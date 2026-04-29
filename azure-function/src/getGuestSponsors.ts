// SPDX-FileCopyrightText: 2026 Workoho GmbH <https://workoho.com>
// SPDX-FileCopyrightText: 2026 Julian Pawlowski <https://github.com/jpawlowski>
// SPDX-License-Identifier: LicenseRef-PolyForm-Shield-1.0.0

import { app, HttpRequest, HttpResponseInit, InvocationContext } from '@azure/functions';
import { createHmac, randomUUID, timingSafeEqual } from 'crypto';
import { DefaultAzureCredential } from '@azure/identity';
import { Client, GraphError, ResponseType } from '@microsoft/microsoft-graph-client';
import { TokenCredentialAuthenticationProvider } from '@microsoft/microsoft-graph-client/authProviders/azureTokenCredentials';
import packageJson from '../package.json';
import { isNewerVersion, latestGitHubVersion, latestGitHubReleaseUrl } from './releaseState.js';

/** Version string exposed in every response via X-Api-Version and used for client/server mismatch detection. */
const FUNCTION_VERSION: string = packageJson.version;

/**
 * Minimum interval between consecutive version-mismatch / update-available warn
 * traces emitted during request handling.  Throttling is per function instance
 * (in-memory) to avoid flooding Application Insights when many guests are active
 * simultaneously.  A 60-minute interval is short enough for a 2-hour KQL alert
 * window to observe at least one trace per evaluation cycle.
 */
const VERSION_WARN_INTERVAL_MS = 3_600_000; // 1 hour
let _lastVersionWarnAt = 0;

/**
 * Throttle map for broken-sponsor-reference warnings.  Keyed by the
 * pseudonymized sponsor OID (redactGuid output); value is the timestamp of the
 * last emitted trace for that sponsor in milliseconds since epoch.
 *
 * One [BROKEN_SPONSOR_REF] warn per sponsor per hour per instance prevents
 * Application Insights flooding when many guests still reference the same
 * deleted account simultaneously.
 */
const BROKEN_SPONSOR_WARN_INTERVAL_MS = 3_600_000; // 1 hour
const _brokenSponsorWarnAt = new Map<string, number>();

/**
 * A guest can have at most 5 sponsors.
 * Keeping the cap explicit avoids unnecessary Graph work and bounds the size
 * of our batch requests.
 */
const MAX_SPONSORS = 5;
const DEFAULT_SPONSOR_LOOKUP_TIMEOUT_MS = 5000;
const DEFAULT_PRESENCE_TIMEOUT_MS = 2500;
const DEFAULT_BATCH_TIMEOUT_MS = 4000;
const DEFAULT_PHOTO_TIMEOUT_MS = 5000;

/**
 * In-memory sliding-window rate limiter, keyed by an arbitrary string.
 *
 * Two tiers:
 *
 * 1. Anonymous / unauthenticated callers (no EasyAuth OID) — always active,
 *    keyed by client IP (`anon:<ip>`).  In production EasyAuth blocks anonymous
 *    requests at the infra level before function code runs; this tier protects
 *    dev environments and acts as a belt-and-suspenders guard should EasyAuth
 *    ever be misconfigured.
 *      5 req / 60 s per IP (hardcoded — anonymous callers are not legitimate).
 *
 * 2. Authenticated callers (valid OID) — disabled by default to avoid false
 *    positives when multiple web parts are on one page or the user reloads
 *    quickly. When enabled, each endpoint gets its own per-user bucket so
 *    sponsor loads, presence polls, and photo fetches do not throttle each other.
 *    Enable only for incident response via:
 *      RATE_LIMIT_ENABLED=true
 *      RATE_LIMIT_MAX_REQUESTS=12   (optional, default 12)
 *      RATE_LIMIT_WINDOW_MS=60000  (optional, default 60 000 ms)
 *
 * Note: on a Consumption plan each instance has its own counter.
 */
const ANON_RATE_LIMIT_MAX_REQUESTS = 5;
const ANON_RATE_LIMIT_WINDOW_MS = 60_000;
const DEFAULT_RATE_LIMIT_ENABLED = false;
const DEFAULT_RATE_LIMIT_MAX_REQUESTS = 12;
const DEFAULT_RATE_LIMIT_WINDOW_MS = 60_000;
const rateLimitMap = new Map<string, number[]>();

/**
 * Presence token — a short-lived HMAC-SHA256-signed string issued by
 * getGuestSponsors and presented to getPresence on every refresh poll.
 *
 * Format:  <base64url(payload)>.<hmac-sha256-hex>
 * Payload: { oid: callerOid, ids: string[], exp: epochMs }
 *
 * The token binds the caller to exactly their set of sponsor IDs, so
 * getPresence can validate requested IDs without keeping server-side state
 * or making an extra Graph call per poll.  This prevents an authenticated
 * tenant member from using the function's Presence.Read.All application
 * permission to probe presence for arbitrary Entra objects.
 *
 * Security properties:
 *   - Tamper-proof:   altering payload or sig invalidates the token.
 *   - Caller-bound:   oid is verified against the EasyAuth identity.
 *   - ID-scoped:      only the exact IDs in the token are accepted.
 *   - Time-limited:   tokens expire after PRESENCE_TOKEN_TTL_MS.
 *   - Unique:         jti (UUID) makes every issued token distinct; aids audit logs.
 *   - Stateless:      works across all Function instances on any hosting plan.
 *
 * When PRESENCE_TOKEN_SECRET is not set, token issuance is skipped and
 * getPresence falls back to a lightweight Graph sponsor lookup for validation.
 */
const PRESENCE_TOKEN_TTL_MS = 2 * 60 * 60 * 1000; // 2 hours

/**
 * Creates a signed presence token for the given caller and sponsor IDs.
 * Returns undefined when PRESENCE_TOKEN_SECRET is not configured.
 */
function createPresenceToken(callerOid: string, sponsorIds: string[], managerIds: string[]): string | undefined {
  const secret = process.env.PRESENCE_TOKEN_SECRET;
  if (!secret) return undefined;
  const payload = JSON.stringify({
    jti: randomUUID(),
    oid: callerOid,
    ids: sponsorIds,
    ...(managerIds.length > 0 ? { managerIds } : {}),
    iat: Date.now(),
    exp: Date.now() + PRESENCE_TOKEN_TTL_MS,
  });
  const payloadB64 = Buffer.from(payload).toString('base64url');
  const sig = createHmac('sha256', secret).update(payloadB64).digest('base64url');
  return `${payloadB64}.${sig}`;
}

/**
 * Verifies a presence token and returns the set of authorised sponsor IDs.
 *
 * Returns undefined when:
 *   - PRESENCE_TOKEN_SECRET is not configured
 *   - The token is missing, malformed, or has an invalid HMAC signature
 *   - The token has expired
 *   - The token's oid does not match the current caller
 *
 * Uses timingSafeEqual to prevent timing side-channel attacks.
 */
function verifyPresenceToken(token: string, callerOid: string): Set<string> | undefined {
  const secret = process.env.PRESENCE_TOKEN_SECRET;
  if (!secret) return undefined;

  const dot = token.lastIndexOf('.');
  if (dot === -1) return undefined;

  const payloadB64 = token.slice(0, dot);
  const providedSig = token.slice(dot + 1);
  const expectedSig = createHmac('sha256', secret).update(payloadB64).digest('base64url');

  // Constant-time HMAC comparison — prevents timing side-channel attacks.
  let sigValid = false;
  try {
    const expected = Buffer.from(expectedSig, 'base64url');
    const provided = Buffer.from(providedSig, 'base64url');
    sigValid = expected.length === provided.length && timingSafeEqual(expected, provided);
  } catch {
    return undefined;
  }
  if (!sigValid) return undefined;

  let payload: { jti?: unknown; oid?: unknown; ids?: unknown; managerIds?: unknown; iat?: unknown; exp?: unknown };
  try {
    payload = JSON.parse(Buffer.from(payloadB64, 'base64url').toString('utf8')) as {
      jti?: unknown; oid?: unknown; ids?: unknown; managerIds?: unknown; iat?: unknown; exp?: unknown;
    };
  } catch {
    return undefined;
  }

  if (typeof payload.exp !== 'number' || payload.exp < Date.now()) return undefined;
  if (payload.oid !== callerOid) return undefined;
  if (!Array.isArray(payload.ids)) return undefined;

  const ids = (payload.ids as unknown[])
    .filter((id): id is string => typeof id === 'string' && isValidGuid(id))
    .slice(0, MAX_SPONSORS);
  // Also include manager IDs in the authorised set — the photo endpoint uses the
  // same token to validate both sponsor and manager photo requests.
  const managerIds = Array.isArray(payload.managerIds)
    ? (payload.managerIds as unknown[])
        .filter((id): id is string => typeof id === 'string' && isValidGuid(id))
        .slice(0, MAX_SPONSORS * 2)
    : [];
  return new Set([...ids, ...managerIds]);
}

/** Cached optional-permission flags for MailboxSettings.Read, Presence.Read.All, and TeamMember.Read.All. */
interface IOptionalPermissions {
  hasMailboxSettings: boolean;
  hasPresenceReadAll: boolean;
  hasTeamMemberReadAll: boolean;
}

let cachedOptionalPermissions: Promise<IOptionalPermissions> | undefined;

/**
 * Inspects the JWT access token obtained by the credential to determine
 * which optional application permissions have been granted.
 *
 * In Azure the DefaultAzureCredential resolves to ManagedIdentityCredential;
 * locally it falls back to AzureCliCredential (or EnvironmentCredential when
 * service-principal env vars are set).  Either way getToken() is cached by the
 * Azure Identity SDK, so this call adds no extra network request.
 *
 * The result is stored as a Promise so that concurrent invocations during
 * a cold start share a single in-flight token inspection instead of each
 * starting their own (eliminates the require-atomic-updates race pattern).
 */
async function detectOptionalPermissions(
  credential: DefaultAzureCredential,
  context: InvocationContext
): Promise<IOptionalPermissions> {
  // Assign synchronously before any await so concurrent callers share one promise.
  if (!cachedOptionalPermissions) {
    cachedOptionalPermissions = (async (): Promise<IOptionalPermissions> => {
      try {
        const token = await credential.getToken('https://graph.microsoft.com/.default');
        // JWT payload is the second segment, base64url-encoded.
        const payloadJson = Buffer.from(token.token.split('.')[1], 'base64url').toString('utf8');
        const payload = JSON.parse(payloadJson) as { roles?: unknown };
        const roles = Array.isArray(payload.roles) ? (payload.roles as string[]) : [];
        const hasUserReadAll = roles.includes('User.Read.All');
        const hasUserReadBasicAll = roles.includes('User.ReadBasic.All');
        const hasPresenceReadAll = roles.includes('Presence.Read.All');
        const hasDirectoryReadAll = roles.includes('Directory.Read.All');
        const hasMailboxSettings = roles.includes('MailboxSettings.Read');
        const hasTeamMemberReadAll = roles.includes('TeamMember.Read.All');
        context.log(
          `Graph app roles: User.Read.All=${hasUserReadAll}, ` +
          `User.ReadBasic.All=${hasUserReadBasicAll}, Presence.Read.All=${hasPresenceReadAll}, ` +
          `Directory.Read.All=${hasDirectoryReadAll}, MailboxSettings.Read=${hasMailboxSettings}, ` +
          `TeamMember.Read.All=${hasTeamMemberReadAll}, count=${roles.length}`
        );
        // The minimum permission required for sponsor lookups is User.ReadBasic.All.
        // If none of the covering permissions are present the very next Graph call will
        // return 403.  Surface this immediately in Application Insights so an admin
        // does not have to hunt for the root cause.
        if (!hasUserReadBasicAll && !hasUserReadAll && !hasDirectoryReadAll) {
          context.error(
            'GRAPH_PERMISSION_DENIED: Managed identity is missing required Microsoft Graph ' +
            'application permissions. Every sponsor lookup will fail with HTTP 403 until fixed.\n' +
            'REQUIRED (at least one): User.ReadBasic.All | User.Read.All | Directory.Read.All\n' +
            'OPTIONAL:               Presence.Read.All | MailboxSettings.Read | TeamMember.Read.All\n' +
            `CURRENTLY GRANTED:      ${roles.length > 0 ? roles.join(', ') : '(none)'}\n` +
            'FIX: Run infra/setup-graph-permissions.ps1 as Global Administrator or ' +
            'Privileged Role Administrator — it grants User.ReadBasic.All and the optional roles.\n' +
            '     Then restart the Function App so it retrieves a fresh token.',
            {
              requiredPermissions: ['User.ReadBasic.All', 'User.Read.All', 'Directory.Read.All'],
              optionalPermissions: ['Presence.Read.All', 'MailboxSettings.Read', 'TeamMember.Read.All'],
              grantedRoles: roles,
              fixAction: 'Run infra/setup-graph-permissions.ps1 and restart the Function App',
            }
          );
        }
        return { hasMailboxSettings, hasPresenceReadAll, hasTeamMemberReadAll };
      } catch (error) {
        // If token inspection fails for any reason, degrade gracefully.
        context.warn('Could not inspect token roles — optional features disabled.', error);
        return { hasMailboxSettings: false, hasPresenceReadAll: false, hasTeamMemberReadAll: false };
      }
    })();
  }
  return cachedOptionalPermissions;
}

/**
 * Extracts the best-effort client IP from X-Forwarded-For.
 * In Azure App Service, the client IP is the first entry added by Azure.
 * Used only for anonymous rate limiting and partial logging — not for trust decisions.
 */
function getClientIp(request: HttpRequest): string {
  const forwarded = request.headers.get('x-forwarded-for');
  if (forwarded) {
    const first = forwarded.split(',')[0].trim();
    // Strip port: IPv4:port → IPv4, [IPv6]:port → IPv6
    return first.replace(/(:\d+)$/, '').replace(/^\[(.+)\]$/, '$1');
  }
  return 'unknown';
}

/**
 * Partially redacts an IP address for logging (GDPR / minimal exposure).
 * IPv4: keeps the /24 prefix, masks the last octet → "192.168.1.x"
 * IPv6: keeps the first four groups, masks the rest → "2001:db8:85a3:0:..."
 */
function redactIp(ip: string): string {
  if (ip === 'unknown') return ip;
  if (/^\d{1,3}(\.\d{1,3}){3}$/.test(ip)) {
    return ip.replace(/\.\d+$/, '.x');
  }
  const parts = ip.split(':');
  if (parts.length >= 4) {
    return `${parts.slice(0, 4).join(':')}:...`;
  }
  return 'x.x.x.x';
}

function isTrue(value: string | undefined): boolean {
  if (!value) return false;
  const v = value.trim().toLowerCase();
  return v === '1' || v === 'true' || v === 'yes' || v === 'on';
}

function getRateLimitConfig(): { enabled: boolean; maxRequests: number; windowMs: number } {
  const enabled = process.env.RATE_LIMIT_ENABLED === undefined
    ? DEFAULT_RATE_LIMIT_ENABLED
    : isTrue(process.env.RATE_LIMIT_ENABLED);

  const maxRaw = Number(process.env.RATE_LIMIT_MAX_REQUESTS);
  const maxRequests = Number.isFinite(maxRaw) && maxRaw > 0
    ? Math.floor(maxRaw)
    : DEFAULT_RATE_LIMIT_MAX_REQUESTS;

  const windowRaw = Number(process.env.RATE_LIMIT_WINDOW_MS);
  const windowMs = Number.isFinite(windowRaw) && windowRaw > 0
    ? Math.floor(windowRaw)
    : DEFAULT_RATE_LIMIT_WINDOW_MS;

  return { enabled, maxRequests, windowMs };
}

type RateLimitEndpoint = 'getGuestSponsors' | 'getPresence' | 'getPhoto';

function getRateLimitKey(endpoint: RateLimitEndpoint, subjectKey: string): string {
  return `${endpoint}:${subjectKey}`;
}

/**
 * Memory bounds for the in-process rate-limiter map.
 *
 * SOFT_CAP — entry count that triggers an expired-entry sweep.
 * HARD_CAP — absolute ceiling: when the map is still above this limit after the
 *            expired-entry sweep, the oldest-accessed (LRU) entries are evicted
 *            down to SOFT_CAP.  This prevents memory exhaustion under a DDoS
 *            with many distinct source IPs / user IDs.
 *
 *   At ~300 bytes/entry (key + timestamp array + Map overhead),
 *   10 000 entries ≈ 3 MB — well within the Azure Functions consumption-plan
 *   per-instance memory budget.
 *
 * MIN_GC_INTERVAL_MS — rate-limits the O(n) sweep so it runs at most once per
 *   second even during a flood, capping CPU overhead under attack.
 * PERIODIC_GC_MS     — also sweeps every minute when the map is below SOFT_CAP,
 *   so expired entries from quiet periods don't accumulate indefinitely.
 */
const RATE_LIMIT_SOFT_CAP = 5_000;
const RATE_LIMIT_HARD_CAP = 10_000;
const RATE_LIMIT_MIN_GC_INTERVAL_MS = 1_000;
const RATE_LIMIT_PERIODIC_GC_MS = 60_000;

let rlLastGcAt = 0;

/**
 * Two-phase GC for the rate-limiter map:
 *   Phase 1 — remove entries whose every recorded timestamp has expired.
 *   Phase 2 — if still above HARD_CAP, evict the oldest-accessed entries
 *             (LRU order is maintained by delete-before-set on every access).
 */
function rateLimitGc(windowStart: number): void {
  for (const [key, ts] of rateLimitMap) {
    if (ts.every(t => t <= windowStart)) rateLimitMap.delete(key);
  }
  if (rateLimitMap.size > RATE_LIMIT_HARD_CAP) {
    let toEvict = rateLimitMap.size - RATE_LIMIT_SOFT_CAP;
    for (const key of rateLimitMap.keys()) {
      rateLimitMap.delete(key);
      if (--toEvict <= 0) break;
    }
  }
}

function checkRateLimit(userId: string, maxRequests: number, windowMs: number): { allowed: boolean; retryAfterSeconds?: number } {
  const now = Date.now();
  const windowStart = now - windowMs;

  // GC: sweep when the map is large, or periodically — but at most once per
  // second so the O(n) pass never becomes a hot path under a DDoS flood.
  if (
    (rateLimitMap.size > RATE_LIMIT_SOFT_CAP || now - rlLastGcAt >= RATE_LIMIT_PERIODIC_GC_MS) &&
    now - rlLastGcAt >= RATE_LIMIT_MIN_GC_INTERVAL_MS
  ) {
    rlLastGcAt = now;
    rateLimitGc(windowStart);
  }

  const recent = (rateLimitMap.get(userId) ?? []).filter(t => t > windowStart);
  if (recent.length >= maxRequests) {
    const oldest = Math.min(...recent);
    const retryAfterSeconds = Math.ceil((oldest + windowMs - now) / 1000);
    // delete-before-set keeps LRU insertion order current for the eviction pass.
    rateLimitMap.delete(userId);
    rateLimitMap.set(userId, recent);
    return { allowed: false, retryAfterSeconds };
  }

  recent.push(now);
  // delete-before-set keeps LRU insertion order current for the eviction pass.
  rateLimitMap.delete(userId);
  rateLimitMap.set(userId, recent);
  return { allowed: true };
}

interface IBatchRequest {
  id: string;
  method: 'GET';
  url: string;
}

interface IBatchResponseItem {
  id: string;
  status: number;
  body?: Record<string, unknown>;
}

class TimeoutError extends Error {
  public constructor(operation: string, timeoutMs: number) {
    super(`${operation} timed out after ${timeoutMs}ms`);
    this.name = 'TimeoutError';
  }
}

/**
 * Converts an ArrayBuffer (binary photo data from Graph) into a base64-encoded
 * data URL suitable for embedding directly in a JSON response body.
 */
function arrayBufferToDataUrl(buffer: ArrayBuffer, mimeType = 'image/jpeg'): string {
  return `data:${mimeType};base64,${Buffer.from(buffer).toString('base64')}`;
}

/**
 * Validates that a string is a well-formed Entra / Azure AD object ID (GUID).
 * Used as a defence-in-depth guard before embedding any ID in a Graph API URL,
 * even though EasyAuth and Microsoft Graph already validate their own inputs.
 */
function isValidGuid(value: string): boolean {
  return /^[0-9a-f]{8}-[0-9a-f]{4}-[0-9a-f]{4}-[0-9a-f]{4}-[0-9a-f]{12}$/i.test(value);
}

function getTimeoutMs(envVarName: string, defaultValue: number): number {
  const value = process.env[envVarName];
  if (!value) return defaultValue;

  const parsed = Number(value);
  return Number.isFinite(parsed) && parsed > 0 ? parsed : defaultValue;
}

function redactGuid(value: string): string {
  return `${value.slice(0, 8)}...${value.slice(-4)}`;
}

async function withTimeout<T>(
  promise: Promise<T>,
  timeoutMs: number,
  operation: string
): Promise<T> {
  let timer: ReturnType<typeof setTimeout> | undefined;
  const timeoutPromise = new Promise<never>((_resolve, reject) => {
    timer = setTimeout(() => reject(new TimeoutError(operation, timeoutMs)), timeoutMs);
  });
  try {
    return await Promise.race([promise, timeoutPromise]);
  } finally {
    clearTimeout(timer);
  }
}

/** Shape of a sponsor returned by this function (matches ISponsor in SPFx). */
interface ISponsor {
  id: string;
  displayName: string;
  givenName?: string;
  surname?: string;
  mail?: string;
  jobTitle?: string;
  department?: string;
  officeLocation?: string;
  streetAddress?: string;
  postalCode?: string;
  state?: string;
  city?: string;
  country?: string;
  businessPhones: string[];
  mobilePhone?: string;
  presence?: string;
  presenceActivity?: string;
  managerDisplayName?: string;
  managerGivenName?: string;
  managerSurname?: string;
  managerJobTitle?: string;
  managerDepartment?: string;
  /** Manager's Entra ID — used by the SPFx client to fetch the manager photo progressively. */
  managerId?: string;
  /** Base64-encoded data URL of the sponsor's profile photo. Populated by the function. */
  photoUrl?: string;
  /** True when the sponsor has an active Microsoft Teams license. */
  hasTeams?: boolean;
}

/**
 * Shape of the JSON response (matches ISponsorsResult in SPFx).
 * unavailableCount includes sponsors that are deleted, soft-deleted, or disabled.
 * guestHasTeamsAccess is undefined when neither TeamMember.Read.All nor Presence.Read.All
 * is granted — the client falls back to showing buttons enabled (fail-open).
 */
interface ISponsorsResult {
  activeSponsors: ISponsor[];
  unavailableCount: number;
  guestHasTeamsAccess?: boolean;
  /**
   * Short-lived HMAC-signed token that authorizes subsequent getPresence calls
   * for exactly this caller and these sponsor IDs.  Present only when
   * PRESENCE_TOKEN_SECRET is configured in the Function App settings.
   */
  presenceToken?: string;
  /**
   * Sponsor profiles that exist (or existed) in the directory but whose
   * account is disabled, soft-deleted, or otherwise unavailable. Populated
   * when all sponsors fall into this category so the client can render
   * read-only tiles alongside the "sponsor not available" notice.
   */
  unavailableSponsors?: ISponsor[];
  /**
   * Entra object IDs of all sponsors in the original Graph response order.
   * Used by the client to walk the list in priority order and let active
   * accounts "nachrücken" when higher-priority sponsors are unavailable,
   * while still displaying unavailable accounts alongside the active ones.
   */
  sponsorOrder?: string[];
}

/**
 * Logs a structured rejection event with redacted identifiers.
 * Includes context about why the request was rejected without exposing sensitive data.
 */
function logRejection(
  context: InvocationContext,
  reasonCode: string,
  reason: string,
  details: Record<string, unknown> = {}
): void {
  context.warn(`Client validation (${reasonCode}): ${reason}`, {
    reasonCode,
    ...details,
    timestamp: new Date().toISOString(),
  });
}

/**
 * EasyAuth client-principal claim entry.
 */
interface IEasyAuthClaim {
  typ: string;
  val: string;
}

/**
 * EasyAuth client-principal structure.
 */
interface IEasyAuthPrincipal {
  auth_typ?: string;
  claims?: IEasyAuthClaim[];
  name_typ?: string;
  role_typ?: string;
  userId?: string;
  userDetails?: string;
  identityProvider?: string;
}

/**
 * Parses the EasyAuth principal from X-MS-CLIENT-PRINCIPAL.
 * Header is Base64-encoded JSON emitted only after successful EasyAuth validation.
 */
function parseEasyAuthPrincipal(request: HttpRequest): IEasyAuthPrincipal | undefined {
  const encodedPrincipal = request.headers.get('x-ms-client-principal');
  if (!encodedPrincipal) return undefined;

  try {
    const json = Buffer.from(encodedPrincipal, 'base64').toString('utf8');
    const parsed = JSON.parse(json) as IEasyAuthPrincipal;
    return parsed;
  } catch {
    return undefined;
  }
}

/**
 * Returns the first claim value that matches any claim type.
 * Supports both short names (tid, aud) and URI claim types used by Entra/EasyAuth.
 */
function getPrincipalClaim(principal: IEasyAuthPrincipal, claimTypes: string[]): string | undefined {
  if (!Array.isArray(principal.claims)) return undefined;
  const normalized = claimTypes.map(t => t.toLowerCase());

  for (const claim of principal.claims) {
    const type = (claim.typ ?? '').toLowerCase();
    if (normalized.includes(type) || normalized.some(t => type.endsWith(`/${t}`))) {
      return claim.val;
    }
  }

  return undefined;
}

/**
 * Validates that the authenticated EasyAuth principal belongs to our tenant and API audience,
 * and that the request originated from a specific Entra application.
 *
 * Checks performed on EasyAuth principal claims:
 * 1. tid claim matches our tenant ID
 * 2. aud claim matches our API's audience URI
 * 3. appid claim matches the well-known SPFx "SharePoint Online Web Client Extensibility" App ID
 *
 * Check 3 ensures that only a specific Entra application (e.g. the "SharePoint
 * Online Web Client Extensibility" app used by SPFx AadHttpClient) can call
 * this API. Without it, any Entra client that has been granted the API scope
 * could make requests.
 *
 * @returns validation result with details
 */
function validateClientAuthorization(
  principal: IEasyAuthPrincipal,
  context: InvocationContext
): { authorized: boolean; reasonCode?: string; tid?: string; reason?: string } {
  if (!principal || !Array.isArray(principal.claims)) {
    return { authorized: false, reasonCode: 'AUTH_PRINCIPAL_MISSING', reason: 'EasyAuth principal missing or invalid' };
  }

  // Check 1: Tenant ID must match our environment.
  const tenantId = process.env.TENANT_ID;
  const tid = getPrincipalClaim(principal, ['tid', 'tenantid', 'http://schemas.microsoft.com/identity/claims/tenantid']);
  if (!tid) {
    return { authorized: false, reasonCode: 'AUTH_CLAIM_MISSING_TID', reason: 'Principal missing tid (tenant ID) claim' };
  }
  if (!tenantId) {
    context.error('TENANT_ID environment variable not configured');
    return { authorized: false, reasonCode: 'AUTH_CONFIG_TENANT_MISSING', reason: 'Server configuration error: TENANT_ID not set' };
  }
  if (tid !== tenantId) {
    return {
      authorized: false,
      reasonCode: 'AUTH_TENANT_MISMATCH',
      tid,
      reason: 'Principal tenant does not match TENANT_ID',
    };
  }

  // Check 2: Audience must match our client ID.
  // With accessTokenAcceptedVersion=2 the aud claim is always the bare
  // client ID (GUID).  ALLOWED_AUDIENCE must be set to the same GUID.
  const allowedAudience = process.env.ALLOWED_AUDIENCE;
  if (!allowedAudience) {
    context.error('ALLOWED_AUDIENCE environment variable not configured');
    return { authorized: false, reasonCode: 'AUTH_CONFIG_AUDIENCE_MISSING', reason: 'Server configuration error: ALLOWED_AUDIENCE not set' };
  }
  const aud = getPrincipalClaim(principal, ['aud']);
  if (!aud) {
    return {
      authorized: false,
      reasonCode: 'AUTH_CLAIM_MISSING_AUD',
      reason: 'Principal missing aud (audience) claim',
    };
  }

  if (aud !== allowedAudience) {
    return {
      authorized: false,
      reasonCode: 'AUTH_AUDIENCE_MISMATCH',
      reason: 'Principal audience does not match ALLOWED_AUDIENCE',
    };
  }

  // Check 3: Calling application must match the expected client app.
  // The appid (v1 tokens) or azp (v2 tokens) claim identifies the Entra
  // application that acquired the token. For SPFx web parts the calling app is
  // always the Microsoft-managed "SharePoint Online Web Client Extensibility"
  // multi-tenant app — its App ID is globally fixed and identical in every
  // Entra tenant.  No env-var is needed; the constant is the authoritative value.
  // Source: https://www.eliostruyf.com/fix-admin-consent-sp-token-retrieval-flows-spfx/
  const allowedClientAppId = '08e18876-6177-487e-b8b5-cf950c1e598c';
  const appid = getPrincipalClaim(principal, [
    'appid',
    'azp',
    'http://schemas.microsoft.com/identity/claims/appid',
  ]);
  if (!appid) {
    return {
      authorized: false,
      reasonCode: 'AUTH_CLAIM_MISSING_APPID',
      reason: 'Principal missing appid/azp (calling application) claim',
    };
  }
  if (appid !== allowedClientAppId) {
    return {
      authorized: false,
      reasonCode: 'AUTH_APPID_MISMATCH',
      reason: 'Calling application is not the authorized SharePoint Web Client Extensibility app',

    };
  }

  return { authorized: true, tid };
}

/**
 * Resolves the OID of the calling user from EasyAuth headers.
 *
 * In production, Azure App Service EasyAuth sets these headers after validating the
 * Bearer token.  The function code never sees an unvalidated token.
 *
 * For local development (EasyAuth absent), the OID can be passed via
 * the X-Dev-User-OID header — only accepted when NODE_ENV !== 'production'.
 */
function resolveCallerOid(request: HttpRequest): string | undefined {
  // EasyAuth sets X-MS-CLIENT-PRINCIPAL-ID to the validated caller OID.
  const easyAuthOid = request.headers.get('x-ms-client-principal-id');
  if (easyAuthOid) return isValidGuid(easyAuthOid) ? easyAuthOid : undefined;

  // Local dev fallback — never accepted in production.
  if (process.env.NODE_ENV !== 'production') {
    const devOid = request.headers.get('x-dev-user-oid');
    if (devOid) return isValidGuid(devOid) ? devOid : undefined;
  }

  return undefined;
}

/**
 * Fetches presence for a list of user IDs in a single batched Graph call.
 * Returns a map of userId → availability/activity strings.
 * Silently returns an empty map on any error or timeout (presence is non-critical).
 */
async function fetchPresences(
  client: Client,
  userIds: string[],
  context: InvocationContext
): Promise<Map<string, { availability?: string; activity?: string }>> {
  const map = new Map<string, { availability?: string; activity?: string }>();
  if (userIds.length === 0) return map;
  // Build a Set of the requested IDs so we can reject any unexpected entries
  // the Graph API might return (defence-in-depth against unexpected response data).
  const requestedIds = new Set(userIds);
  try {
    const response = await withTimeout(
      client
        .api('/communications/getPresencesByUserId')
        .post({ ids: userIds }),
      getTimeoutMs('PRESENCE_TIMEOUT_MS', DEFAULT_PRESENCE_TIMEOUT_MS),
      'Graph presence lookup'
    );
    if (response?.value) {
      for (const entry of response.value as Array<{ id: string; availability?: string; activity?: string }>) {
        // Only accept entries whose ID was in our request and whose availability
        // is a plain alphabetic string (no injection vectors).
        const availability = entry.availability && /^[A-Za-z]+$/.test(entry.availability)
          ? entry.availability
          : undefined;
        const activity = entry.activity && /^[A-Za-z]+$/.test(entry.activity)
          ? entry.activity
          : undefined;
        if (
          entry.id &&
          requestedIds.has(entry.id) &&
          (availability || activity)
        ) {
          map.set(entry.id, { availability, activity });
        }
      }
    }
  } catch (error) {
    // Presence is supplemental — silently degrade when the call fails.
    context.warn('Presence lookup degraded.', error);
  }
  return map;
}

/**
 * Executes a single Microsoft Graph $batch request and returns a map keyed by
 * the request id for efficient lookup by the caller.
 */
async function executeBatch(
  client: Client,
  requests: IBatchRequest[],
  timeoutMs: number,
  operation: string
): Promise<Map<string, IBatchResponseItem>> {
  const results = new Map<string, IBatchResponseItem>();
  if (requests.length === 0) return results;

  const response = await withTimeout(
    client
      .api('/$batch')
      .post({ requests }),
    timeoutMs,
    operation
  );

  for (const item of (response?.responses ?? []) as IBatchResponseItem[]) {
    if (item?.id) {
      results.set(item.id, item);
    }
  }

  return results;
}

function getStringValue(source: Record<string, unknown>, key: string): string | undefined {
  const value = source[key];
  return typeof value === 'string' && value.length > 0 ? value : undefined;
}

function getBooleanValue(source: Record<string, unknown>, key: string): boolean | undefined {
  const value = source[key];
  return typeof value === 'boolean' ? value : undefined;
}

/**
 * Returns true when a plan entry has an active-or-grace-period capability status.
 * Both 'Enabled' (fully active) and 'Warning' (grace period — plan will be removed
 * but the service is still accessible) count as "available" for sponsor eligibility.
 */
function isPlanActive(capabilityStatus: unknown): boolean {
  return capabilityStatus === 'Enabled' || capabilityStatus === 'Warning';
}

/**
 * Returns true when the assignedPlans array contains at least one active plan of any type.
 * Used for the 'any' sponsor filter mode.
 */
function assignedPlansHaveAny(plans: unknown): boolean {
  if (!Array.isArray(plans)) return false;
  return (plans as Array<{ capabilityStatus?: unknown }>)
    .some(p => isPlanActive(p.capabilityStatus));
}

/**
 * Returns true when the assignedPlans array contains at least one active Exchange Online plan.
 * Exchange plans are identified by service === 'exchange'.
 * Both 'Enabled' and 'Warning' capabilityStatus count as active.
 */
function assignedPlansHaveExchange(plans: unknown): boolean {
  if (!Array.isArray(plans)) return false;
  return (plans as Array<{ service?: unknown; capabilityStatus?: unknown }>)
    .some(p => p.service === 'exchange' && isPlanActive(p.capabilityStatus));
}

/**
 * Returns true when the assignedPlans array contains at least one active Teams plan.
 * Teams plans are identified by service === 'TeamspaceAPI'.
 * Both 'Enabled' and 'Warning' capabilityStatus count as active.
 */
function assignedPlansHaveTeams(plans: unknown): boolean {
  if (!Array.isArray(plans)) return false;
  return (plans as Array<{ service?: unknown; capabilityStatus?: unknown }>)
    .some(p => p.service === 'TeamspaceAPI' && isPlanActive(p.capabilityStatus));
}

/**
 * Allowed values for the `sponsorFilter` query parameter.
 * Validated server-side before use — the raw query string value is never passed
 * directly to Graph or used in a filter expression.
 */
const VALID_SPONSOR_FILTERS = ['any', 'exchange', 'teams'] as const;
type SponsorFilter = typeof VALID_SPONSOR_FILTERS[number];

/**
 * Validates and normalizes HTTP status codes. Returns a valid HTTP status in the
 * range 200–599; defaults to 500 if the provided code is undefined or out of range.
 * This guards against GraphError or other exceptions providing invalid status codes
 * that would crash the Azure Functions runtime.
 */
function getValidHttpStatus(code: unknown): number {
  if (typeof code === 'number' && code >= 200 && code <= 599) return code;
  return 500;
}

function getCorrelationId(request: HttpRequest): string {
  return request.headers.get('x-correlation-id')
    ?? request.headers.get('x-ms-request-id')
    ?? request.headers.get('x-arr-log-id')
    ?? randomUUID();
}

function jsonErrorResponse(
  request: HttpRequest,
  status: number,
  correlationId: string,
  error: string,
  reasonCode: string,
  message: string,
  retryable: boolean
): HttpResponseInit {
  return {
    status,
    body: JSON.stringify({ error, reasonCode, message, retryable, referenceId: correlationId }),
    headers: {
      'Content-Type': 'application/json',
      'x-correlation-id': correlationId,
      'x-api-version': FUNCTION_VERSION,
      ...corsHeaders(request),
    },
  };
}

/**
 * HTTP GET - returns the sponsors of the calling user.
 *
 * Authentication is handled by EasyAuth (Azure App Service Authentication).
 * The function reads the caller OID from the X-MS-CLIENT-PRINCIPAL-ID header
 * that EasyAuth sets after validating the Bearer token.
 *
 * DefaultAzureCredential is used to call Microsoft Graph with application
 * permissions.  In Azure it resolves to ManagedIdentityCredential; locally it
 * falls back to Azure CLI or environment-variable credentials.
 * No client secrets are stored anywhere.
 */
export async function getGuestSponsors(
  request: HttpRequest,
  context: InvocationContext
): Promise<HttpResponseInit> {
  const correlationId = getCorrelationId(request);

  // CORS preflight.
  if (request.method === 'OPTIONS') {
    return {
      status: 204,
      headers: { 'x-correlation-id': correlationId, ...corsHeaders(request) },
    };
  }

  // Method guard: only GET is supported. Reject early with 405 to avoid exposing
  // internal logic to method-probing or accidental non-GET calls.
  if (request.method !== 'GET') {
    return {
      status: 405,
      body: JSON.stringify({ error: 'Method Not Allowed' }),
      headers: {
        'Content-Type': 'application/json',
        Allow: 'GET, OPTIONS',
        'x-correlation-id': correlationId,
        ...corsHeaders(request),
      },
    };
  }

  const callerOid = resolveCallerOid(request);
  if (!callerOid) {
    // Anonymous / unauthenticated caller — apply strict IP-based rate limiting
    // before rejecting, to prevent hammering in dev or on EasyAuth misconfiguration.
    const clientIp = getClientIp(request);
    const anonLimit = checkRateLimit(
      getRateLimitKey('getGuestSponsors', `anon:${clientIp}`),
      ANON_RATE_LIMIT_MAX_REQUESTS,
      ANON_RATE_LIMIT_WINDOW_MS
    );
    if (!anonLimit.allowed) {
      context.warn('Anonymous caller rate-limited', {
        reasonCode: 'AUTH_RATE_LIMITED',
        clientIp: redactIp(clientIp),
        correlationId,
      });
      const resp = jsonErrorResponse(
        request,
        429,
        correlationId,
        'Too Many Requests',
        'AUTH_RATE_LIMITED',
        'Request rate limit exceeded. Retry after the provided delay.',
        true
      );
      resp.headers = {
        ...(resp.headers ?? {}),
        'Retry-After': String(anonLimit.retryAfterSeconds ?? 60),
      };
      return resp;
    }
    context.warn('Anonymous caller rejected — EasyAuth not configured or bypassed', {
      reasonCode: 'AUTH_CALLER_OID_MISSING',
      clientIp: redactIp(clientIp),
      correlationId,
    });
    return jsonErrorResponse(
      request,
      401,
      correlationId,
      'Unauthorized',
      'AUTH_CALLER_OID_MISSING',
      'Authenticated caller could not be resolved from EasyAuth headers',
      false
    );
  }

  // Validate the EasyAuth principal belongs to our tenant and API audience.
  // Skip in development mode unless NODE_ENV is explicitly 'production'.
  if (process.env.NODE_ENV === 'production') {
    const principal = parseEasyAuthPrincipal(request);
    if (!principal) {
      logRejection(context, 'AUTH_PRINCIPAL_MISSING', 'No valid EasyAuth principal header present', {
        step: 'client-validation-start',
        correlationId,
      });
      return jsonErrorResponse(
        request,
        401,
        correlationId,
        'Unauthorized',
        'AUTH_PRINCIPAL_MISSING',
        'EasyAuth principal header is missing or invalid',
        false
      );
    }

    const validation = validateClientAuthorization(principal, context);
    if (!validation.authorized) {
      const reasonCode = validation.reasonCode ?? 'AUTH_VALIDATION_FAILED';
      logRejection(context, reasonCode, validation.reason ?? 'Unknown authorization failure', {
        step: 'client-validation-failed',
        tid: validation.tid,
        callerOid: redactGuid(callerOid),
        correlationId,
      });
      return jsonErrorResponse(
        request,
        403,
        correlationId,
        'Forbidden',
        reasonCode,
        'Request token failed tenant or audience validation',
        false
      );
    }

    context.log(`Client authorization validated for tenant ${validation.tid}`);
  }

  const rateLimitConfig = getRateLimitConfig();
  if (rateLimitConfig.enabled) {
    const rateLimitResult = checkRateLimit(
      getRateLimitKey('getGuestSponsors', callerOid),
      rateLimitConfig.maxRequests,
      rateLimitConfig.windowMs
    );
    if (!rateLimitResult.allowed) {
      const response = jsonErrorResponse(
        request,
        429,
        correlationId,
        'Too Many Requests',
        'AUTH_RATE_LIMITED',
        'Request rate limit exceeded. Retry after the provided delay.',
        true
      );
      response.headers = {
        ...(response.headers ?? {}),
        'Retry-After': String(rateLimitResult.retryAfterSeconds ?? 60),
      };
      return response;
    }
  }

  // Version compatibility check.
  //
  // When the web part sends X-Client-Version and it does not match the running
  // function version, log a structured warning so Azure Monitor can alert.
  // Throttled per function instance (1 h) to avoid flooding Application Insights
  // when many guests are online simultaneously.
  //
  // Hierarchy of logged sentinel tokens (used by KQL alert rules):
  //   [VERSION_MISMATCH]          — versions differ; no GitHub context yet
  //   [WEBPART_UPDATE_AVAILABLE]  — S3: function is on latest, web part is behind
  //   [FUNCTION_UPDATE_AVAILABLE] — S4: web part is on latest, function is behind
  //
  // The [NEW_RELEASE_AVAILABLE] token (emitted by checkGitHubRelease.ts) covers
  // the case where both components are at the same version but GitHub has a newer
  // release.
  const clientVersion = request.headers.get('x-client-version');
  if (clientVersion && clientVersion !== FUNCTION_VERSION) {
    const now = Date.now();
    if (now - _lastVersionWarnAt >= VERSION_WARN_INTERVAL_MS) {
      _lastVersionWarnAt = now;
      const cached = latestGitHubVersion; // snapshot; may be undefined on first cold start
      const funcIsNewer = isNewerVersion(FUNCTION_VERSION, clientVersion);
      const clientIsNewer = isNewerVersion(clientVersion, FUNCTION_VERSION);

      if (funcIsNewer && cached && !isNewerVersion(cached, FUNCTION_VERSION)) {
        // S3: function is on (or ahead of) the latest GitHub release; web part is behind.
        context.warn(
          `[WEBPART_UPDATE_AVAILABLE] webPartVersion=${clientVersion}` +
          ` functionVersion=${FUNCTION_VERSION} latestVersion=${cached}`
        );
      } else if (clientIsNewer && cached && !isNewerVersion(cached, clientVersion)) {
        // S4: web part is on (or ahead of) the latest GitHub release; function is behind.
        context.warn(
          `[FUNCTION_UPDATE_AVAILABLE] functionVersion=${FUNCTION_VERSION}` +
          ` webPartVersion=${clientVersion} latestVersion=${cached}`
        );
      } else {
        // Generic mismatch — no GitHub context available yet, or both are behind latest.
        const olderComponent = funcIsNewer ? 'webpart' : 'function';
        context.warn(
          `[VERSION_MISMATCH] functionVersion=${FUNCTION_VERSION}` +
          ` webPartVersion=${clientVersion} olderComponent=${olderComponent}` +
          (cached ? ` latestVersion=${cached}` : '')
        );
      }
    }
  }

  // Mock mode — return realistic demo data without Graph credentials.
  // Only accepted outside production to prevent accidental use.
  if (process.env.NODE_ENV !== 'production' && isTrue(process.env.MOCK_MODE)) {
    context.log('MOCK_MODE active — returning demo sponsor data');
    const mockResult: ISponsorsResult = {
      activeSponsors: [
        {
          id: '00000000-0000-0000-0000-000000000001',
          displayName: 'Anna Müller',
          givenName: 'Anna',
          surname: 'Müller',
          mail: 'anna.mueller@contoso.com',
          jobTitle: 'IT Manager',
          department: 'Information Technology',
          officeLocation: 'BER-HQ / Bldg A / Floor 4 / A4-12',
          city: 'Berlin',
          country: 'Germany',
          businessPhones: ['+49 30 12345678'],
          presence: 'Available',
          presenceActivity: 'Available',
          managerDisplayName: 'Thomas Schneider',
          managerGivenName: 'Thomas',
          managerSurname: 'Schneider',
          managerJobTitle: 'Head of IT',
          managerId: '00000000-0000-0000-0000-000000000003',
          hasTeams: true,
        },
        {
          id: '00000000-0000-0000-0000-000000000002',
          displayName: 'James Anderson',
          givenName: 'James',
          surname: 'Anderson',
          mail: 'james.anderson@contoso.com',
          jobTitle: 'Project Lead',
          department: 'Business Development',
          officeLocation: 'MUC-03 / Bldg C / Floor 2 / C2-08',
          city: 'Munich',
          country: 'Germany',
          businessPhones: [],
          mobilePhone: '+49 151 98765432',
          presence: 'Busy',
          presenceActivity: 'InACall',
          managerDisplayName: 'Sarah Webb',
          managerGivenName: 'Sarah',
          managerSurname: 'Webb',
          managerJobTitle: 'VP Business Development',
          managerId: '00000000-0000-0000-0000-000000000004',
          hasTeams: true,
        },
      ],
      unavailableCount: 0,
      guestHasTeamsAccess: true,
    };
    return jsonResponse(mockResult, 200, request, correlationId);
  }

  // Parse and validate sponsor eligibility filter parameters sent by the web part.
  // Both values are strictly validated against known-good constants before use —
  // the raw query string is never forwarded to Graph or embedded in a filter string.
  const rawFilter = request.query.get('sponsorFilter') ?? 'teams';
  const sponsorFilter: SponsorFilter = (VALID_SPONSOR_FILTERS as readonly string[]).includes(rawFilter)
    ? rawFilter as SponsorFilter
    : 'teams';
  // requireUserMailbox: true unless the web part explicitly sends 'false'.
  const requireUserMailbox = request.query.get('requireUserMailbox') !== 'false';

  context.log(`Fetching sponsors for caller ${redactGuid(callerOid)}`
    + ` [sponsorFilter=${sponsorFilter}, requireUserMailbox=${requireUserMailbox}]`);

  try {
    const credential = new DefaultAzureCredential();
    const { hasMailboxSettings, hasPresenceReadAll, hasTeamMemberReadAll } = await detectOptionalPermissions(credential, context);
    const authProvider = new TokenCredentialAuthenticationProvider(credential, {
      scopes: ['https://graph.microsoft.com/.default'],
    });
    // Let the Graph SDK create its default middleware pipeline from authProvider
    // (includes auth + default handlers). Passing middleware[0] was invalid and
    // caused runtime failures: "Cannot read properties of undefined (reading 'execute')".
    const client = Client.initWithMiddleware({ authProvider });
    let response: Record<string, unknown> | undefined;
    try {
      response = await withTimeout(
        client
          .api(`/users/${callerOid}/sponsors`)
          .select('id,displayName,givenName,surname,mail,jobTitle,department,officeLocation,streetAddress,postalCode,state,city,country,businessPhones,mobilePhone')
          // Enforce the cap at the Graph level to avoid fetching more items than we will process.
          .top(MAX_SPONSORS)
          .get(),
        getTimeoutMs('SPONSOR_LOOKUP_TIMEOUT_MS', DEFAULT_SPONSOR_LOOKUP_TIMEOUT_MS),
        'Graph sponsor lookup'
      ) as Record<string, unknown>;
    } catch (error) {
      if (error instanceof GraphError) {
        context.error('Graph sponsor lookup failed:', {
          statusCode: error.statusCode,
          code: error.code,
          requestId: error.requestId,
          message: error.message,
          callerOid: redactGuid(callerOid),
        });
      }
      throw error;
    }

    if (!response?.value) {
      const result: ISponsorsResult = { activeSponsors: [], unavailableCount: 0 };
      return jsonResponse(result, 200, request, correlationId);
    }

    const items = response.value as Record<string, unknown>[];
    // Explicitly map only the fields we requested via $select and validate each
    // sponsor ID is a GUID before making any further per-sponsor Graph calls.
    const candidates: ISponsor[] = items
      .filter(item => typeof item.id === 'string' && isValidGuid(item.id))
      .slice(0, MAX_SPONSORS)
      .map(item => ({
        id: item.id as string,
        displayName: (item.displayName as string) || '',
        givenName: (item.givenName as string) || undefined,
        surname: (item.surname as string) || undefined,
        mail: (item.mail as string) || undefined,
        jobTitle: (item.jobTitle as string) || undefined,
        department: (item.department as string) || undefined,
        officeLocation: (item.officeLocation as string) || undefined,
        streetAddress: (item.streetAddress as string) || undefined,
        postalCode: (item.postalCode as string) || undefined,
        state: (item.state as string) || undefined,
        city: (item.city as string) || undefined,
        country: (item.country as string) || undefined,
        businessPhones: Array.isArray(item.businessPhones)
          ? (item.businessPhones as unknown[]).filter((p): p is string => typeof p === 'string')
          : [],
        mobilePhone: (item.mobilePhone as string) || undefined,
      }));

    const sponsorIds = candidates.map(s => s.id);

    // Each sponsor requires only a single batch sub-request because the manager
    // object is inlined via $expand instead of fetched through a second request.
    // This halves the number of batch operations compared to the previous approach
    // (N requests instead of 2N) and stays well within the Graph $batch limit of 20.
    const sponsorBatchRequests: IBatchRequest[] = candidates.map((sponsor, index) => ({
      id: `exists-${index}`,
      method: 'GET',
      // Include mailboxSettings only when the permission is granted AND the admin
      // configured "require user mailbox type". Without that setting, the mailbox
      // type check is skipped entirely, making the extra field wasted bandwidth.
      url: `/users/${sponsor.id}?$select=id,accountEnabled,isResourceAccount,assignedPlans${hasMailboxSettings && requireUserMailbox ? ',mailboxSettings' : ''}&$expand=manager($select=id,displayName,givenName,surname,jobTitle,department,accountEnabled)`,
    }));

    // When TeamMember.Read.All is granted, add a joinedTeams sub-request to the
    // same batch — saves an extra HTTP round-trip compared to a standalone call.
    if (hasTeamMemberReadAll) {
      sponsorBatchRequests.push({
        id: 'joinedTeams',
        method: 'GET',
        url: `/users/${callerOid}/joinedTeams?$select=id&$top=1`,
      });
    }

    const [presenceMap, batchResults, rawPhotoResults] = await Promise.all([
      // Include callerOid so we can use the guest's own presence as a provisioning signal.
      hasPresenceReadAll ? fetchPresences(client, [...sponsorIds, callerOid], context) : Promise.resolve(new Map<string, { availability?: string; activity?: string }>()),
      executeBatch(
        client,
        sponsorBatchRequests,
        getTimeoutMs('BATCH_TIMEOUT_MS', DEFAULT_BATCH_TIMEOUT_MS),
        'Graph sponsor detail batch'
      ),
      // Fetch profile photos for all candidate sponsors in parallel with the batch.
      // Each photo is fetched independently; failures silently fall back to initials.
      Promise.allSettled(
        candidates.map(s =>
          withTimeout(
            (client.api(`/users/${s.id}/photo/$value`).responseType(ResponseType.ARRAYBUFFER).get() as Promise<ArrayBuffer>)
              .then(buf => ({ id: s.id, dataUrl: arrayBufferToDataUrl(buf) }))
              .catch(() => ({ id: s.id, dataUrl: undefined })),
            getTimeoutMs('PHOTO_TIMEOUT_MS', DEFAULT_PHOTO_TIMEOUT_MS),
            `photo for ${redactGuid(s.id)}`
          ).catch(() => ({ id: s.id, dataUrl: undefined }))
        )
      ),
    ]);

    // Build a photo lookup map to attach photoUrl to each active sponsor below.
    const photoMap = new Map<string, string>();
    for (const r of rawPhotoResults) {
      if (r.status === 'fulfilled' && r.value.dataUrl !== undefined) {
        photoMap.set(r.value.id, r.value.dataUrl);
      }
    }

    // Extract joinedTeams count from the batch response (if included).
    let guestJoinedTeamsCount: number | undefined;
    if (hasTeamMemberReadAll) {
      const jtResponse = batchResults.get('joinedTeams');
      if (jtResponse?.status === 200 && jtResponse.body) {
        const jtValue = (jtResponse.body as Record<string, unknown>).value;
        guestJoinedTeamsCount = Array.isArray(jtValue) ? jtValue.length : 0;
      }
      // else: undefined — error or missing, falls through to presence fallback
    }

    const perSponsorResults = candidates.map((sponsor, index) => {
      const existsResponse = batchResults.get(`exists-${index}`);
      const existsBody = existsResponse?.body;
      const sponsorEnabled = existsBody !== undefined
        ? getBooleanValue(existsBody, 'accountEnabled')
        : undefined;
      // isResourceAccount flags Teams Room devices, Common Area Phones, and other
      // resource accounts — these are never valid sponsors regardless of their licenses.
      // Fail-open when the property is absent (undefined) so a transient Graph response
      // does not silently hide a legitimately configured sponsor.
      const isResourceAccount = existsBody !== undefined
        ? getBooleanValue(existsBody, 'isResourceAccount') === true
        : false;

      const assignedPlans = existsBody?.assignedPlans;

      // License eligibility: controlled by the sponsorFilter query parameter.
      //   'any'      — at least one plan with Enabled or Warning capabilityStatus
      //   'exchange' — at least one Exchange Online plan (service='exchange') active
      //   'teams'    — at least one Teams plan (service='TeamspaceAPI') active (default)
      const meetsLicenseFilter =
        sponsorFilter === 'any'      ? assignedPlansHaveAny(assignedPlans) :
        sponsorFilter === 'exchange' ? assignedPlansHaveExchange(assignedPlans) :
                                       assignedPlansHaveTeams(assignedPlans);

      // Mailbox eligibility: controlled by the requireUserMailbox query parameter.
      //
      // requireUserMailbox=true (default):
      //   Check mailboxSettings.userPurpose when MailboxSettings.Read is granted.
      //   Accepted values: 'user' (cloud mailbox) and 'linked' (on-premises mailbox).
      //   Any other value (shared, room, equipment, …) marks this as non-eligible.
      //   Fail-open when MailboxSettings.Read is NOT granted so a missing permission
      //   never silently hides a legitimately configured sponsor.
      //
      // requireUserMailbox=false:
      //   Skip the mailbox-type check entirely. Use an active Exchange Online license
      //   as a proxy for "has some mailbox" — works without MailboxSettings.Read.
      let hasUserMailbox = true;
      if (requireUserMailbox) {
        if (hasMailboxSettings && existsBody !== undefined) {
          const mb = existsBody.mailboxSettings;
          if (mb !== null && mb !== undefined) {
            const userPurpose = typeof (mb as Record<string, unknown>).userPurpose === 'string'
              ? (mb as Record<string, unknown>).userPurpose as string
              : undefined;
            if (userPurpose !== undefined) {
              hasUserMailbox = userPurpose === 'user' || userPurpose === 'linked';
            }
          }
        }
        // else hasMailboxSettings=false → fail-open (hasUserMailbox stays true)
      } else {
        // "Any mailbox" mode: Exchange Online license is the proxy for mailbox existence.
        hasUserMailbox = assignedPlansHaveExchange(assignedPlans);
      }

      // Distinguish between three outcomes:
      //
      //   exists = true          — sponsor is active and meets all filters.
      //                            Shown as a full interactive card.
      //
      //   exists = false         — sponsor account exists but is disabled or is a
      //     resource account (Teams Room, Common Area Phone, …).  Shown as a
      //     read-only tigeminile so the guest can still see who their sponsors are.
      //
      //   filterMismatch = true  — sponsor is completely excluded from the response:
      //     • Hard-deleted (Graph 404): no display name available; nothing useful
      //       to show the guest.  The [BROKEN_SPONSOR_REF] warning already alerts
      //       admins via Application Insights.
      //     • License/mailbox filter not met (e.g. no Teams plan when
      //       sponsorFilter=teams): sponsor intentionally out-of-scope.
      const is404 = existsResponse?.status === 404;
      const filterMismatch = is404 || !meetsLicenseFilter || !hasUserMailbox;
      const exists = !filterMismatch && sponsorEnabled !== false && !isResourceAccount;

      // Emit a throttled warn for hard-deleted sponsors (Graph 404).  A persistent
      // 404 indicates a broken sponsor reference — the Entra object no longer exists
      // and an admin should remove it from the guest's sponsor list.  One trace per
      // sponsor per hour per instance avoids flooding Application Insights.
      if (existsResponse?.status === 404) {
        const key = redactGuid(sponsor.id);
        const now = Date.now();
        if ((now - (_brokenSponsorWarnAt.get(key) ?? 0)) >= BROKEN_SPONSOR_WARN_INTERVAL_MS) {
          _brokenSponsorWarnAt.set(key, now);
          context.warn(`[BROKEN_SPONSOR_REF] sponsorId=${key}`);
        }
      }

      const hasTeams = assignedPlansHaveTeams(existsBody?.assignedPlans);

      // Manager is inlined via $expand — read directly from the user body.
      const managerBody = existsBody !== undefined && existsBody.manager !== null
        ? existsBody.manager as Record<string, unknown> | undefined
        : undefined;

      let managerDisplayName: string | undefined;
      let managerGivenName: string | undefined;
      let managerSurname: string | undefined;
      let managerJobTitle: string | undefined;
      let managerDepartment: string | undefined;
      let managerId: string | undefined;

      if (managerBody !== undefined) {
        const managerEnabled = getBooleanValue(managerBody, 'accountEnabled');
        if (managerEnabled !== false) {
          managerDisplayName = getStringValue(managerBody, 'displayName');
          managerGivenName = getStringValue(managerBody, 'givenName');
          managerSurname = getStringValue(managerBody, 'surname');
          managerJobTitle = getStringValue(managerBody, 'jobTitle');
          managerDepartment = getStringValue(managerBody, 'department');
          const mid = getStringValue(managerBody, 'id');
          if (mid && isValidGuid(mid)) managerId = mid;
        }
      }

      return {
        sponsor: {
          ...sponsor,
          managerDisplayName,
          managerGivenName,
          managerSurname,
          managerJobTitle,
          managerDepartment,
          managerId,
          hasTeams,
        },
        exists,
        filterMismatch,
      };
    });

    const activeSponsors = perSponsorResults
      .filter(r => r.exists)
      // Explicitly enumerate all output fields so no unexpected properties from
      // intermediate objects can leak into the response payload.
      .map(r => {
        const s = r.sponsor;
        const out: ISponsor = {
          id: s.id,
          displayName: s.displayName,
          businessPhones: s.businessPhones,
        };
        if (s.mail !== undefined)                out.mail = s.mail;
        if (s.givenName !== undefined)           out.givenName = s.givenName;
        if (s.surname !== undefined)             out.surname = s.surname;
        if (s.jobTitle !== undefined)            out.jobTitle = s.jobTitle;
        if (s.department !== undefined)          out.department = s.department;
        if (s.officeLocation !== undefined)      out.officeLocation = s.officeLocation;
        if (s.streetAddress !== undefined)       out.streetAddress = s.streetAddress;
        if (s.postalCode !== undefined)          out.postalCode = s.postalCode;
        if (s.state !== undefined)               out.state = s.state;
        if (s.city !== undefined)               out.city = s.city;
        if (s.country !== undefined)            out.country = s.country;
        if (s.mobilePhone !== undefined)         out.mobilePhone = s.mobilePhone;
        if (s.managerDisplayName !== undefined)  out.managerDisplayName = s.managerDisplayName;
        if (s.managerGivenName !== undefined)    out.managerGivenName = s.managerGivenName;
        if (s.managerSurname !== undefined)      out.managerSurname = s.managerSurname;
        if (s.managerJobTitle !== undefined)     out.managerJobTitle = s.managerJobTitle;
        if (s.managerDepartment !== undefined)   out.managerDepartment = s.managerDepartment;
        if (s.managerId !== undefined)           out.managerId = s.managerId;
        out.hasTeams = s.hasTeams;  // always a boolean from the proxy; undefined only in direct path
        const presence = presenceMap.get(s.id);
        if (presence?.availability !== undefined) out.presence = presence.availability;
        if (presence?.activity !== undefined)     out.presenceActivity = presence.activity;
        const photo = photoMap.get(s.id);
        if (photo !== undefined)                  out.photoUrl = photo;
        return out;
      });
    // Only include account-level unavailable sponsors (disabled / deleted / resource).
    // Sponsors excluded by the license or mailbox filter are omitted entirely.
    const unavailableCount = perSponsorResults.filter(r => !r.exists && !r.filterMismatch).length;
    const unavailableSponsors: ISponsor[] = perSponsorResults
      .filter(r => !r.exists && !r.filterMismatch)
      .map(r => {
        const s = r.sponsor;
        const out: ISponsor = {
          id: s.id,
          displayName: s.displayName,
          businessPhones: s.businessPhones ?? [],
        };
        if (s.givenName)   out.givenName   = s.givenName;
        if (s.surname)     out.surname     = s.surname;
        if (s.jobTitle)    out.jobTitle    = s.jobTitle;
        if (s.department)  out.department  = s.department;
        return out;
      });

    // Determine whether the guest's Teams service account has been provisioned.
    //
    // Signal hierarchy (strongest → weakest):
    //   1. joinedTeams non-empty  → definitively provisioned (account + resource exist)
    //   2. joinedTeams empty      → account MAY still exist if the guest was removed from
    //                              their last Team; use presence as tie-breaker:
    //        • real presence (not PresenceUnknown) → account exists, no current membership
    //        • PresenceUnknown / absent             → account not yet provisioned
    //   3. Neither permission granted               → unknown; client fails open
    let guestHasTeamsAccess: boolean | undefined;
    const guestPresenceSnapshot = presenceMap.get(callerOid);
    const guestPresenceKnown =
      guestPresenceSnapshot !== undefined &&
      guestPresenceSnapshot.availability !== undefined &&
      guestPresenceSnapshot.availability !== 'PresenceUnknown';

    if (hasTeamMemberReadAll) {
      if (guestJoinedTeamsCount === undefined) {
        // Error fetching joinedTeams — fall back to presence if available, else unknown.
        guestHasTeamsAccess = hasPresenceReadAll ? guestPresenceKnown : undefined;
      } else if (guestJoinedTeamsCount > 0) {
        guestHasTeamsAccess = true;
      } else {
        // No joined teams — presence distinguishes "removed from last team" from "never provisioned".
        guestHasTeamsAccess = hasPresenceReadAll ? guestPresenceKnown : false;
      }
    } else if (hasPresenceReadAll) {
      // TeamMember.Read.All not granted — presence alone as signal.
      guestHasTeamsAccess = guestPresenceKnown ? true : false;
    }
    // else → undefined: neither permission granted, client falls open.

    const result: ISponsorsResult = { activeSponsors, unavailableCount };
    result.sponsorOrder = sponsorIds;
    if (unavailableSponsors.length > 0) result.unavailableSponsors = unavailableSponsors;
    if (guestHasTeamsAccess !== undefined) result.guestHasTeamsAccess = guestHasTeamsAccess;

    // Issue a signed presence token so subsequent getPresence and getPhoto polls can be
    // validated without server-side state or extra Graph calls.
    // Include manager IDs so the photo endpoint can authorise manager photo requests
    // using the same token without an additional Graph lookup.
    const managerIds = [...new Set(
      activeSponsors
        .filter(s => s.managerId !== undefined)
        .map(s => s.managerId as string)
    )];
    const presenceToken = createPresenceToken(callerOid, sponsorIds, managerIds);
    if (presenceToken !== undefined) result.presenceToken = presenceToken;

    return jsonResponse(result, 200, request, correlationId);
  } catch (error) {
    if (error instanceof GraphError) {
      if (error.statusCode === 403) {
        // HTTP 403 from Graph means the managed identity was never granted the required
        // application permissions.  Log the exact missing permission so the ops team
        // knows precisely what to fix without reading Microsoft's generic error messages.
        context.error(
          'GRAPH_PERMISSION_DENIED: Microsoft Graph returned HTTP 403 Forbidden.\n' +
          'CAUSE: The managed identity has not been granted the Microsoft Graph application\n' +
          '       permission "User.ReadBasic.All" (minimum required for sponsor lookups).\n' +
          'REQUIRED (at least one of):\n' +
          '  • User.ReadBasic.All  — read basic profile of any user (displayName, mail, photo)\n' +
          '  • User.Read.All       — read full profile of any user\n' +
          '  • Directory.Read.All  — read all directory data\n' +
          'OPTIONAL (for additional features):\n' +
          '  • Presence.Read.All   — real-time Teams presence indicators\n' +
          '  • MailboxSettings.Read — filter shared/room mailboxes out of sponsor list\n' +
          '  • TeamMember.Read.All — detect Teams provisioning for guests\n' +
          'FIX: Run infra/setup-graph-permissions.ps1 as Global Administrator or\n' +
          '     Privileged Role Administrator. The script grants all roles above.\n' +
          '     Then restart the Function App to pick up the new token.',
          {
            statusCode: error.statusCode,
            code: error.code,
            requestId: error.requestId,
            correlationId,
            requiredPermissions: ['User.ReadBasic.All', 'User.Read.All', 'Directory.Read.All'],
            fixAction: 'Run infra/setup-graph-permissions.ps1 and restart the Function App',
          }
        );
        return jsonErrorResponse(
          request,
          403,
          correlationId,
          'Forbidden',
          'GRAPH_PERMISSION_DENIED',
          'The function managed identity is missing required Microsoft Graph permissions. ' +
          'Contact your administrator to run setup-graph-permissions.ps1.',
          false
        );
      }
      // GraphError is thrown by the Graph SDK for HTTP-level failures.
      // requestId is the Graph correlation ID — log it so it can be looked up
      // in Azure Monitor or provided to Microsoft Support for debugging.
      context.error('Graph error fetching sponsors:', {
        statusCode: error.statusCode,
        code: error.code,
        requestId: error.requestId,
        message: error.message,
      });
    } else if (error instanceof TimeoutError) {
      context.warn(`[SPONSOR_LOOKUP_TIMEOUT] timeoutMs=${getTimeoutMs('SPONSOR_LOOKUP_TIMEOUT_MS', DEFAULT_SPONSOR_LOOKUP_TIMEOUT_MS)}`);
    } else {
      context.error('Error fetching sponsors:', error);
    }
    const status = error instanceof TimeoutError
      ? 504
      : getValidHttpStatus(
        error instanceof GraphError ? error.statusCode : (error as { statusCode?: number }).statusCode
      );
    const retryable = status >= 500 || status === 429;
    return jsonErrorResponse(
      request,
      status,
      correlationId,
      'Failed to retrieve sponsor information.',
      status === 504 ? 'SPONSOR_LOOKUP_TIMEOUT' : 'SPONSOR_LOOKUP_FAILED',
      'Sponsor retrieval failed in backend processing',
      retryable
    );
  }
}

function jsonResponse(body: unknown, status: number, request: HttpRequest, correlationId: string): HttpResponseInit {
  return {
    status,
    body: JSON.stringify(body),
    headers: {
      'Content-Type': 'application/json',
      'x-correlation-id': correlationId,
      'x-api-version': FUNCTION_VERSION,
      ...corsHeaders(request),
    },
  };
}

/**
 * Returns CORS headers allowing the request's origin when it matches the configured origin.
 * The allowed origin is configured via the CORS_ALLOWED_ORIGIN environment variable
 * (e.g. "https://contoso.sharepoint.com"). When the env var is not set, no
 * Access-Control-Allow-Origin header is emitted — the Azure Function-level CORS
 * settings (if any) will handle enforcement instead.
 *
 * Also applies security hardening headers to prevent content type sniffing,
 * clickjacking, browser cache issues, and enforce HTTPS transmission.
 */
function corsHeaders(request: HttpRequest): Record<string, string> {
  const origin = request.headers.get('origin') ?? '';
  const allowedOrigin = process.env.CORS_ALLOWED_ORIGIN ?? '';

  const headers: Record<string, string> = {
    // CORS
    'Access-Control-Allow-Methods': 'GET, OPTIONS',
    'Access-Control-Allow-Headers': 'Authorization, Content-Type, X-Presence-Token',
    'Access-Control-Expose-Headers': 'x-correlation-id',
    // Minimal security headers for API responses.
    // HSTS, X-Frame-Options, and other browser-level protections are better handled
    // by SharePoint Online or Azure Application Gateway, not at the API level.
    'X-Content-Type-Options': 'nosniff',           // Prevent content-type sniffing attacks
  };

  if (allowedOrigin && origin === allowedOrigin) {
    headers['Access-Control-Allow-Origin'] = origin;
  }

  return headers;
}

/**
 * Final guard at the function boundary: normalize response status codes and
 * ensure unexpected exceptions still return a valid HTTP response.
 */
async function safeGetGuestSponsors(
  request: HttpRequest,
  context: InvocationContext
): Promise<HttpResponseInit> {
  try {
    const response = await getGuestSponsors(request, context);
    return {
      ...response,
      status: getValidHttpStatus(response.status),
    };
  } catch (error) {
    const correlationId = getCorrelationId(request);
    context.error('Unhandled exception in safeGetGuestSponsors:', error);
    return jsonErrorResponse(
      request,
      500,
      correlationId,
      'Failed to retrieve sponsor information.',
      'UNHANDLED_EXCEPTION',
      'Unexpected backend exception occurred',
      true
    );
  }
}

app.http('getGuestSponsors', {
  methods: ['GET', 'OPTIONS'],
  authLevel: 'anonymous', // Authentication is enforced by EasyAuth, not the function key.
  handler: safeGetGuestSponsors,
});

/**
 * HTTP GET - returns Microsoft Teams presence for a list of Entra object IDs.
 *
 * Designed to serve presence refresh polls from the SPFx web part after the
 * initial sponsor load.  Using application permissions (Managed Identity) keeps
 * presence up-to-date reliably for guest callers — unlike the delegated
 * Presence.Read.All scope, which may silently return empty results for guests
 * on tenants with restrictive guest-access policies.
 *
 * Query parameter:
 *   ids  - comma-separated list of up to MAX_SPONSORS Entra object IDs (GUIDs).
 *
 * If Presence.Read.All has not been granted to the Managed Identity the endpoint
 * returns HTTP 200 with an empty presences array and logs a warning.  The caller
 * should preserve any previously loaded presence data on screen.
 */
export async function getPresence(
  request: HttpRequest,
  context: InvocationContext
): Promise<HttpResponseInit> {
  const correlationId = getCorrelationId(request);

  // CORS preflight.
  if (request.method === 'OPTIONS') {
    return {
      status: 204,
      headers: { 'x-correlation-id': correlationId, ...corsHeaders(request) },
    };
  }

  if (request.method !== 'GET') {
    return {
      status: 405,
      body: JSON.stringify({ error: 'Method Not Allowed' }),
      headers: {
        'Content-Type': 'application/json',
        Allow: 'GET, OPTIONS',
        'x-correlation-id': correlationId,
        ...corsHeaders(request),
      },
    };
  }

  // Parse and validate the 'ids' query parameter.
  const rawIds = request.query.get('ids') ?? '';
  const ids = rawIds
    .split(',')
    .map(s => s.trim())
    .filter(s => isValidGuid(s))
    .slice(0, MAX_SPONSORS); // cap to prevent abusive large requests

  if (ids.length === 0) {
    return jsonErrorResponse(
      request,
      400,
      correlationId,
      'Bad Request',
      'INVALID_IDS',
      'ids must be a comma-separated list of up to 5 valid Entra object ID GUIDs',
      false
    );
  }

  const callerOid = resolveCallerOid(request);
  if (!callerOid) {
    const clientIp = getClientIp(request);
    const anonLimit = checkRateLimit(
      getRateLimitKey('getPresence', `anon:${clientIp}`),
      ANON_RATE_LIMIT_MAX_REQUESTS,
      ANON_RATE_LIMIT_WINDOW_MS
    );
    if (!anonLimit.allowed) {
      const resp = jsonErrorResponse(
        request,
        429,
        correlationId,
        'Too Many Requests',
        'AUTH_RATE_LIMITED',
        'Request rate limit exceeded. Retry after the provided delay.',
        true
      );
      resp.headers = {
        ...(resp.headers ?? {}),
        'Retry-After': String(anonLimit.retryAfterSeconds ?? 60),
      };
      return resp;
    }
    context.warn('Anonymous caller rejected — EasyAuth not configured or bypassed', {
      reasonCode: 'AUTH_CALLER_OID_MISSING',
      clientIp: redactIp(clientIp),
      correlationId,
    });
    return jsonErrorResponse(
      request,
      401,
      correlationId,
      'Unauthorized',
      'AUTH_CALLER_OID_MISSING',
      'Authenticated caller could not be resolved from EasyAuth headers',
      false
    );
  }

  if (process.env.NODE_ENV === 'production') {
    const principal = parseEasyAuthPrincipal(request);
    if (!principal) {
      logRejection(context, 'AUTH_PRINCIPAL_MISSING', 'No valid EasyAuth principal header present', {
        step: 'client-validation-start',
        correlationId,
      });
      return jsonErrorResponse(
        request,
        401,
        correlationId,
        'Unauthorized',
        'AUTH_PRINCIPAL_MISSING',
        'EasyAuth principal header is missing or invalid',
        false
      );
    }

    const validation = validateClientAuthorization(principal, context);
    if (!validation.authorized) {
      const reasonCode = validation.reasonCode ?? 'AUTH_VALIDATION_FAILED';
      logRejection(context, reasonCode, validation.reason ?? 'Unknown authorization failure', {
        step: 'client-validation-failed',
        tid: validation.tid,
        callerOid: redactGuid(callerOid),
        correlationId,
      });
      return jsonErrorResponse(
        request,
        403,
        correlationId,
        'Forbidden',
        reasonCode,
        'Request token failed tenant or audience validation',
        false
      );
    }
  }

  const rateLimitConfig = getRateLimitConfig();
  if (rateLimitConfig.enabled) {
    const rateLimitResult = checkRateLimit(
      getRateLimitKey('getPresence', callerOid),
      rateLimitConfig.maxRequests,
      rateLimitConfig.windowMs
    );
    if (!rateLimitResult.allowed) {
      const response = jsonErrorResponse(
        request,
        429,
        correlationId,
        'Too Many Requests',
        'AUTH_RATE_LIMITED',
        'Request rate limit exceeded. Retry after the provided delay.',
        true
      );
      response.headers = {
        ...(response.headers ?? {}),
        'Retry-After': String(rateLimitResult.retryAfterSeconds ?? 60),
      };
      return response;
    }
  }

  // Mock mode — return demo presence without Graph credentials.
  if (process.env.NODE_ENV !== 'production' && isTrue(process.env.MOCK_MODE)) {
    context.log('MOCK_MODE active — returning demo presence data');
    const mockAvailabilities = ['Available', 'Busy', 'Away', 'DoNotDisturb', 'BeRightBack'];
    const presences = ids.map((id, i) => ({
      id,
      availability: mockAvailabilities[i % mockAvailabilities.length],
      activity: mockAvailabilities[i % mockAvailabilities.length],
    }));
    return jsonResponse({ presences }, 200, request, correlationId);
  }

  context.log(`Fetching presence for ${ids.length} id(s) on behalf of ${redactGuid(callerOid)}`);

  try {
    const credential = new DefaultAzureCredential();
    const { hasPresenceReadAll } = await detectOptionalPermissions(credential, context);

    if (!hasPresenceReadAll) {
      context.warn(
        'getPresence: Presence.Read.All not granted — ' +
        'returning empty presences array. ' +
        'Run infra/setup-graph-permissions.ps1 to enable presence polling via the proxy.'
      );
      return jsonResponse({ presences: [] }, 200, request, correlationId);
    }

    const authProvider = new TokenCredentialAuthenticationProvider(credential, {
      scopes: ['https://graph.microsoft.com/.default'],
    });
    const client = Client.initWithMiddleware({ authProvider });

    // Validate which IDs this caller is allowed to query presence for.
    //
    // Fast path (preferred): verify the HMAC-signed token issued by getGuestSponsors.
    //   - Stateless, zero extra Graph calls, works across all Function instances.
    //   - Sent as X-Presence-Token header (not URL query param) to keep it out of
    //     server logs and Application Insights URL traces.
    //
    // Fallback: when no valid token is provided (old client, first poll after a
    //   Function restart with a new secret, or PRESENCE_TOKEN_SECRET not configured),
    //   fetch the caller's sponsor list from Graph to build the allowed-ID set.
    //   Logs a warning when a token was supplied but failed verification.
    //
    // Either way, any requested ID not in the authorised set is silently dropped
    // so the response never leaks presence for non-sponsors.
    const rawToken = request.headers.get('x-presence-token') ?? '';
    let authorizedIds: Set<string> | undefined;

    if (rawToken) {
      authorizedIds = verifyPresenceToken(rawToken, callerOid);
      if (!authorizedIds) {
        context.warn('getPresence: presence token invalid or expired — falling back to Graph validation', {
          callerOid: redactGuid(callerOid),
          correlationId,
          hasSecret: !!process.env.PRESENCE_TOKEN_SECRET,
        });
      }
    }

    if (!authorizedIds) {
      // Token absent or invalid — fetch sponsor IDs directly from Graph.
      try {
        const sponsorResponse = await withTimeout(
          client
            .api(`/users/${callerOid}/sponsors`)
            .select('id')
            .get(),
          getTimeoutMs('SPONSOR_LOOKUP_TIMEOUT_MS', DEFAULT_SPONSOR_LOOKUP_TIMEOUT_MS),
          'sponsor ID lookup for presence validation'
        );
        const graphIds = ((sponsorResponse?.value ?? []) as Array<{ id?: unknown }>)
          .map(u => u.id)
          .filter((id): id is string => typeof id === 'string' && isValidGuid(id))
          .slice(0, MAX_SPONSORS);
        authorizedIds = new Set(graphIds);
      } catch (sponsorError) {
        // Fail closed: if we cannot confirm the caller's sponsor set, return no
        // presence data for this poll rather than authorizing arbitrary IDs.
        context.warn('getPresence: sponsor validation lookup failed — returning empty presences', sponsorError);
        return jsonResponse({ presences: [] }, 200, request, correlationId);
      }
    }

    // Filter to only the IDs the caller is authorized to query.
    const validatedIds = ids.filter(id => (authorizedIds as Set<string>).has(id));
    if (validatedIds.length < ids.length) {
      context.warn('getPresence: dropping unauthorized IDs from presence request', {
        requested: ids.length,
        authorized: validatedIds.length,
        callerOid: redactGuid(callerOid),
        correlationId,
      });
    }

    if (validatedIds.length === 0) {
      return jsonResponse({ presences: [] }, 200, request, correlationId);
    }

    const presenceMap = await fetchPresences(client, validatedIds, context);

    const presences = validatedIds.map(id => {
      const p = presenceMap.get(id);
      const entry: { id: string; availability?: string; activity?: string } = { id };
      if (p?.availability !== undefined) entry.availability = p.availability;
      if (p?.activity !== undefined) entry.activity = p.activity;
      return entry;
    });

    return jsonResponse({ presences }, 200, request, correlationId);
  } catch (error) {
    if (error instanceof GraphError && error.statusCode === 403) {
      context.error(
        'GRAPH_PERMISSION_DENIED: Microsoft Graph returned HTTP 403 for presence lookup.\n' +
        'CAUSE: Presence.Read.All application permission not granted to the managed identity.\n' +
        'FIX: Run infra/setup-graph-permissions.ps1 and restart the Function App.'
      );
      return jsonResponse({ presences: [] }, 200, request, correlationId);
    }
    context.error('Presence lookup failed:', error);
    return jsonErrorResponse(
      request,
      502,
      correlationId,
      'Presence lookup failed',
      'PRESENCE_LOOKUP_FAILED',
      'Failed to retrieve presence data from Microsoft Graph',
      true
    );
  }
}

async function safeGetPresence(
  request: HttpRequest,
  context: InvocationContext
): Promise<HttpResponseInit> {
  try {
    const response = await getPresence(request, context);
    return {
      ...response,
      status: getValidHttpStatus(response.status),
    };
  } catch (error) {
    const correlationId = getCorrelationId(request);
    context.error('Unhandled exception in safeGetPresence:', error);
    return jsonErrorResponse(
      request,
      500,
      correlationId,
      'Failed to retrieve presence information.',
      'UNHANDLED_EXCEPTION',
      'Unexpected backend exception occurred',
      true
    );
  }
}

app.http('getPresence', {
  methods: ['GET', 'OPTIONS'],
  authLevel: 'anonymous', // Authentication is enforced by EasyAuth, not the function key.
  handler: safeGetPresence,
});

/**
 * HTTP GET - lightweight health-check endpoint.
 *
 * Returns HTTP 200 with the function version.  No authentication checks,
 * no Graph calls — EasyAuth still gates access at the infrastructure level,
 * but the function itself performs no caller validation.  This lets the SPFx
 * web part verify connectivity to the Function App in edit mode without
 * triggering sponsor lookups or permission errors for non-guest editors.
 */
app.http('ping', {
  methods: ['GET', 'OPTIONS'],
  authLevel: 'anonymous',
  handler: async (request: HttpRequest): Promise<HttpResponseInit> => {
    if (request.method === 'OPTIONS') {
      return { status: 204, headers: corsHeaders(request) };
    }
    return {
      status: 200,
      body: JSON.stringify({ status: 'ok', version: FUNCTION_VERSION }),
      headers: {
        'Content-Type': 'application/json',
        'x-api-version': FUNCTION_VERSION,
        ...corsHeaders(request),
      },
    };
  },
});

/**
 * HTTP GET - returns the latest published GitHub release version cached in memory.
 *
 * The timer trigger in `checkGitHubRelease` fetches the GitHub Releases API once
 * every six hours (and on every cold start) and stores the result in module-level
 * in-memory state (`releaseState.ts`).  This endpoint exposes that cached state so
 * the SPFx web part can learn about new releases without ever calling GitHub
 * directly from the browser.  Multiple simultaneous clients therefore share a
 * single outbound GitHub request per six-hour window.
 *
 * Response body (JSON):
 *   latestVersion  - semver string of the latest GitHub release (no leading "v"),
 *                    or null when the timer has not yet completed its first run.
 *   url            - HTML URL of the GitHub release page, or null.
 *   functionVersion - the currently deployed function version.
 *
 * No caller-identity or Graph work is performed here.  EasyAuth gates access
 * at the infrastructure level (same as /api/ping).
 */
/**
 * HTTP GET - returns the profile photo of a sponsor or their manager, proxied via
 * the function's Managed Identity (User.Read.All).
 *
 * This allows the SPFx web part to lazy-load manager photos without making direct
 * Microsoft Graph calls. All data fetching is centralised in the Azure Function
 * so the web part needs no Graph permissions of its own.
 *
 * Query parameter:
 *   userId - Entra object ID of the user whose photo to return.
 *            Must be a sponsor or manager of one of the caller's sponsors
 *            (validated via the presence token or a live Graph lookup).
 *
 * Header (optional but recommended):
 *   X-Presence-Token - the signed token issued by getGuestSponsors.
 *                      When present, enables stateless zero-Graph-call validation.
 *
 * Returns:
 *   200 { photoUrl: "data:image/jpeg;base64,…" }
 *   404 when Graph returns 404 (no photo set for this user)
 */
export async function getPhoto(
  request: HttpRequest,
  context: InvocationContext
): Promise<HttpResponseInit> {
  const correlationId = getCorrelationId(request);

  if (request.method === 'OPTIONS') {
    return { status: 204, headers: { 'x-correlation-id': correlationId, ...corsHeaders(request) } };
  }

  if (request.method !== 'GET') {
    return {
      status: 405,
      body: JSON.stringify({ error: 'Method Not Allowed' }),
      headers: {
        'Content-Type': 'application/json',
        Allow: 'GET, OPTIONS',
        'x-correlation-id': correlationId,
        ...corsHeaders(request),
      },
    };
  }

  const userId = request.query.get('userId') ?? '';
  if (!isValidGuid(userId)) {
    return jsonErrorResponse(
      request, 400, correlationId,
      'Bad Request', 'INVALID_USER_ID',
      'userId must be a valid Entra object ID (GUID)', false
    );
  }

  const callerOid = resolveCallerOid(request);
  if (!callerOid) {
    const clientIp = getClientIp(request);
    const anonLimit = checkRateLimit(
      getRateLimitKey('getPhoto', `anon:${clientIp}`),
      ANON_RATE_LIMIT_MAX_REQUESTS,
      ANON_RATE_LIMIT_WINDOW_MS
    );
    if (!anonLimit.allowed) {
      const resp = jsonErrorResponse(
        request, 429, correlationId,
        'Too Many Requests', 'AUTH_RATE_LIMITED',
        'Request rate limit exceeded. Retry after the provided delay.', true
      );
      resp.headers = { ...(resp.headers ?? {}), 'Retry-After': String(anonLimit.retryAfterSeconds ?? 60) };
      return resp;
    }
    context.warn('getPhoto: anonymous caller rejected', { correlationId });
    return jsonErrorResponse(
      request, 401, correlationId,
      'Unauthorized', 'AUTH_CALLER_OID_MISSING',
      'Authenticated caller could not be resolved from EasyAuth headers', false
    );
  }

  if (process.env.NODE_ENV === 'production') {
    const principal = parseEasyAuthPrincipal(request);
    if (!principal) {
      return jsonErrorResponse(
        request, 401, correlationId,
        'Unauthorized', 'AUTH_PRINCIPAL_MISSING',
        'EasyAuth principal header is missing or invalid', false
      );
    }
    const validation = validateClientAuthorization(principal, context);
    if (!validation.authorized) {
      const reasonCode = validation.reasonCode ?? 'AUTH_VALIDATION_FAILED';
      logRejection(context, reasonCode, validation.reason ?? 'Authorization failure', { correlationId });
      return jsonErrorResponse(
        request, 403, correlationId,
        'Forbidden', reasonCode,
        'Request token failed tenant validation', false
      );
    }
  }

  const rateLimitConfig = getRateLimitConfig();
  if (rateLimitConfig.enabled) {
    const rateLimitResult = checkRateLimit(
      getRateLimitKey('getPhoto', callerOid),
      rateLimitConfig.maxRequests,
      rateLimitConfig.windowMs
    );
    if (!rateLimitResult.allowed) {
      const resp = jsonErrorResponse(
        request, 429, correlationId,
        'Too Many Requests', 'AUTH_RATE_LIMITED',
        'Request rate limit exceeded. Retry after the provided delay.', true
      );
      resp.headers = { ...(resp.headers ?? {}), 'Retry-After': String(rateLimitResult.retryAfterSeconds ?? 60) };
      return resp;
    }
  }

  // Validate that the caller is authorised to fetch this user's photo.
  // Fast path: verify the HMAC-signed token issued by getGuestSponsors
  // (which now includes both sponsor IDs and manager IDs).
  // Fallback: re-fetch sponsor list + manager IDs from Graph.
  const rawToken = request.headers.get('x-presence-token') ?? '';
  let authorizedIds: Set<string> | undefined;

  if (rawToken) {
    authorizedIds = verifyPresenceToken(rawToken, callerOid);
    if (!authorizedIds) {
      context.warn('getPhoto: presence token invalid or expired — falling back to Graph validation', {
        callerOid: redactGuid(callerOid),
        correlationId,
        hasSecret: !!process.env.PRESENCE_TOKEN_SECRET,
      });
    }
  }

  const credential = new DefaultAzureCredential();
  const authProvider = new TokenCredentialAuthenticationProvider(credential, {
    scopes: ['https://graph.microsoft.com/.default'],
  });
  const client = Client.initWithMiddleware({ authProvider });

  if (!authorizedIds) {
    // Fallback: build the authorised ID set from Graph.
    try {
      const sponsorResponse = await withTimeout(
        client.api(`/users/${callerOid}/sponsors`).select('id').top(MAX_SPONSORS).get(),
        getTimeoutMs('SPONSOR_LOOKUP_TIMEOUT_MS', DEFAULT_SPONSOR_LOOKUP_TIMEOUT_MS),
        'sponsor ID lookup for photo validation'
      );
      const sponsorIds = ((sponsorResponse?.value ?? []) as Array<{ id?: unknown }>)
        .map(u => u.id)
        .filter((id): id is string => typeof id === 'string' && isValidGuid(id))
        .slice(0, MAX_SPONSORS);

      // Also include manager IDs so callers can load manager photos without a token.
      const expandedIds = [...sponsorIds];
      if (sponsorIds.length > 0) {
        const batchReqs: IBatchRequest[] = sponsorIds.map((id, i) => ({
          id: `m-${i}`, method: 'GET',
          url: `/users/${id}?$select=id&$expand=manager($select=id)`,
        }));
        const batchRes = await executeBatch(
          client, batchReqs,
          getTimeoutMs('BATCH_TIMEOUT_MS', DEFAULT_BATCH_TIMEOUT_MS),
          'manager ID batch for photo validation'
        );
        for (let i = 0; i < sponsorIds.length; i++) {
          const item = batchRes.get(`m-${i}`);
          const mgr = item?.body?.manager as Record<string, unknown> | null | undefined;
          if (mgr && typeof mgr.id === 'string' && isValidGuid(mgr.id)) {
            expandedIds.push(mgr.id);
          }
        }
      }
      authorizedIds = new Set(expandedIds);
    } catch (err) {
      // Fail closed: if caller-to-user authorization cannot be validated, do not
      // fetch the photo with application permissions.
      context.warn('getPhoto: sponsor validation lookup failed — denying photo request', err);
      return jsonErrorResponse(
        request, 503, correlationId,
        'Failed to validate photo access.', 'PHOTO_VALIDATION_FAILED',
        'Unable to validate whether the requested photo belongs to the caller\'s authorized sponsor set',
        true
      );
    }
  }

  if (!authorizedIds.has(userId)) {
    context.warn('getPhoto: requested userId not in authorized set', {
      correlationId,
      callerOid: redactGuid(callerOid),
      userId: redactGuid(userId),
    });
    return jsonErrorResponse(
      request, 403, correlationId,
      'Forbidden', 'PHOTO_ACCESS_DENIED',
      'The requested user ID is not in the caller\'s authorized photo set', false
    );
  }

  try {
    const buffer = await withTimeout(
      client.api(`/users/${userId}/photo/$value`).responseType(ResponseType.ARRAYBUFFER).get() as Promise<ArrayBuffer>,
      getTimeoutMs('PHOTO_TIMEOUT_MS', DEFAULT_PHOTO_TIMEOUT_MS),
      'Graph photo fetch'
    );
    return jsonResponse({ photoUrl: arrayBufferToDataUrl(buffer) }, 200, request, correlationId);
  } catch (error) {
    if (error instanceof GraphError && error.statusCode === 404) {
      return {
        status: 404,
        body: JSON.stringify({ error: 'Not Found', reasonCode: 'PHOTO_NOT_FOUND' }),
        headers: {
          'Content-Type': 'application/json',
          'x-correlation-id': correlationId,
          'x-api-version': FUNCTION_VERSION,
          ...corsHeaders(request),
        },
      };
    }
    context.warn('getPhoto: error fetching photo', { userId: redactGuid(userId), error });
    const status = error instanceof TimeoutError
      ? 504
      : getValidHttpStatus(
          error instanceof GraphError ? error.statusCode : (error as { statusCode?: number }).statusCode
        );
    return jsonErrorResponse(
      request, status, correlationId,
      'Failed to retrieve photo.',
      status === 504 ? 'PHOTO_TIMEOUT' : 'PHOTO_FETCH_FAILED',
      'Photo retrieval failed in backend processing',
      status >= 500 || status === 429
    );
  }
}

async function safeGetPhoto(
  request: HttpRequest,
  context: InvocationContext
): Promise<HttpResponseInit> {
  try {
    const response = await getPhoto(request, context);
    return { ...response, status: getValidHttpStatus(response.status) };
  } catch (error) {
    const correlationId = getCorrelationId(request);
    context.error('Unhandled exception in safeGetPhoto:', error);
    return jsonErrorResponse(
      request, 500, correlationId,
      'Failed to retrieve photo.', 'UNHANDLED_EXCEPTION',
      'Unexpected backend exception occurred', true
    );
  }
}

app.http('getPhoto', {
  methods: ['GET', 'OPTIONS'],
  authLevel: 'anonymous', // Authentication is enforced by EasyAuth, not the function key.
  handler: safeGetPhoto,
});

app.http('getLatestRelease', {
  methods: ['GET', 'OPTIONS'],
  authLevel: 'anonymous', // Authentication is enforced by EasyAuth, not the function key.
  handler: async (request: HttpRequest): Promise<HttpResponseInit> => {
    if (request.method === 'OPTIONS') {
      return { status: 204, headers: corsHeaders(request) };
    }
    return {
      status: 200,
      body: JSON.stringify({
        latestVersion: latestGitHubVersion ?? null,
        url: latestGitHubReleaseUrl ?? null,
        functionVersion: FUNCTION_VERSION,
      }),
      headers: {
        'Content-Type': 'application/json',
        'x-api-version': FUNCTION_VERSION,
        ...corsHeaders(request),
      },
    };
  },
});
