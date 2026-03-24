import { app, HttpRequest, HttpResponseInit, InvocationContext } from '@azure/functions';
import { createHmac, randomUUID, timingSafeEqual } from 'crypto';
import { DefaultAzureCredential } from '@azure/identity';
import { Client, GraphError } from '@microsoft/microsoft-graph-client';
import { TokenCredentialAuthenticationProvider } from '@microsoft/microsoft-graph-client/authProviders/azureTokenCredentials';
import packageJson from '../package.json';

/** Version string exposed in every response via X-Api-Version and used for client/server mismatch detection. */
const FUNCTION_VERSION: string = packageJson.version;

/**
 * A guest can have at most 5 sponsors.
 * Keeping the cap explicit avoids unnecessary Graph work and bounds the size
 * of our batch requests.
 */
const MAX_SPONSORS = 5;
const DEFAULT_SPONSOR_LOOKUP_TIMEOUT_MS = 5000;
const DEFAULT_PRESENCE_TIMEOUT_MS = 2500;
const DEFAULT_BATCH_TIMEOUT_MS = 4000;

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
 *      10 req / 60 s per IP (hardcoded — anonymous callers are not legitimate).
 *
 * 2. Authenticated callers (valid OID) — disabled by default to avoid false
 *    positives when multiple web parts are on one page or the user reloads
 *    quickly.  Enable only for incident response via:
 *      RATE_LIMIT_ENABLED=true
 *      RATE_LIMIT_MAX_REQUESTS=20   (optional, default 20)
 *      RATE_LIMIT_WINDOW_MS=60000  (optional, default 60 000 ms)
 *
 * Note: on a Consumption plan each instance has its own counter.
 */
const ANON_RATE_LIMIT_MAX_REQUESTS = 10;
const ANON_RATE_LIMIT_WINDOW_MS = 60_000;
const DEFAULT_RATE_LIMIT_ENABLED = false;
const DEFAULT_RATE_LIMIT_MAX_REQUESTS = 20;
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
function createPresenceToken(callerOid: string, sponsorIds: string[]): string | undefined {
  const secret = process.env.PRESENCE_TOKEN_SECRET;
  if (!secret) return undefined;
  const payload = JSON.stringify({
    jti: randomUUID(),
    oid: callerOid,
    ids: sponsorIds,
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

  let payload: { jti?: unknown; oid?: unknown; ids?: unknown; iat?: unknown; exp?: unknown };
  try {
    payload = JSON.parse(Buffer.from(payloadB64, 'base64url').toString('utf8')) as {
      jti?: unknown; oid?: unknown; ids?: unknown; iat?: unknown; exp?: unknown;
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
  return new Set(ids);
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
 * Validates that the authenticated EasyAuth principal belongs to our tenant and API audience.
 *
 * Performs two checks on EasyAuth principal claims:
 * 1. tid claim matches our tenant ID
 * 2. aud claim matches our API's audience URI
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
 * Returns true when the assignedPlans array contains at least one active Teams plan.
 * Teams plans are identified by service === 'TeamspaceAPI' with capabilityStatus === 'Enabled'.
 * Returns false for any non-array input (undefined, null, unexpected type).
 */
function assignedPlansHaveTeams(plans: unknown): boolean {
  if (!Array.isArray(plans)) return false;
  return (plans as Array<{ service?: unknown; capabilityStatus?: unknown }>)
    .some(p => p.service === 'TeamspaceAPI' && p.capabilityStatus === 'Enabled');
}

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
 * HTTP GET – returns the sponsors of the calling user.
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
    const anonLimit = checkRateLimit(`anon:${clientIp}`, ANON_RATE_LIMIT_MAX_REQUESTS, ANON_RATE_LIMIT_WINDOW_MS);
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
    const rateLimitResult = checkRateLimit(callerOid, rateLimitConfig.maxRequests, rateLimitConfig.windowMs);
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

  // Version compatibility check — warn in Azure Monitor when the client and function versions differ.
  const clientVersion = request.headers.get('x-client-version');
  if (clientVersion && clientVersion !== FUNCTION_VERSION) {
    context.warn('Client/function version mismatch', {
      clientVersion,
      functionVersion: FUNCTION_VERSION,
      correlationId,
    });
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

  context.log(`Fetching sponsors for caller ${redactGuid(callerOid)}`);

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
      url: `/users/${sponsor.id}?$select=id,accountEnabled,assignedPlans${hasMailboxSettings ? ',mailboxSettings' : ''}&$expand=manager($select=id,displayName,givenName,surname,jobTitle,department,accountEnabled)`,
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

    const [presenceMap, batchResults] = await Promise.all([
      // Include callerOid so we can use the guest's own presence as a provisioning signal.
      hasPresenceReadAll ? fetchPresences(client, [...sponsorIds, callerOid], context) : Promise.resolve(new Map<string, { availability?: string; activity?: string }>()),
      executeBatch(
        client,
        sponsorBatchRequests,
        getTimeoutMs('BATCH_TIMEOUT_MS', DEFAULT_BATCH_TIMEOUT_MS),
        'Graph sponsor detail batch'
      ),
    ]);

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
      // Determine userPurpose from mailboxSettings when the permission is available.
      // Accepted values: 'user' (cloud mailbox) and 'linked' (on-premises mailbox).
      // Any other value (shared, room, equipment, …) means this is not a real person
      // that should appear as a sponsor.  When mailboxSettings is absent (permission
      // not granted or user has no Exchange mailbox object) we fail-open so that a
      // missing permission never silently hides sponsors.
      let hasUserMailbox = true;
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
      // Treat missing, soft-deleted, disabled, or non-user-mailbox sponsors as unavailable.
      const exists = existsResponse?.status === 404 ? false
        : sponsorEnabled !== false && hasUserMailbox;

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
        return out;
      });
    const unavailableCount = perSponsorResults.filter(r => !r.exists).length;
    const unavailableSponsors: ISponsor[] = perSponsorResults
      .filter(r => !r.exists)
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

    // Issue a signed presence token so subsequent getPresence polls can be
    // validated without server-side state or extra Graph calls.
    const presenceToken = createPresenceToken(callerOid, sponsorIds);
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
 * HTTP GET – returns Microsoft Teams presence for a list of Entra object IDs.
 *
 * Designed to serve presence refresh polls from the SPFx web part after the
 * initial sponsor load.  Using application permissions (Managed Identity) keeps
 * presence up-to-date reliably for guest callers — unlike the delegated
 * Presence.Read.All scope, which may silently return empty results for guests
 * on tenants with restrictive guest-access policies.
 *
 * Query parameter:
 *   ids  – comma-separated list of up to MAX_SPONSORS Entra object IDs (GUIDs).
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
      `anon:${clientIp}`,
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
    const rateLimitResult = checkRateLimit(callerOid, rateLimitConfig.maxRequests, rateLimitConfig.windowMs);
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
        // Fail-open: a transient Graph error should not permanently block presence
        // polls.  Log and treat all requested IDs as authorized for this call only.
        context.warn('getPresence: sponsor validation lookup failed — failing open for this request', sponsorError);
        authorizedIds = new Set(ids);
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
 * HTTP GET – lightweight health-check endpoint.
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
