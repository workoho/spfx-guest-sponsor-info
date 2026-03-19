import { app, HttpRequest, HttpResponseInit, InvocationContext } from '@azure/functions';
import { ManagedIdentityCredential } from '@azure/identity';
import { Client, GraphError, MiddlewareFactory } from '@microsoft/microsoft-graph-client';
import { TokenCredentialAuthenticationProvider } from '@microsoft/microsoft-graph-client/authProviders/azureTokenCredentials';

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
 * In-memory sliding-window rate limiter (per caller OID, per warm instance).
 * 20 requests / 60 s is generous for legitimate use (one page load = one request)
 * while still capping runaway request loops or unintentional retry storms.
 * Note: on a Consumption plan each instance has its own counter; the effective
 * limit across N concurrent instances is N × 20/min — acceptable for this use case.
 */
const RATE_LIMIT_MAX_REQUESTS = 20;
const RATE_LIMIT_WINDOW_MS = 60_000;
const rateLimitMap = new Map<string, number[]>();

/**
 * Bot and scanner protection: detect and block suspicious User-Agent headers.
 * Legitimate browser/SPFx clients have recognizable User-Agents; scanners often
 * use minimal or no User-Agent, or use tools like curl, wget, Nmap, sqlmap, etc.
 */
const SUSPICIOUS_USER_AGENTS = /^(curl|wget|python|nmap|masscan|nikto|sqlmap|dirbuster|burpsuite|zaproxy|metasploit|nessus)|\b(bot|scraper|crawler|spider|scanner|recon)\b|^-?$/i;
const MAX_HEADER_SIZE_BYTES = 8192;  // Per-header limit; browsers typically send <1KB
const MAX_BODY_SIZE_BYTES = 1024;    // GET requests have no body; 1KB guards against malformed requests

/**
 * Cached result of permission detection for MailboxSettings.Read.
 * undefined = not yet checked; true/false = result cached for the lifetime
 * of this warm function instance.
 */
let cachedMailboxSettingsDetection: Promise<boolean> | undefined;

/**
 * Inspects the JWT access token obtained by the Managed Identity to determine
 * whether the MailboxSettings.Read application permission has been granted.
 *
 * The token is already fetched by the Azure Identity SDK for Graph API calls;
 * calling getToken() here returns the same cached token — no extra network
 * request is made.  The result is cached at module level so the check only
 * runs once per warm function instance.
 *
 * The result is stored as a Promise so that concurrent invocations during
 * a cold start share a single in-flight token inspection instead of each
 * starting their own (eliminates the require-atomic-updates race pattern).
 *
 * When this permission is present the caller can add mailboxSettings to the
 * per-sponsor $select and use userPurpose to filter out non-user mailboxes
 * (shared, room, equipment, …).  When it is absent the filter is silently
 * skipped — the function degrades gracefully without any error.
 */
async function detectMailboxSettingsPermission(
  credential: ManagedIdentityCredential,
  context: InvocationContext
): Promise<boolean> {
  // Assign synchronously before any await so concurrent callers share one promise.
  if (!cachedMailboxSettingsDetection) {
    cachedMailboxSettingsDetection = (async (): Promise<boolean> => {
      try {
        const token = await credential.getToken('https://graph.microsoft.com/.default');
        // JWT payload is the second segment, base64url-encoded.
        const payloadJson = Buffer.from(token.token.split('.')[1], 'base64url').toString('utf8');
        const payload = JSON.parse(payloadJson) as { roles?: unknown };
        const roles = Array.isArray(payload.roles) ? (payload.roles as string[]) : [];
        const result = roles.includes('MailboxSettings.Read');
        context.log(`MailboxSettings.Read permission: ${result}`);
        return result;
      } catch (error) {
        // If token inspection fails for any reason, degrade gracefully.
        context.warn('Could not inspect token roles — mailbox filter disabled.', error);
        return false;
      }
    })();
  }
  return cachedMailboxSettingsDetection;
}

function checkRateLimit(userId: string): { allowed: boolean; retryAfterSeconds?: number } {
  const now = Date.now();
  const windowStart = now - RATE_LIMIT_WINDOW_MS;

  // Lazy cleanup: purge entries with no recent timestamps to bound map memory usage.
  if (rateLimitMap.size > 500) {
    for (const [key, ts] of rateLimitMap) {
      if (ts.every(t => t <= windowStart)) rateLimitMap.delete(key);
    }
  }

  const recent = (rateLimitMap.get(userId) ?? []).filter(t => t > windowStart);
  if (recent.length >= RATE_LIMIT_MAX_REQUESTS) {
    const oldest = Math.min(...recent);
    const retryAfterSeconds = Math.ceil((oldest + RATE_LIMIT_WINDOW_MS - now) / 1000);
    rateLimitMap.set(userId, recent);
    return { allowed: false, retryAfterSeconds };
  }

  recent.push(now);
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

/**
 * Evaluates whether a request looks suspicious based on User-Agent.
 * Returns true for obvious scanner/bot signatures (curl, wget, nmap, etc.) and
 * requests with missing or clearly malicious User-Agents.
 */
function isSuspiciousUserAgent(userAgent: string | null): boolean {
  if (!userAgent || userAgent.trim().length === 0) {
    return true;  // Missing User-Agent is suspicious in a browser context.
  }
  return SUSPICIOUS_USER_AGENTS.test(userAgent);
}

/**
 * Validates that request headers do not exceed reasonable size (defence against
 * header-injection and Denial of Service attacks via oversized headers).
 */
function validateRequestSize(headers: Record<string, string | string[]>): { valid: boolean; reason?: string } {
  for (const [key, value] of Object.entries(headers)) {
    const headerStr = typeof value === 'string' ? value : value.join(',');
    if (headerStr.length > MAX_HEADER_SIZE_BYTES) {
      return { valid: false, reason: `${key} exceeds max header size` };
    }
  }
  return { valid: true };
}

/**
 * Returns a short, loggable description of suspicious request patterns.
 */
function describeRequestAnomaly(userAgent: string | null, method: string): string {
  const parts: string[] = [];
  if (!userAgent || userAgent.trim().length === 0) {
    parts.push('missing-user-agent');
  } else if (SUSPICIOUS_USER_AGENTS.test(userAgent)) {
    parts.push(`bot-like-ua:${userAgent.substring(0, 30)}`);
  }
  if (method !== 'GET' && method !== 'OPTIONS') {
    parts.push(`non-get-method:${method}`);
  }
  return parts.join(' | ') || 'none';
}

async function withTimeout<T>(
  promise: Promise<T>,
  timeoutMs: number,
  operation: string
): Promise<T> {
  return await new Promise<T>((resolve, reject) => {
    const timer = setTimeout(() => reject(new TimeoutError(operation, timeoutMs)), timeoutMs);

    void promise.then(
      value => {
        clearTimeout(timer);
        resolve(value);
      },
      error => {
        clearTimeout(timer);
        reject(error);
      }
    );
  });
}

/** Shape of a sponsor returned by this function (matches ISponsor in SPFx). */
interface ISponsor {
  id: string;
  displayName: string;
  mail?: string;
  jobTitle?: string;
  department?: string;
  officeLocation?: string;
  businessPhones: string[];
  mobilePhone?: string;
  presence?: string;
  managerDisplayName?: string;
  managerJobTitle?: string;
  /** Manager's Entra ID — used by the SPFx client to fetch the manager photo progressively. */
  managerId?: string;
  /** True when the sponsor has an active Microsoft Teams license. */
  hasTeams?: boolean;
}

/**
 * Shape of the JSON response (matches ISponsorsResult in SPFx).
 * unavailableCount includes sponsors that are deleted, soft-deleted, or disabled.
 */
interface ISponsorsResult {
  activeSponsors: ISponsor[];
  unavailableCount: number;
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
function resolveCallerOid(request: HttpRequest): string | null {
  // EasyAuth sets X-MS-CLIENT-PRINCIPAL-ID to the validated caller OID.
  const easyAuthOid = request.headers.get('x-ms-client-principal-id');
  if (easyAuthOid) return isValidGuid(easyAuthOid) ? easyAuthOid : null;

  // Local dev fallback — never accepted in production.
  if (process.env['NODE_ENV'] !== 'production') {
    const devOid = request.headers.get('x-dev-user-oid');
    if (devOid) return isValidGuid(devOid) ? devOid : null;
  }

  return null;
}

/**
 * Fetches presence for a list of user IDs in a single batched Graph call.
 * Returns a map of userId → availability string.
 * Silently returns an empty map on any error or timeout (presence is non-critical).
 */
async function fetchPresences(
  client: Client,
  userIds: string[],
  context: InvocationContext
): Promise<Map<string, string>> {
  const map = new Map<string, string>();
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
      for (const entry of response.value as Array<{ id: string; availability: string }>) {
        // Only accept entries whose ID was in our request and whose availability
        // is a plain alphabetic string (no injection vectors).
        if (
          entry.id &&
          requestedIds.has(entry.id) &&
          entry.availability &&
          /^[A-Za-z]+$/.test(entry.availability)
        ) {
          map.set(entry.id, entry.availability);
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
 * HTTP GET – returns the sponsors of the calling user.
 *
 * Authentication is handled by EasyAuth (Azure App Service Authentication).
 * The function reads the caller OID from the X-MS-CLIENT-PRINCIPAL-ID header
 * that EasyAuth sets after validating the Bearer token.
 *
 * The Managed Identity of the Function App is used to call Microsoft Graph
 * with application permissions (User.Read.All, Presence.Read.All).
 * No client secrets are stored anywhere.
 */
export /**
 * Validates and normalizes HTTP status codes. Returns a valid HTTP status in the
 * range 200–599; defaults to 500 if the provided code is undefined or out of range.
 * This guards against GraphError or other exceptions providing invalid status codes
 * that would crash the Azure Functions runtime.
 */
function getValidHttpStatus(code: unknown): number {
  if (typeof code === 'number' && code >= 200 && code <= 599) return code;
  return 500;
}

async function getGuestSponsors(
  request: HttpRequest,
  context: InvocationContext
): Promise<HttpResponseInit> {
  // CORS preflight.
  if (request.method === 'OPTIONS') {
    return {
      status: 204,
      headers: corsHeaders(request),
    };
  }

  // Reject requests that are clearly not from a real browser/client.
  const userAgent = request.headers.get('user-agent');
  if (isSuspiciousUserAgent(userAgent)) {
    const anomaly = describeRequestAnomaly(userAgent, request.method);
    context.warn(`Suspicious request rejected. Anomalies: ${anomaly}`);
    return {
      status: 403,
      body: JSON.stringify({ error: 'Forbidden' }),
      headers: { 'Content-Type': 'application/json', ...corsHeaders(request) },
    };
  }

  // Validate request header sizes (defence against injection and DoS via oversized headers).
  const headerValidation = validateRequestSize(Object.fromEntries(
    Array.from(request.headers.entries()).map(([k, v]) => [k, v ?? ''])
  ));
  if (!headerValidation.valid) {
    context.warn(`Request validation failed: ${headerValidation.reason}`);
    return {
      status: 400,
      body: JSON.stringify({ error: 'Bad Request' }),
      headers: { 'Content-Type': 'application/json', ...corsHeaders(request) },
    };
  }

  const callerOid = resolveCallerOid(request);
  if (!callerOid) {
    context.warn('No caller OID found — EasyAuth may not be configured.');
    return {
      status: 401,
      body: JSON.stringify({ error: 'Unauthorized' }),
      headers: { 'Content-Type': 'application/json', ...corsHeaders(request) },
    };
  }

  const rateLimitResult = checkRateLimit(callerOid);
  if (!rateLimitResult.allowed) {
    return {
      status: 429,
      body: JSON.stringify({ error: 'Too Many Requests' }),
      headers: {
        'Content-Type': 'application/json',
        'Retry-After': String(rateLimitResult.retryAfterSeconds ?? 60),
        ...corsHeaders(request),
      },
    };
  }

  context.log(`Fetching sponsors for caller ${redactGuid(callerOid)}`);

  try {
    const credential = new ManagedIdentityCredential();
    const hasMailboxSettings = await detectMailboxSettingsPermission(credential, context);
    const authProvider = new TokenCredentialAuthenticationProvider(credential, {
      scopes: ['https://graph.microsoft.com/.default'],
    });
    // Build a middleware chain that includes automatic retry with exponential
    // back-off for 429 (throttled) and transient 5xx responses (up to 3 retries,
    // starting at 3 s — the library default).  RedirectHandler and the auth
    // handler are also included via the factory's default chain.
    const middleware = MiddlewareFactory.getDefaultMiddlewareChain(authProvider);
    const client = Client.initWithMiddleware({ middleware: middleware[0] });
    const response = await withTimeout(
      client
        .api(`/users/${callerOid}/sponsors`)
        .select('id,displayName,mail,jobTitle,department,officeLocation,businessPhones,mobilePhone')
        // Enforce the cap at the Graph level to avoid fetching more items than we will process.
        .top(MAX_SPONSORS)
        .get(),
      getTimeoutMs('SPONSOR_LOOKUP_TIMEOUT_MS', DEFAULT_SPONSOR_LOOKUP_TIMEOUT_MS),
      'Graph sponsor lookup'
    );

    if (!response?.value) {
      const result: ISponsorsResult = { activeSponsors: [], unavailableCount: 0 };
      return jsonResponse(result, 200, request);
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
        mail: (item.mail as string) || undefined,
        jobTitle: (item.jobTitle as string) || undefined,
        department: (item.department as string) || undefined,
        officeLocation: (item.officeLocation as string) || undefined,
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
      url: `/users/${sponsor.id}?$select=id,accountEnabled,assignedPlans${hasMailboxSettings ? ',mailboxSettings' : ''}&$expand=manager($select=id,displayName,jobTitle,accountEnabled)`,
    }));

    const [presenceMap, sponsorBatchResults] = await Promise.all([
      fetchPresences(client, sponsorIds, context),
      executeBatch(
        client,
        sponsorBatchRequests,
        getTimeoutMs('BATCH_TIMEOUT_MS', DEFAULT_BATCH_TIMEOUT_MS),
        'Graph sponsor detail batch'
      ),
    ]);

    const perSponsorResults = candidates.map((sponsor, index) => {
      const existsResponse = sponsorBatchResults.get(`exists-${index}`);
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
        const mb = existsBody['mailboxSettings'];
        if (mb !== null && mb !== undefined) {
          const userPurpose = typeof (mb as Record<string, unknown>)['userPurpose'] === 'string'
            ? (mb as Record<string, unknown>)['userPurpose'] as string
            : undefined;
          if (userPurpose !== undefined) {
            hasUserMailbox = userPurpose === 'user' || userPurpose === 'linked';
          }
        }
      }
      // Treat missing, soft-deleted, disabled, or non-user-mailbox sponsors as unavailable.
      const exists = existsResponse?.status === 404 ? false
        : sponsorEnabled !== false && hasUserMailbox;

      const hasTeams = assignedPlansHaveTeams(existsBody?.['assignedPlans']);

      // Manager is inlined via $expand — read directly from the user body.
      const managerBody = existsBody !== undefined && existsBody['manager'] !== null
        ? existsBody['manager'] as Record<string, unknown> | undefined
        : undefined;

      let managerDisplayName: string | undefined;
      let managerJobTitle: string | undefined;
      let managerId: string | undefined;

      if (managerBody !== undefined) {
        const managerEnabled = getBooleanValue(managerBody, 'accountEnabled');
        if (managerEnabled !== false) {
          managerDisplayName = getStringValue(managerBody, 'displayName');
          managerJobTitle = getStringValue(managerBody, 'jobTitle');
          const mid = getStringValue(managerBody, 'id');
          if (mid && isValidGuid(mid)) managerId = mid;
        }
      }

      return {
        sponsor: {
          ...sponsor,
          managerDisplayName,
          managerJobTitle,
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
        if (s.jobTitle !== undefined)            out.jobTitle = s.jobTitle;
        if (s.department !== undefined)          out.department = s.department;
        if (s.officeLocation !== undefined)      out.officeLocation = s.officeLocation;
        if (s.mobilePhone !== undefined)         out.mobilePhone = s.mobilePhone;
        if (s.managerDisplayName !== undefined)  out.managerDisplayName = s.managerDisplayName;
        if (s.managerJobTitle !== undefined)     out.managerJobTitle = s.managerJobTitle;
        if (s.managerId !== undefined)           out.managerId = s.managerId;
        out.hasTeams = s.hasTeams;  // always a boolean from the proxy; undefined only in direct path
        const presence = presenceMap.get(s.id);
        if (presence !== undefined)              out.presence = presence;
        return out;
      });
    const unavailableCount = perSponsorResults.filter(r => !r.exists).length;

    const result: ISponsorsResult = { activeSponsors, unavailableCount };
    return jsonResponse(result, 200, request);
  } catch (error) {
    if (error instanceof GraphError) {
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
    return {
      status,
      body: JSON.stringify({ error: 'Failed to retrieve sponsor information.' }),
      headers: { 'Content-Type': 'application/json', ...corsHeaders(request) },
    };
  }
}

function jsonResponse(body: unknown, status: number, request: HttpRequest): HttpResponseInit {
  return {
    status,
    body: JSON.stringify(body),
    headers: { 'Content-Type': 'application/json', ...corsHeaders(request) },
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
  const allowedOrigin = process.env['CORS_ALLOWED_ORIGIN'] ?? '';

  const headers: Record<string, string> = {
    // CORS
    'Access-Control-Allow-Methods': 'GET, OPTIONS',
    'Access-Control-Allow-Headers': 'Authorization, Content-Type',
    // Security headers (defence-in-depth even without a WAF)
    'X-Content-Type-Options': 'nosniff',           // Prevent content-type sniffing attacks
    'X-Frame-Options': 'DENY',                      // Block framing in other sites
    'Strict-Transport-Security': 'max-age=31536000; includeSubDomains; preload',  // Enforce HTTPS
    'X-XSS-Protection': '1; mode=block',            // Legacy XSS filter (for older browsers)
  };

  if (allowedOrigin && origin === allowedOrigin) {
    headers['Access-Control-Allow-Origin'] = origin;
  }

  return headers;
}

app.http('getGuestSponsors', {
  methods: ['GET', 'OPTIONS'],
  authLevel: 'anonymous', // Authentication is enforced by EasyAuth, not the function key.
  handler: getGuestSponsors,
});
