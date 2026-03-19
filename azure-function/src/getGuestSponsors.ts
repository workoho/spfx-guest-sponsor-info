import { app, HttpRequest, HttpResponseInit, InvocationContext } from '@azure/functions';
import { ManagedIdentityCredential } from '@azure/identity';
import { Client } from '@microsoft/microsoft-graph-client';
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
const DEFAULT_PHOTO_BATCH_TIMEOUT_MS = 2500;

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
  photoUrl?: string;
  presence?: string;
  managerDisplayName?: string;
  managerJobTitle?: string;
  managerPhotoUrl?: string;
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
 * Converts an ArrayBuffer containing JPEG bytes into a base64-encoded data URL.
 */
function arrayBufferToDataUrl(buffer: ArrayBuffer): string {
  const bytes = new Uint8Array(buffer);
  let binary = '';
  for (let i = 0; i < bytes.byteLength; i++) {
    binary += String.fromCharCode(bytes[i]);
  }
  return `data:image/jpeg;base64,${Buffer.from(bytes).toString('base64')}`;
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

function getPhotoFromBatchResponse(item: IBatchResponseItem | undefined): string | undefined {
  const photoBody = item?.body;
  const base64 = photoBody !== undefined ? getStringValue(photoBody, 'value') : undefined;
  return base64 !== undefined ? `data:image/jpeg;base64,${base64}` : undefined;
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
export async function getGuestSponsors(
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

  const credential = new ManagedIdentityCredential();
  const authProvider = new TokenCredentialAuthenticationProvider(credential, {
    scopes: ['https://graph.microsoft.com/.default'],
  });
  const client = Client.initWithMiddleware({ authProvider });

  try {
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

    const sponsorBatchRequests: IBatchRequest[] = candidates.flatMap((sponsor, index) => ([
      {
        id: `exists-${index}`,
        method: 'GET',
        url: `/users/${sponsor.id}?$select=id,accountEnabled`,
      },
      {
        id: `photo-${index}`,
        method: 'GET',
        url: `/users/${sponsor.id}/photo/$value`,
      },
      {
        id: `manager-${index}`,
        method: 'GET',
        url: `/users/${sponsor.id}/manager?$select=id,displayName,jobTitle,accountEnabled`,
      },
    ]));

    const [presenceMap, sponsorBatchResults] = await Promise.all([
      fetchPresences(client, sponsorIds, context),
      executeBatch(
        client,
        sponsorBatchRequests,
        getTimeoutMs('BATCH_TIMEOUT_MS', DEFAULT_BATCH_TIMEOUT_MS),
        'Graph sponsor detail batch'
      ),
    ]);

    const managerPhotoRequests: IBatchRequest[] = [];
    const managerPhotoRequestIds = new Map<string, string>();

    for (const [index, sponsor] of candidates.entries()) {
      const managerBody = sponsorBatchResults.get(`manager-${index}`)?.body;
      if (managerBody === undefined) continue;

      const managerId = getStringValue(managerBody, 'id');
      const managerEnabled = getBooleanValue(managerBody, 'accountEnabled');
      if (!managerId || !isValidGuid(managerId) || managerEnabled === false) continue;

      const requestId = `manager-photo-${index}`;
      managerPhotoRequests.push({
        id: requestId,
        method: 'GET',
        url: `/users/${managerId}/photo/$value`,
      });
      managerPhotoRequestIds.set(sponsor.id, requestId);
    }

    let managerPhotoBatchResults = new Map<string, IBatchResponseItem>();
    try {
      managerPhotoBatchResults = await executeBatch(
        client,
        managerPhotoRequests,
        getTimeoutMs('PHOTO_BATCH_TIMEOUT_MS', DEFAULT_PHOTO_BATCH_TIMEOUT_MS),
        'Graph manager photo batch'
      );
    } catch (error) {
      // Manager photos are supplemental — degrade to initials if the photo batch is slow or fails.
      context.warn('Manager photo batch degraded.', error);
    }

    const perSponsorResults = candidates.map((sponsor, index) => {
      const existsResponse = sponsorBatchResults.get(`exists-${index}`);
      const existsBody = existsResponse?.body;
      const sponsorEnabled = existsBody !== undefined
        ? getBooleanValue(existsBody, 'accountEnabled')
        : undefined;
      // Treat missing, soft-deleted, or disabled sponsors as unavailable.
      const exists = existsResponse?.status === 404 ? false : sponsorEnabled !== false;

      const managerBody = sponsorBatchResults.get(`manager-${index}`)?.body;
      const photoUrl = getPhotoFromBatchResponse(sponsorBatchResults.get(`photo-${index}`));

      let managerDisplayName: string | undefined;
      let managerJobTitle: string | undefined;
      let managerPhotoUrl: string | undefined;

      if (managerBody !== undefined) {
        const managerEnabled = getBooleanValue(managerBody, 'accountEnabled');
        // Managers are shown only when they still resolve through the active user view
        // and are not disabled.
        if (managerEnabled !== false) {
          managerDisplayName = getStringValue(managerBody, 'displayName');
          managerJobTitle = getStringValue(managerBody, 'jobTitle');
          const managerPhotoRequestId = managerPhotoRequestIds.get(sponsor.id);
          managerPhotoUrl = managerPhotoRequestId !== undefined
            ? getPhotoFromBatchResponse(managerPhotoBatchResults.get(managerPhotoRequestId))
            : undefined;
        }
      }

      return {
        sponsor: {
          ...sponsor,
          photoUrl,
          managerDisplayName,
          managerJobTitle,
          managerPhotoUrl,
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
        if (s.photoUrl !== undefined)            out.photoUrl = s.photoUrl;
        if (s.managerDisplayName !== undefined)  out.managerDisplayName = s.managerDisplayName;
        if (s.managerJobTitle !== undefined)     out.managerJobTitle = s.managerJobTitle;
        if (s.managerPhotoUrl !== undefined)     out.managerPhotoUrl = s.managerPhotoUrl;
        const presence = presenceMap.get(s.id);
        if (presence !== undefined)              out.presence = presence;
        return out;
      });
    const unavailableCount = perSponsorResults.filter(r => !r.exists).length;

    const result: ISponsorsResult = { activeSponsors, unavailableCount };
    return jsonResponse(result, 200, request);
  } catch (error) {
    context.error('Error fetching sponsors:', error);
    const status = error instanceof TimeoutError
      ? 504
      : ((error as { statusCode?: number }).statusCode ?? 500);
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
 * Returns CORS headers allowing the request's origin when it matches a known pattern.
 * The allowed origin is configured via the CORS_ALLOWED_ORIGIN environment variable
 * (e.g. "https://contoso.sharepoint.com"). Falls back to the Function App CORS settings
 * configured in Azure when the env var is not set.
 */
function corsHeaders(request: HttpRequest): Record<string, string> {
  const origin = request.headers.get('origin') ?? '';
  const allowedOrigin = process.env['CORS_ALLOWED_ORIGIN'] ?? '';

  const headers: Record<string, string> = {
    'Access-Control-Allow-Methods': 'GET, OPTIONS',
    'Access-Control-Allow-Headers': 'Authorization, Content-Type',
  };

  if (allowedOrigin && origin === allowedOrigin) {
    headers['Access-Control-Allow-Origin'] = origin;
  } else if (!allowedOrigin && origin) {
    // When no env var is set, echo the origin back (Azure Function CORS handles enforcement).
    headers['Access-Control-Allow-Origin'] = origin;
  }

  return headers;
}

app.http('getGuestSponsors', {
  methods: ['GET', 'OPTIONS'],
  authLevel: 'anonymous', // Authentication is enforced by EasyAuth, not the function key.
  handler: getGuestSponsors,
});
