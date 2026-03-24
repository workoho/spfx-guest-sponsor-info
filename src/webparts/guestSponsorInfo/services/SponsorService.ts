import { MSGraphClientV3, AadHttpClient } from '@microsoft/sp-http';
import { ResponseType } from '@microsoft/microsoft-graph-client';
import { ISponsor } from './ISponsor';

/** Result returned by getSponsors. */
export interface ISponsorsResult {
  /** Sponsor accounts whose directory object still exists in Entra. */
  activeSponsors: ISponsor[];
  /**
   * Number of sponsor entries that were excluded because their directory object
   * could no longer be found (HTTP 404 – hard-deleted or past the soft-delete
   * recycle-bin period).  Accounts that are merely disabled (accountEnabled ===
   * false) still appear in activeSponsors because reading that property requires
   * User.Read.All, which exceeds the declared permission scope.
   */
  unavailableCount: number;
  /**
   * Whether the guest user’s Teams service account has been provisioned in this
   * tenant.  True = the guest can use Teams (chat/call buttons are functional).
   * False = the guest has not yet been added to a Team and cannot use Teams.
   * Undefined = could not be determined (function proxy not configured, or
   * neither TeamMember.Read.All nor Presence.Read.All is granted); the client
   * fails open and shows the buttons enabled.
   */
  guestHasTeamsAccess?: boolean;  /**
   * Short-lived HMAC-signed token issued by the Azure Function that authorizes
   * subsequent getPresence calls for exactly this caller and sponsor ID set.
   * Present only when the function has PRESENCE_TOKEN_SECRET configured.
   * The web part passes this as the `token` query parameter to `/api/getPresence`
   * so the function can validate IDs without server-side state or extra Graph calls.
   */
  presenceToken?: string;
  /**
   * Version of the Azure Function that served this response, read from the
   * X-Api-Version response header.  Only set on the proxy path; undefined
   * when sponsors are fetched directly via Graph (no function involved).
   */
  functionVersion?: string;
  /**
   * Sponsor profiles that were found in the directory but whose account is
   * disabled, deleted, or otherwise unavailable. Populated when all sponsors
   * are unavailable so the client can still render read-only tiles alongside
   * the "sponsor not available" notice.
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

interface IPresenceSnapshot {
  availability?: string;
  activity?: string;
}

/**
 * Returns true when the SPFx login name belongs to a Microsoft Entra guest account.
 * Guest UPNs always contain the "#EXT#" marker introduced by Entra external identity.
 */
export function isGuestUser(loginName: string): boolean {
  return loginName.indexOf('#EXT#') !== -1;
}

/**
 * Converts an ArrayBuffer containing JPEG bytes into a base64-encoded data URL.
 * Avoids Blob-URL leaks because data URLs do not require explicit cleanup.
 *
 * Uses chunked processing (8 KB at a time) so the main thread is never blocked
 * by a long loop over tens of thousands of individual bytes.
 */
function arrayBufferToDataUrl(buffer: ArrayBuffer): string {
  const bytes = new Uint8Array(buffer);
  const chunkSize = 8192;
  const chunks: string[] = [];
  for (let i = 0; i < bytes.byteLength; i += chunkSize) {
    chunks.push(String.fromCharCode(...Array.from(bytes.subarray(i, i + chunkSize))));
  }
  return `data:image/jpeg;base64,${btoa(chunks.join(''))}`;
}

/**
 * Checks whether a sponsor user still exists and retrieves their manager profile
 * in a single Graph call via $expand (saving one request per sponsor compared
 * to two separate calls).
 *
 * Returns `exists: false` only on HTTP 404 (hard-deleted or past soft-delete
 * recycle-bin period).  Any other error is treated as "still exists" (fail-open)
 * so a transient Graph outage does not incorrectly hide a sponsor card.
 *
 * Note: accountEnabled requires User.Read.All, which exceeds the declared
 * permission scope and is therefore not checked here.
 */
async function fetchSponsorDetails(
  client: MSGraphClientV3,
  userId: string
): Promise<{ exists: boolean; managerDisplayName?: string; managerGivenName?: string; managerSurname?: string; managerJobTitle?: string; managerDepartment?: string; managerId?: string }> {
  try {
    const user = await client
      .api(`/users/${userId}`)
      .select('id')
      .expand('manager($select=id,displayName,givenName,surname,jobTitle,department)')
      .get() as Record<string, unknown>;
    const managerRaw = user.manager as Record<string, unknown> | null | undefined;
    if (!managerRaw) return { exists: true };
    return {
      exists: true,
      managerId: (managerRaw.id as string) || undefined,
      managerDisplayName: (managerRaw.displayName as string) || undefined,
      managerGivenName: (managerRaw.givenName as string) || undefined,
      managerSurname: (managerRaw.surname as string) || undefined,
      managerJobTitle: (managerRaw.jobTitle as string) || undefined,
      managerDepartment: (managerRaw.department as string) || undefined,
    };
  } catch (error) {
    if ((error as { statusCode?: number }).statusCode === 404) return { exists: false };
    return { exists: true }; // Fail-open on transient errors.
  }
}

/**
 * Fetches the sponsors of the signed-in user via Microsoft Graph.
 * For each sponsor the function checks existence and retrieves manager details
 * in a single Graph call (via $expand), then fetches presence for all sponsors
 * in one batched call — both fan-outs run concurrently.
 *
 * Required delegated permissions (declared in package-solution.json):
 *   - User.Read          – read the signed-in user's own /me/sponsors relationship.
 *   - User.ReadBasic.All – read existence and profile photos of the sponsor users.
 *                          This is the narrowest permission that covers reading
 *                          another user's directory object.  "ReadBasic" exposes
 *                          only: displayName, givenName, surname, mail, photo.
 *                          It does NOT expose accountEnabled (which requires
 *                          User.Read.All and is therefore out of scope).
 */
/**
 * Fetches presence for a list of user IDs in a single batched Graph call.
 * Uses the Presence.Read.All delegated permission when consented.
 * Returns a map of userId → availability/activity strings.
 * Silently returns an empty map on any error (presence is optional).
 */
export async function fetchPresences(
  client: MSGraphClientV3,
  userIds: string[]
): Promise<Map<string, IPresenceSnapshot>> {
  const map = new Map<string, IPresenceSnapshot>();
  if (userIds.length === 0) return map;
  try {
    const response = await client
      .api('/communications/getPresencesByUserId')
      .post({ ids: userIds });
    if (response?.value) {
      for (const entry of response.value as Array<{ id: string; availability?: string; activity?: string }>) {
        const availability = entry.availability && /^[A-Za-z]+$/.test(entry.availability)
          ? entry.availability
          : undefined;
        const activity = entry.activity && /^[A-Za-z]+$/.test(entry.activity)
          ? entry.activity
          : undefined;
        if (entry.id && (availability || activity)) {
          map.set(entry.id, { availability, activity });
        }
      }
    }
  } catch {
    // Presence is supplemental — silently degrade when the call fails
    // (e.g. permission not yet granted by tenant admin).
  }
  return map;
}

export async function getSponsors(client: MSGraphClientV3): Promise<ISponsorsResult> {
  const response = await client
    .api('/me/sponsors')
    .select('id,displayName,givenName,surname,mail,jobTitle,department,officeLocation,streetAddress,postalCode,state,city,country,businessPhones,mobilePhone')
    .get();

  if (!response?.value) return { activeSponsors: [], unavailableCount: 0 };

  const items = response.value as Record<string, unknown>[];
  const candidates: ISponsor[] = items.map(item => ({
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
    businessPhones: (item.businessPhones as string[]) || [],
    mobilePhone: (item.mobilePhone as string) || undefined,
  }));

  const sponsorIds = candidates.map(s => s.id);

  // Fetch presence for all sponsors in one batched call, concurrently with per-sponsor work.
  const [presenceMap, perSponsorResults] = await Promise.all([
    fetchPresences(client, sponsorIds),
    Promise.all(
      candidates.map(async sponsor => {
        const { exists, ...managerInfo } = await fetchSponsorDetails(client, sponsor.id);
        return { sponsor: { ...sponsor, ...managerInfo }, exists };
      })
    ),
  ]);

  const activeSponsors = perSponsorResults
    .filter(r => r.exists)
    .map(r => {
      const presence = presenceMap.get(r.sponsor.id);
      return {
        ...r.sponsor,
        presence: presence?.availability,
        presenceActivity: presence?.activity,
      };
    });
  const unavailableSponsors = perSponsorResults
    .filter(r => !r.exists)
    .map(r => {
      const { id, displayName, givenName, surname, jobTitle, department } = r.sponsor;
      return { id, displayName, givenName, surname, jobTitle, department } as ISponsor;
    });
  const unavailableCount = unavailableSponsors.length;
  return {
    activeSponsors,
    unavailableCount,
    sponsorOrder: candidates.map(s => s.id),
    ...(unavailableSponsors.length > 0 ? { unavailableSponsors } : {}),
  };
}

/**
 * Fetches sponsor data via the Azure Function proxy instead of calling Graph directly.
 * The proxy authenticates the caller via EasyAuth and calls Graph with application
 * permissions (User.Read.All, optionally Presence.Read.All) using its Managed Identity.
 *
 * @param proxyUrl      - Full URL of the Azure Function endpoint.
 * @param aadHttpClient - Pre-acquired AAD HTTP client scoped to the function App Registration.
 * @param clientVersion - Optional web part version string sent as X-Client-Version request header
 *                        so the function can log a warning when versions differ.
 */

/**
/**
 * Lightweight health check against the Azure Function's `/api/ping` endpoint.
 * No Graph calls or guest-context required — used in edit mode to verify connectivity.
 * Returns the function version string from the `x-api-version` response header,
 * or `undefined` when the header is absent (older function deployments).
 * Throws when the function is unreachable or responds with a non-2xx status.
 */
export async function pingProxy(
  pingUrl: string,
  aadHttpClient: AadHttpClient
): Promise<string | undefined> {
  const response = await aadHttpClient.get(pingUrl, AadHttpClient.configurations.v1);
  if (!response.ok) {
    throw new Error(`Ping returned ${response.status}`);
  }
  return response.headers.get('x-api-version') ?? undefined;
}

export async function getSponsorsViaProxy(
  proxyUrl: string,
  aadHttpClient: AadHttpClient,
  clientVersion?: string
): Promise<ISponsorsResult> {
  const options = clientVersion
    ? { headers: { 'X-Client-Version': clientVersion } }
    : undefined;
  const response = await aadHttpClient.get(proxyUrl, AadHttpClient.configurations.v1, options);
  const functionVersion = response.headers.get('x-api-version') ?? undefined;
  if (!response.ok) {
    let reasonCode: string | undefined;
    let referenceId: string | undefined;
    let retryable: boolean | undefined;

    try {
      const payload = await response.json() as {
        reasonCode?: string;
        referenceId?: string;
        retryable?: boolean;
      };
      reasonCode = payload.reasonCode;
      referenceId = payload.referenceId;
      retryable = payload.retryable;
    } catch {
      // Ignore body parse issues and fall back to status-only error metadata.
    }

    const correlationFromHeader = response.headers.get('x-correlation-id') ?? undefined;
    const err = new Error(`Proxy returned ${response.status}${reasonCode ? ` (${reasonCode})` : ''}`);
    (err as { statusCode?: number }).statusCode = response.status;
    (err as { reasonCode?: string }).reasonCode = reasonCode;
    (err as { referenceId?: string }).referenceId = referenceId ?? correlationFromHeader;
    (err as { retryable?: boolean }).retryable = retryable;
    throw err;
  }
  let result: ISponsorsResult;
  try {
    result = await response.json() as ISponsorsResult;
  } catch (parseError) {
    // The proxy returned a non-JSON body (e.g. an HTML login-redirect or gateway
    // error page with a 2xx status).  Wrap it so the caller gets a structured error
    // with statusCode and retryable set — rather than a bare SyntaxError with no
    // metadata that leaves all error-handling fields undefined.
    const err = new Error(
      `Proxy returned non-JSON response (HTTP ${response.status}): ${(parseError as Error).message}`
    );
    (err as { statusCode?: number }).statusCode = response.status;
    (err as { retryable?: boolean }).retryable = false;
    throw err;
  }
  if (functionVersion) result.functionVersion = functionVersion;
  return result;
}

/**
 * Fetches presence for a list of user IDs via the Azure Function proxy.
 *
 * The proxy calls Graph with application permissions (Managed Identity), which
 * work reliably for guest callers — unlike the delegated Presence.Read.All scope
 * that may silently return empty results for guests on tenants with restrictive
 * guest-access policies.
 *
 * Fails open (empty map) on any error so the existing presence data stays on
 * screen — identical degradation behaviour to the direct delegated path.
 *
 * @param presenceUrl   - Full URL of the `/api/getPresence` Function endpoint.
 * @param aadHttpClient - Pre-acquired AAD HTTP client scoped to the Function App Registration.
 * @param userIds       - Entra object IDs whose presence should be refreshed.
 * @param presenceToken - Optional signed token from the last getGuestSponsors response.
 *                        When present, the function validates IDs against the token
 *                        instead of making an extra Graph call.
 */
export async function getPresencesViaProxy(
  presenceUrl: string,
  aadHttpClient: AadHttpClient,
  userIds: string[],
  presenceToken?: string
): Promise<{ map: Map<string, { availability?: string; activity?: string }>; presenceToken?: string }> {
  const map = new Map<string, { availability?: string; activity?: string }>();
  if (userIds.length === 0) return { map };
  try {
    const params = new URLSearchParams({ ids: userIds.join(',') });
    const url = `${presenceUrl}?${params.toString()}`;
    const headers: Record<string, string> = {};
    if (presenceToken) headers['X-Presence-Token'] = presenceToken;
    const response = await aadHttpClient.get(url, AadHttpClient.configurations.v1,
      Object.keys(headers).length > 0 ? { headers } : undefined
    );
    if (!response.ok) return { map }; // fail-open: preserve existing presence data
    const data = await response.json() as {
      presences?: Array<{ id: string; availability?: string; activity?: string }>;
      presenceToken?: string;
    };
    for (const entry of data.presences ?? []) {
      if (entry.id) {
        map.set(entry.id, {
          availability: entry.availability,
          activity: entry.activity,
        });
      }
    }
    return { map, presenceToken: data.presenceToken };
  } catch {
    // Presence is supplemental — silently degrade when the proxy call fails.
  }
  return { map };
}

/**
 * Progressively fetches photos for a list of sponsors and their managers,
 * calling onUpdate for each sponsor as soon as its photo (or its manager's
 * photo) arrives.  Errors are silently swallowed — the card stays in its
 * initials-fallback state.
 *
 * Fire-and-forget: the caller does not need to await this.  Each resolved
 * photo triggers an individual React state update so cards light up one
 * by one instead of all at once after a long wait.
 *
 * @param client    MSGraphClientV3 with User.ReadBasic.All permission.
 * @param sponsors  The list returned by getSponsors / getSponsorsViaProxy.
 * @param onUpdate  Called per sponsor once photo data is ready.
 *                  Receives the sponsor ID plus the two (possibly undefined)
 *                  data-URL strings so the caller can do a targeted state update.
 */
export function loadPhotosProgressively(
  client: MSGraphClientV3,
  sponsors: ISponsor[],
  onUpdate: (sponsorId: string, photoUrl: string | undefined, managerPhotoUrl: string | undefined) => void
): void {
  const fetchOne = async (sponsor: ISponsor): Promise<void> => {
    let photoUrl: string | undefined;
    let managerPhotoUrl: string | undefined;

    // Fetch sponsor photo and manager photo concurrently.
    await Promise.all([
      (async () => {
        try {
          const buffer: ArrayBuffer = await client
            .api(`/users/${sponsor.id}/photo/$value`)
            .responseType(ResponseType.ARRAYBUFFER)
            .get();
          photoUrl = arrayBufferToDataUrl(buffer);
        } catch {
          // No photo — initials fallback stays.
        }
      })(),
      (async () => {
        if (!sponsor.managerId) return;
        try {
          const buffer: ArrayBuffer = await client
            .api(`/users/${sponsor.managerId}/photo/$value`)
            .responseType(ResponseType.ARRAYBUFFER)
            .get();
          managerPhotoUrl = arrayBufferToDataUrl(buffer);
        } catch {
          // No manager photo — initials fallback stays.
        }
      })(),
    ]);

    onUpdate(sponsor.id, photoUrl, managerPhotoUrl);
  };

  for (const sponsor of sponsors) {
    fetchOne(sponsor).catch(() => undefined);
  }
}
