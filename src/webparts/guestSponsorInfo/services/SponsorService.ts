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
  guestHasTeamsAccess?: boolean;
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
 * Checks whether a user object still exists in the directory.
 *
 * Returns false only on an explicit HTTP 404.  Any other error (throttling,
 * transient network failure) is treated as "still exists" so that a temporary
 * Graph outage does not incorrectly hide a sponsor card.
 *
 * Note: a *disabled* account (accountEnabled === false) still returns 200 here
 * because reading that flag on other users' objects requires User.Read.All, which
 * we intentionally do not request (least-privilege).
 */
async function userExists(client: MSGraphClientV3, userId: string): Promise<boolean> {
  try {
    await client.api(`/users/${userId}`).select('id').get();
    return true;
  } catch (error) {
    if ((error as { statusCode?: number }).statusCode === 404) return false;
    return true;
  }
}

/**
 * Fetches the sponsors of the signed-in user via Microsoft Graph.
 * For each sponsor the function concurrently:
 *   1. Verifies the user object still exists (404 → unavailable).
 *   2. Fetches the profile photo (silent fallback to initials on any error).
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

/**
 * Fetches the manager of a user (name and job title only, no photo).
 * Photo is deferred to loadPhotosProgressively so the card renders immediately.
 * Requires User.ReadBasic.All (delegated).
 * Returns undefined fields when no manager is set or on any error.
 */
async function fetchManager(
  client: MSGraphClientV3,
  userId: string
): Promise<{ managerDisplayName?: string; managerJobTitle?: string; managerId?: string }> {
  try {
    const manager = await client
      .api(`/users/${userId}/manager`)
      .select('id,displayName,jobTitle')
      .get() as Record<string, unknown>;
    const managerId = manager.id as string | undefined;
    const managerDisplayName = (manager.displayName as string) || undefined;
    const managerJobTitle = (manager.jobTitle as string) || undefined;
    return { managerDisplayName, managerJobTitle, managerId };
  } catch {
    // No manager set (404) or permission error — non-critical.
    return {};
  }
}

export async function getSponsors(client: MSGraphClientV3): Promise<ISponsorsResult> {
  const response = await client
    .api('/me/sponsors')
    .select('id,displayName,mail,jobTitle,department,officeLocation,businessPhones,mobilePhone')
    .get();

  if (!response?.value) return { activeSponsors: [], unavailableCount: 0 };

  const items = response.value as Record<string, unknown>[];
  const candidates: ISponsor[] = items.map(item => ({
    id: item.id as string,
    displayName: (item.displayName as string) || '',
    mail: (item.mail as string) || undefined,
    jobTitle: (item.jobTitle as string) || undefined,
    department: (item.department as string) || undefined,
    officeLocation: (item.officeLocation as string) || undefined,
    businessPhones: (item.businessPhones as string[]) || [],
    mobilePhone: (item.mobilePhone as string) || undefined,
  }));

  const sponsorIds = candidates.map(s => s.id);

  // Fetch presence for all sponsors in one batched call, concurrently with per-sponsor work.
  const [presenceMap, perSponsorResults] = await Promise.all([
    fetchPresences(client, sponsorIds),
    Promise.all(
      candidates.map(async sponsor => {
        const [exists, managerInfo] = await Promise.all([
          userExists(client, sponsor.id),
          fetchManager(client, sponsor.id),
        ]);
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
  const unavailableCount = perSponsorResults.filter(r => !r.exists).length;
  return { activeSponsors, unavailableCount };
}

/**
 * Fetches sponsor data via the Azure Function proxy instead of calling Graph directly.
 * The proxy authenticates the caller via EasyAuth and calls Graph with application
 * permissions (User.Read.All, optionally Presence.Read.All) using its Managed Identity.
 *
 * @param proxyUrl     - Full URL of the Azure Function endpoint.
 * @param aadHttpClient - Pre-acquired AAD HTTP client scoped to the function App Registration.
 */
export async function getSponsorsViaProxy(
  proxyUrl: string,
  aadHttpClient: AadHttpClient
): Promise<ISponsorsResult> {
  const response = await aadHttpClient.get(proxyUrl, AadHttpClient.configurations.v1);
  if (!response.ok) {
    const err = new Error(`Proxy returned ${response.status}`);
    (err as { statusCode?: number }).statusCode = response.status;
    throw err;
  }
  return response.json() as Promise<ISponsorsResult>;
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
