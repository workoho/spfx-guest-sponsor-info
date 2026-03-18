import { MSGraphClientV3 } from '@microsoft/sp-http';
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
 */
function arrayBufferToDataUrl(buffer: ArrayBuffer): string {
  const bytes = new Uint8Array(buffer);
  let binary = '';
  for (let i = 0; i < bytes.byteLength; i++) {
    binary += String.fromCharCode(bytes[i]);
  }
  return `data:image/jpeg;base64,${btoa(binary)}`;
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
 * Requires the Presence.Read.All delegated permission.
 * Returns a map of userId → availability string.
 * Silently returns an empty map on any error (presence is non-critical).
 */
async function fetchPresences(
  client: MSGraphClientV3,
  userIds: string[]
): Promise<Map<string, string>> {
  const map = new Map<string, string>();
  if (userIds.length === 0) return map;
  try {
    const response = await client
      .api('/communications/getPresencesByUserId')
      .post({ ids: userIds });
    if (response?.value) {
      for (const entry of response.value as Array<{ id: string; availability: string }>) {
        if (entry.id && entry.availability) {
          map.set(entry.id, entry.availability);
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
 * Fetches the manager of a user and their profile photo.
 * Requires User.ReadBasic.All (delegated).
 * Returns undefined fields when no manager is set or on any error.
 */
async function fetchManager(
  client: MSGraphClientV3,
  userId: string
): Promise<{ managerDisplayName?: string; managerJobTitle?: string; managerPhotoUrl?: string }> {
  try {
    const manager = await client
      .api(`/users/${userId}/manager`)
      .select('id,displayName,jobTitle')
      .get() as Record<string, unknown>;
    const managerId = manager.id as string | undefined;
    const managerDisplayName = (manager.displayName as string) || undefined;
    const managerJobTitle = (manager.jobTitle as string) || undefined;
    let managerPhotoUrl: string | undefined;
    if (managerId) {
      try {
        const buffer: ArrayBuffer = await client
          .api(`/users/${managerId}/photo/$value`)
          .responseType(ResponseType.ARRAYBUFFER)
          .get();
        managerPhotoUrl = arrayBufferToDataUrl(buffer);
      } catch {
        // No manager photo — initials fallback.
      }
    }
    return { managerDisplayName, managerJobTitle, managerPhotoUrl };
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
        const [exists, photoUrl, managerInfo] = await Promise.all([
          userExists(client, sponsor.id),
          (async (): Promise<string | undefined> => {
            try {
              const buffer: ArrayBuffer = await client
                .api(`/users/${sponsor.id}/photo/$value`)
                .responseType(ResponseType.ARRAYBUFFER)
                .get();
              return arrayBufferToDataUrl(buffer);
            } catch {
              // No photo available – initials fallback will be used.
              return undefined;
            }
          })(),
          fetchManager(client, sponsor.id),
        ]);
        return { sponsor: { ...sponsor, photoUrl, ...managerInfo }, exists };
      })
    ),
  ]);

  const activeSponsors = perSponsorResults
    .filter(r => r.exists)
    .map(r => ({ ...r.sponsor, presence: presenceMap.get(r.sponsor.id) }));
  const unavailableCount = perSponsorResults.filter(r => !r.exists).length;
  return { activeSponsors, unavailableCount };
}
