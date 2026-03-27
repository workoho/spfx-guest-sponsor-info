// SPDX-FileCopyrightText: 2026 Workoho GmbH <https://workoho.com>
// SPDX-FileCopyrightText: 2026 Julian Pawlowski <https://github.com/jpawlowski>
// SPDX-License-Identifier: LicenseRef-PolyForm-Shield-1.0.0

import { AadHttpClient } from '@microsoft/sp-http';
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

/**
 * Returns true when the SPFx login name belongs to a Microsoft Entra guest account.
 * Guest UPNs always contain the "#EXT#" marker introduced by Entra external identity.
 */
export function isGuestUser(loginName: string): boolean {
  return loginName.indexOf('#EXT#') !== -1;
}

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
  clientVersion?: string,
  sponsorFilter?: 'any' | 'exchange' | 'teams',
  requireUserMailbox?: boolean
): Promise<ISponsorsResult> {
  // Append sponsor eligibility filter parameters as query string values.
  // The Azure Function validates them server-side before use.
  const url = new URL(proxyUrl);
  url.searchParams.set('sponsorFilter', sponsorFilter ?? 'teams');
  url.searchParams.set('requireUserMailbox', String(requireUserMailbox ?? true));

  const options = clientVersion
    ? { headers: { 'X-Client-Version': clientVersion } }
    : undefined;
  const response = await aadHttpClient.get(url.toString(), AadHttpClient.configurations.v1, options);
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
 * Progressively fetches manager photos for a list of sponsors via the Azure Function
 * photo proxy endpoint (`/api/getPhoto`), calling onUpdate for each sponsor as soon
 * as the manager photo arrives.  Errors are silently swallowed — the card falls back
 * to initials.
 *
 * Sponsor photos are already included in the initial `getSponsorsViaProxy` response
 * and do not need to be fetched separately.
 *
 * Fire-and-forget: the caller does not need to await this.  Each resolved photo
 * triggers an individual React state update so manager avatars light up one by one.
 *
 * @param photoUrl      - Full URL of the Azure Function `/api/getPhoto` endpoint.
 * @param aadHttpClient - Pre-acquired AAD HTTP client scoped to the Function App Registration.
 * @param presenceToken - Optional signed token from the last getGuestSponsors response.
 *                        When present, the function validates the userId without an extra Graph call.
 * @param sponsors      - The active sponsor list from getSponsorsViaProxy.
 * @param onUpdate      - Called per sponsor once the manager photo is ready (or undefined on failure).
 */
export function loadManagerPhotosViaProxy(
  photoUrl: string,
  aadHttpClient: AadHttpClient,
  presenceToken: string | undefined,
  sponsors: ISponsor[],
  onUpdate: (sponsorId: string, managerPhotoUrl: string | undefined) => void
): void {
  const fetchOne = async (sponsor: ISponsor): Promise<void> => {
    if (!sponsor.managerId) return;
    try {
      const params = new URLSearchParams({ userId: sponsor.managerId });
      const url = `${photoUrl}?${params.toString()}`;
      const headers: Record<string, string> = {};
      if (presenceToken) headers['X-Presence-Token'] = presenceToken;
      const response = await aadHttpClient.get(
        url,
        AadHttpClient.configurations.v1,
        Object.keys(headers).length > 0 ? { headers } : undefined
      );
      if (!response.ok) {
        onUpdate(sponsor.id, undefined);
        return;
      }
      const data = await response.json() as { photoUrl?: string };
      onUpdate(sponsor.id, data.photoUrl);
    } catch {
      onUpdate(sponsor.id, undefined);
    }
  };

  for (const sponsor of sponsors) {
    fetchOne(sponsor).catch(() => undefined);
  }
}

