import { MSGraphClientV3 } from '@microsoft/sp-http';
import { DisplayMode } from '@microsoft/sp-core-library';

export interface IGuestSponsorInfoProps {
  /** SPFx login name (UPN) of the current user – used as a fallback for guest detection. */
  loginName: string;
  /**
   * Direct flag from `pageContext.user.isExternalGuestUser` — the authoritative indicator
   * that the current user is an Entra B2B guest. More reliable than the `#EXT#` heuristic
   * when the SharePoint user profile has not yet been initialised (SP.UserProfile 500).
   */
  isExternalGuestUser: boolean;
  /** Current display mode supplied by the SPFx page context. */
  displayMode: DisplayMode;
  /** Initialised Graph client; undefined while onInit is pending. */
  graphClient: MSGraphClientV3 | undefined;
  /** Optional heading shown above the sponsor cards. */
  title: string;
  /**
   * When true, the web part behaves as if the current user is a guest and
   * renders fictitious sponsor cards from MockSponsorService instead of
   * making live Graph calls.
   */
  mockMode: boolean;
  /**
   * Entra ID tenant ID of the host tenant (where the sponsors live).
   * Used to generate Teams deep links that open in the guest-account context.
   */
  hostTenantId: string;
}

