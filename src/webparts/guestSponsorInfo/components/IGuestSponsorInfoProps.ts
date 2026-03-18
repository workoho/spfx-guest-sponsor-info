import { MSGraphClientV3 } from '@microsoft/sp-http';
import { DisplayMode } from '@microsoft/sp-core-library';

export interface IGuestSponsorInfoProps {
  /** SPFx login name (UPN) of the current user – used to detect guest accounts. */
  loginName: string;
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
   * AAD tenant ID of the host tenant (where the sponsors live).
   * Used to generate Teams deep links that open in the guest-account context.
   */
  hostTenantId: string;
}

