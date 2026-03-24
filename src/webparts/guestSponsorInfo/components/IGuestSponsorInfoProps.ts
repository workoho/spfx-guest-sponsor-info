import { MSGraphClientV3, AadHttpClient } from '@microsoft/sp-http';
import { DisplayMode } from '@microsoft/sp-core-library';
import type { IReadonlyTheme } from '@microsoft/sp-component-base';

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
  /** Maximum number of sponsors to display on the live page (1–5). Default: 2. */
  maxSponsorCount: number;
  /** Number of mock sponsor cards to display in demo mode (1–5). Default: 2. */
  mockSponsorCount: number;
  /** Notification to simulate in demo mode. Default: 'none'. */
  mockSimulatedHint: 'none' | 'teamsAccessPending' | 'versionMismatch' | 'sponsorUnavailable' | 'noSponsors';
  /** Show the "Teams not set up yet" notice to guest users. Default: true. */
  showTeamsAccessPendingHint: boolean;
  /** Show the "Update available" notice when the web part and Azure Function versions differ. Default: true. */
  showVersionMismatchHint: boolean;
  /** Show the "Sponsor not available" notice when all assigned sponsors are inactive. Default: true. */
  showSponsorUnavailableHint: boolean;
  /** Show the "No sponsors found" notice when no sponsors are assigned. Default: true. */
  showNoSponsorsHint: boolean;
  /** Card layout mode: 'full' (136px tiles), 'compact' (horizontal rows), or 'auto' (switches based on count). */
  cardLayout: 'auto' | 'full' | 'compact';
  /** Minimum number of sponsors that triggers compact layout in 'auto' mode. Default: 3. */
  cardLayoutAutoThreshold: number;
  /**
   * Entra ID tenant ID of the host tenant (where the sponsors live).
   * Used to generate Teams deep links that open in the guest-account context.
   */
  hostTenantId: string;
  /** URL of the Azure Function proxy endpoint. Undefined when not configured. */
  functionUrl: string | undefined;
  /**
   * URL of the Azure Function presence endpoint (`/api/getPresence`).  Derived from
   * the same Function App base URL as `functionUrl`.  When set, presence refresh polls
   * call this endpoint (application permissions via Managed Identity) instead of Graph
   * directly (delegated permissions), ensuring reliable results for guest users.
   */
  presenceUrl: string | undefined;
  /**
   * URL of the Azure Function ping endpoint (`/api/ping`).  Derived from
   * the same Function App base URL as `functionUrl`.  Used in edit mode to
   * verify connectivity without triggering sponsor lookups or permission errors.
   */
  pingUrl: string | undefined;
  /** Client ID of the App Registration used by the Azure Function proxy. Undefined when not configured. */
  functionClientId: string | undefined;
  /** Pre-acquired AAD HTTP client scoped to the function App Registration. Undefined when not configured. */
  aadHttpClient: AadHttpClient | undefined;
  /** Show business phone numbers in the contact card. */
  showBusinessPhones: boolean;
  /** Show the mobile phone number in the contact card. */
  showMobilePhone: boolean;
  /** Show the work location field in the contact card. */
  showWorkLocation: boolean;
  /** Show the sponsor's city. Default: false. */
  showCity: boolean;
  /** Show the sponsor's country or region. Default: false. */
  showCountry: boolean;
  /** Show the sponsor's street address. Default: false. */
  showStreetAddress: boolean;
  /** Show the sponsor's postal code. Default: false. */
  showPostalCode: boolean;
  /** Show the sponsor's state or province. Default: false. */
  showState: boolean;
  /** Optional Azure Maps key used for inline map preview. */
  azureMapsSubscriptionKey: string | undefined;
  /** External map provider used for fallback links. */
  externalMapProvider: 'bing' | 'google' | 'apple' | 'openstreetmap' | 'here' | 'none';
  /** Show the manager section in the contact card. */
  showManager: boolean;
  /** Show the presence status indicator and label. Default: true. */
  showPresence: boolean;
  /** Show the sponsor's job title in the contact card. Default: true. */
  showSponsorJobTitle: boolean;
  /** Show the manager's job title in the contact card. Default: true. */
  showManagerJobTitle: boolean;
  /** Show the sponsor's department in the Organization section. Default: false. */
  showSponsorDepartment: boolean;
  /** Show the manager's department below the manager's job title. Default: false. */
  showManagerDepartment: boolean;
  /** Show the sponsor's profile photo. When false, only initials are shown. Default: true. */
  showSponsorPhoto: boolean;
  /** Show the manager's profile photo. When false, only initials are shown. Default: true. */
  showManagerPhoto: boolean;
  /** Use informal address ("du"/"tu") instead of formal ("Sie"/"vous"). Default: false. */
  useInformalAddress: boolean;
  /**
   * Version string of the web part (from the manifest).
   * Sent as X-Client-Version on proxy requests so the Azure Function can log
   * a warning when client and server versions differ.
   */
  clientVersion: string;
  /**
   * Called whenever the Azure Function proxy connectivity status changes (edit mode only).
   * Allows the web part class to reflect the status in the property pane without
   * holding the proxyStatus state in both the component and the web part.
   */
  onProxyStatusChange?: (status: 'checking' | 'ok' | 'error') => void;
  /**
   * Called when a version mismatch between the web part and Azure Function is detected or
   * cleared (edit mode only). Allows the web part class to surface the notice in the
   * property pane near the version number without keeping this state in two places.
   */
  onVersionMismatch?: (detected: boolean) => void;
  /**
   * Unique prefix derived from the SPFx web part instance ID.
   * Passed as `id` to every FluentProvider so multiple web part instances
   * on the same page do not produce conflicting Fluent UI style-tag IDs.
   */
  fluentProviderId: string;
  /**
   * SPFx host site theme supplied by the ThemeProvider service.
   * Passed to createV9Theme to produce a matching Fluent UI v9 theme for the
   * FluentProvider that wraps the entire web part component tree.
   * When undefined the FluentProvider falls back to webLightTheme.
   */
  theme?: IReadonlyTheme;
}

