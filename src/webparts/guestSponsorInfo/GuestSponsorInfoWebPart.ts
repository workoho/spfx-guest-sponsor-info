// SPDX-FileCopyrightText: 2026 Workoho GmbH <https://workoho.com>
// SPDX-FileCopyrightText: 2026 Julian Pawlowski <https://github.com/jpawlowski>
// SPDX-License-Identifier: LicenseRef-PolyForm-Shield-1.0.0

import * as React from 'react';
import * as ReactDom from 'react-dom';
import { Version } from '@microsoft/sp-core-library';
import {
  type IPropertyPaneConfiguration,
  type IPropertyPaneDropdownOption,
  type IPropertyPaneField,
  PropertyPaneHorizontalRule,
  PropertyPaneTextField,
  PropertyPaneCheckbox,
  PropertyPaneToggle,
  PropertyPaneDropdown,
  PropertyPaneSlider,
} from '@microsoft/sp-property-pane';
import { PropertyPaneCustomField } from '@microsoft/sp-property-pane/lib/propertyPaneFields/propertyPaneCustomField/PropertyPaneCustomField';
import { BaseClientSideWebPart } from '@microsoft/sp-webpart-base';
import { AadHttpClient, SPHttpClient } from '@microsoft/sp-http';
import { ThemeProvider, IReadonlyTheme, ThemeChangedEventArgs } from '@microsoft/sp-component-base';
import { FluentProvider, IdPrefixProvider, MessageBar, MessageBarBody, Tooltip, webLightTheme, webDarkTheme, type Theme } from '@fluentui/react-components';
import { createV9Theme } from '@fluentui/react-migration-v8-v9';
import { RendererProvider } from '@griffel/react';
import { InfoRegular } from '@fluentui/react-icons';

import * as strings from 'GuestSponsorInfoWebPartStrings';
import GuestSponsorInfo from './components/GuestSponsorInfo';
import { IGuestSponsorInfoProps } from './components/IGuestSponsorInfoProps';
import { isValidFunctionUrl, isValidGuid } from './utils/fieldValidation';
import { MapProvider, MapProviderConfig, getEffectiveMapProvider } from './utils/mapProviderUtils';
import workohoDefaultLogo from './assets/workoho-default-logo.svg';
import appIcon from './assets/app-icon-rounded.svg';
import { griffelRenderer } from './griffelRenderer';

/**
 * Builds a MapProviderConfig object from flat property-pane storage properties.
 */
function buildMapProviderConfig(props: IGuestSponsorInfoWebPartProps): MapProviderConfig {
  if (props.mapProviderConfigMode) {
    return {
      mode: props.mapProviderConfigMode,
      manualProvider: props.mapProviderConfigManualProvider ?? 'bing',
      iosProvider: props.mapProviderConfigIosProvider ?? 'apple',
      androidProvider: props.mapProviderConfigAndroidProvider ?? 'google',
      windowsProvider: props.mapProviderConfigWindowsProvider ?? 'bing',
      macosProvider: props.mapProviderConfigMacosProvider ?? 'apple',
      linuxProvider: props.mapProviderConfigLinuxProvider ?? 'openstreetmap',
    };
  }

  return {
    mode: 'auto',
    windowsProvider: 'bing',
    macosProvider: 'apple',
    linuxProvider: 'openstreetmap',
  };
}

export interface IGuestSponsorInfoWebPartProps {
  title: string;
  /** Show the title above the sponsor cards. Default: true. */
  showTitle?: boolean;
  /** Size of the web part title. Defaults to 'h3' (24 px). */
  titleSize: 'h2' | 'h3' | 'h4' | 'normal';
  mockMode: boolean;
  /**
   * Notification to simulate in demo mode.
   * Replaces the former mockTeamsUnavailable boolean with a dropdown selection.
   * Default: 'none'.
   */
  /** Maximum number of sponsors shown to visitors on the live page (1-5). Default: 2. */
  maxSponsorCount: number;
  mockSimulatedHint: 'none' | 'teamsAccessPending' | 'versionMismatch' | 'sponsorUnavailable' | 'noSponsors';
  /** Show the "Teams not set up yet" notice to guest users. Default: true. */
  showTeamsAccessPendingHint: boolean;
  /** Show the "Update available" notice when the web part and Azure Function versions differ. Default: true. */
  showVersionMismatchHint: boolean;
  /** Show the "Sponsor not available" notice when all assigned sponsors are inactive. Default: true. */
  showSponsorUnavailableHint: boolean;
  /** Show the "No sponsors found" notice when no sponsors are assigned. Default: true. */
  showNoSponsorsHint: boolean;
  /** Card layout: 'auto' switches to compact when sponsors reach the threshold. Default: 'auto'. */
  cardLayout: 'auto' | 'full' | 'compact';
  /** Number of sponsors at which 'auto' switches to compact layout. Default: 3. */
  cardLayoutAutoThreshold: number;
  functionUrl: string;
  functionClientId: string;
  /** Show business phone numbers in the contact card. Default: true. */
  showBusinessPhones: boolean;
  /** Show the mobile phone number in the contact card. Default: true. */
  showMobilePhone: boolean;
  /** Show the work location field in the contact card. Default: true. */
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
  /**
   * Sponsor license eligibility filter sent to the Azure Function.
   *   'any'      — at least one active license
   *   'exchange' — active Exchange Online plan
   *   'teams'    — active Teams plan (default)
   */
  sponsorFilter: 'any' | 'exchange' | 'teams';
  /**
   * When true (default), require the sponsor to have a user-type mailbox
   * (mailboxSettings.userPurpose ∈ {'user','linked'}).
   * When false, an active Exchange Online license serves as the mailbox proxy.
   */
  requireUserMailbox: boolean;
  /** Azure Maps subscription key used for inline map preview. */
  azureMapsSubscriptionKey: string;
  /** Map provider mode: 'auto' for OS-specific, 'manual' for unified. Default: 'auto'. */
  mapProviderConfigMode?: 'auto' | 'manual' | 'none';
  /** iOS map provider (auto mode). Default: 'apple'. */
  mapProviderConfigIosProvider?: MapProvider;
  /** Android map provider (auto mode). Default: 'google'. */
  mapProviderConfigAndroidProvider?: MapProvider;
  /** Windows map provider (auto mode). Default: 'bing'. */
  mapProviderConfigWindowsProvider?: MapProvider;
  /** macOS map provider (auto mode). Default: 'apple'. */
  mapProviderConfigMacosProvider?: MapProvider;
  /** Linux map provider (auto mode). Default: 'openstreetmap'. */
  mapProviderConfigLinuxProvider?: MapProvider;
  /** Manual map provider (manual mode). */
  mapProviderConfigManualProvider?: MapProvider;
  /** Show the manager section in the contact card. Default: true. */
  showManager: boolean;
  /** Show the presence status indicator and label. Default: true. */
  showPresence: boolean;
  /** Show the sponsor's profile photo. Default: true. */
  showSponsorPhoto: boolean;
  /** Show the manager's profile photo. Default: true. */
  showManagerPhoto: boolean;
  /** Show the sponsor's job title. Default: true. */
  showSponsorJobTitle: boolean;
  /** Show the manager's job title. Default: true. */
  showManagerJobTitle: boolean;
  /** Show the sponsor's department. Default: false. */
  showSponsorDepartment: boolean;
  /** Show the manager's department. Default: false. */
  showManagerDepartment: boolean;
  /** Use informal address ("du"/"tu") instead of formal ("Sie"/"vous"). Default: false. */
  useInformalAddress: boolean;
  /**
   * Tracks whether the first-run welcome dialog has been dismissed for this instance.
   * Stored as a web part property so all users and admins share the same flag —
   * the dialog disappears for everyone once any editor has clicked through it.
   * Resets automatically when the web part is removed and re-added (new instance).
   */
  welcomeSeen: boolean;
}

export default class GuestSponsorInfoWebPart extends BaseClientSideWebPart<IGuestSponsorInfoWebPartProps> {

  private _aadHttpClient: AadHttpClient | undefined;
  private _proxyStatus: 'checking' | 'ok' | 'error' = 'checking';
  private _versionMismatch = false;
  private _newVersionAvailable: { version: string; url: string } | false = false;
  private _githubCheckDone = false;
  private _diagnosticsCheckDone = false;
  private _diagnosticsResult:
    | {
        assetSource:
          | 'publicCdn'
          | 'privateCdn'
          | 'tenantCatalog'
          | 'siteCollectionCatalog'
          | 'unknown';
        /** Only populated when assetSource === 'tenantCatalog'. */
        tenantCatalogEveryoneAccess: 'ok' | 'denied' | 'unknown';
        siteEveryoneAccess: 'ok' | 'missing' | 'unknown';
      }
    | undefined;

  /**
   * sessionStorage key for the cached GitHub release check result.
   * Uses the manifest component ID as a namespace so the key is globally unique
   * within the shared sessionStorage of a SharePoint page and cannot collide
   * with keys written by other web parts or first-party SharePoint code.
   */
  private get _githubCacheKey(): string {
    return `${this.manifest.id}/github-release`;
  }
  /** Cache TTL in milliseconds (1 hour). */
  private static readonly _GITHUB_CACHE_TTL = 3_600_000;
  private _themeProvider: ThemeProvider | undefined;
  private _theme: IReadonlyTheme | undefined;
  /** Resolved Fluent v9 theme — derived from `_theme` via `createV9Theme`, falling back to `webLightTheme`. */
  private _v9Theme: Theme = webLightTheme;

  public render(): void {
    const element: React.ReactElement<IGuestSponsorInfoProps> = React.createElement(
      GuestSponsorInfo,
      {
        loginName: this.context.pageContext.user.loginName,
        isExternalGuestUser: this.context.pageContext.user.isExternalGuestUser,
        displayMode: this.displayMode,
        title: this.properties.title,
        showTitle: this.properties.showTitle ?? true,
        titleSize: this.properties.titleSize ?? 'h2',
        mockMode: this.properties.mockMode ?? false,
        maxSponsorCount: this.properties.maxSponsorCount ?? 2,
        mockSponsorCount: 5,
        mockSimulatedHint: this.properties.mockSimulatedHint ?? 'none',
        showTeamsAccessPendingHint: this.properties.showTeamsAccessPendingHint ?? true,
        showVersionMismatchHint: this.properties.showVersionMismatchHint ?? false,
        showSponsorUnavailableHint: this.properties.showSponsorUnavailableHint ?? true,
        showNoSponsorsHint: this.properties.showNoSponsorsHint ?? true,
        cardLayout: this.properties.cardLayout ?? 'auto',
        cardLayoutAutoThreshold: this.properties.cardLayoutAutoThreshold ?? 3,
        hostTenantId: this.context.pageContext.aadInfo.tenantId.toString(),
        functionUrl: this.properties.functionUrl
          ? `https://${this.properties.functionUrl.replace(/^https?:\/\//i, '').replace(/\/$/, '')}/api/getGuestSponsors`
          : undefined,
        presenceUrl: this.properties.functionUrl
          ? `https://${this.properties.functionUrl.replace(/^https?:\/\//i, '').replace(/\/$/, '')}/api/getPresence`
          : undefined,
        pingUrl: this.properties.functionUrl
          ? `https://${this.properties.functionUrl.replace(/^https?:\/\//i, '').replace(/\/$/, '')}/api/ping`
          : undefined,
        photoUrl: this.properties.functionUrl
          ? `https://${this.properties.functionUrl.replace(/^https?:\/\//i, '').replace(/\/$/, '')}/api/getPhoto`
          : undefined,
        functionClientId: this.properties.functionClientId || undefined,
        aadHttpClient: this._aadHttpClient,
        showBusinessPhones: this.properties.showBusinessPhones ?? true,
        showMobilePhone: this.properties.showMobilePhone ?? true,
        showWorkLocation: this.properties.showWorkLocation ?? true,
        showCity: this.properties.showCity ?? false,
        showCountry: this.properties.showCountry ?? false,
        showStreetAddress: this.properties.showStreetAddress ?? false,
        showPostalCode: this.properties.showPostalCode ?? false,
        showState: this.properties.showState ?? false,
        sponsorFilter: this.properties.sponsorFilter ?? 'teams',
        requireUserMailbox: this.properties.requireUserMailbox ?? true,
        azureMapsSubscriptionKey: this.properties.azureMapsSubscriptionKey || undefined,
        externalMapProvider: getEffectiveMapProvider(
          navigator.userAgent,
          buildMapProviderConfig(this.properties)
        ),
        showManager: this.properties.showManager ?? true,
        showPresence: this.properties.showPresence ?? true,
        showSponsorPhoto: this.properties.showSponsorPhoto ?? true,
        showManagerPhoto: this.properties.showManagerPhoto ?? true,
        showSponsorJobTitle: this.properties.showSponsorJobTitle ?? true,
        showManagerJobTitle: this.properties.showManagerJobTitle ?? true,
        showSponsorDepartment: this.properties.showSponsorDepartment ?? false,
        showManagerDepartment: this.properties.showManagerDepartment ?? false,
        useInformalAddress: this.properties.useInformalAddress ?? false,
        clientVersion: this.manifest.version,
        welcomeSeen: this.properties.welcomeSeen ?? false,
        onWelcomeComplete: ({ chosenPath, apiUrl, clientId }) => {
          if (chosenPath === 'api') {
            this.properties.welcomeSeen = true;
            // Strip protocol and trailing slash to match the stored URL format.
            this.properties.functionUrl = (apiUrl ?? '').replace(/^https?:\/\//i, '').replace(/\/$/, '');
            this.properties.functionClientId = clientId ?? '';
            this.properties.mockMode = false;
          } else if (chosenPath === 'demo') {
            this.properties.welcomeSeen = true;
            this.properties.mockMode = true;
          }
          // 'skip': welcomeSeen stays false → wizard reappears on next edit session.
          this.render();
          // Property pane is opened by onWelcomeFinish once the admin dismisses
          // the Done step, so the pane only appears after the wizard is gone.
        },
        onWelcomeFinish: () => {
          // "Let's go" clicked on the Done step — wizard is closing. Now open
          // the property pane so the admin can review and fine-tune all settings.
          // Deferred by one tick to let React finish unmounting the wizard first.
          setTimeout(() => this.context.propertyPane.open(), 0);
        },
        fluentProviderId: `gsi-${this.context.instanceId}`,
        onWelcomeSkip: () => {
          // Admin skipped via X or "Not now" — open the property pane so they
          // have a direct path to configure the web part manually.
          setTimeout(() => this.context.propertyPane.open(), 0);
        },
        onProxyStatusChange: (status: 'checking' | 'ok' | 'error') => {
          this._proxyStatus = status;
          if (this.context.propertyPane.isPropertyPaneOpen()) {
            this.context.propertyPane.refresh();
          }
        },
        onVersionMismatch: (detected: boolean) => {
          if (this._versionMismatch === detected) return;
          this._versionMismatch = detected;
          if (this.context.propertyPane.isPropertyPaneOpen()) {
            this.context.propertyPane.refresh();
          }
        },
        onTitleChange: (newTitle: string) => {
          this.properties.title = newTitle;
        },
        theme: this._theme,
      }
    );

    // Wrap in IdPrefixProvider so that every internal ID generated by Fluent v9
    // (FluentProvider CSS-variables container, Popover/Tooltip ARIA IDs, etc.)
    // gets a unique prefix. Without this, SharePoint Online's own FluentProvider
    // generates the same auto-incremented IDs and the CSS custom-properties
    // containers collide — overwriting theme tokens and causing missing backgrounds,
    // borders, and shadows.
    //
    // Also wrap in the scoped RendererProvider so that useWebPartStyles() inside
    // GuestSponsorInfo (called before its own inner RendererProvider JSX)
    // uses the salted renderer instead of the default one.
    this._renderReactTree(
      React.createElement(IdPrefixProvider, { value: `gsi-${this.context.instanceId}-` },
        React.createElement(RendererProvider, { renderer: griffelRenderer } as React.ComponentProps<typeof RendererProvider>, element)
      ),
      this.domElement
    );
  }

  protected async onInit(): Promise<void> {
    await super.onInit();
    // Pre-populate the title when the web part is first added to a page.
    // strings is resolved by SPFx to the editor's current UI culture, so this
    // produces a localized default (e.g. "My Sponsors" / "Meine Sponsoren") at
    // the moment of first drop without any extra runtime language detection.
    // Use strict undefined check — an empty string means the editor explicitly
    // cleared the title and wants it to remain empty.
    if (this.properties.title === undefined) {
      this.properties.title = strings.TitlePlaceholder;
    }
    // Consume the SPFx ThemeProvider service so the FluentProvider can receive the
    // host site's v8-style theme and convert it to a v9 theme via createV9Theme.
    this._themeProvider = this.context.serviceScope.consume(ThemeProvider.serviceKey);
    this._theme = this._themeProvider?.tryGetTheme();
    this._v9Theme = this._buildV9Theme(this._theme);
    this._themeProvider?.themeChangedEvent.add(this, this._handleThemeChanged);
    // Acquire Graph and AAD clients in the background — do NOT await them here.
    // SPFx awaits the onInit() Promise before rendering any web part on the page,
    // so blocking here would delay the entire page. Instead we resolve immediately
    // and re-render once clients are ready. The React component already shows a
    // shimmer while clients are undefined, so there is no visible flash.
    this._acquireClientsInBackground();
  }

  private _handleThemeChanged = (args: ThemeChangedEventArgs): void => {
    this._theme = args.theme;
    this._v9Theme = this._buildV9Theme(args.theme);
    this.render();
  };

  /** Creates a Fluent v9 theme from an SPFx v8 IReadonlyTheme, falling back to webLightTheme. */
  private _buildV9Theme(theme: IReadonlyTheme | undefined): Theme {
    if (!theme) return webLightTheme;
    return createV9Theme(
      theme as unknown as Parameters<typeof createV9Theme>[0],
      theme.isInverted ? webDarkTheme : webLightTheme
    );
  }

  /**
   * Queries the Azure Function's `/api/getLatestRelease` endpoint for the latest
   * published GitHub release version.  The function caches the result of its own
   * periodic GitHub API call in memory, so multiple web part clients share a
   * single outbound GitHub request per six-hour window without any browser → GitHub
   * traffic.  When the function is not configured this method is a no-op.
   *
   * Called lazily from onPropertyPaneOpened the first time the pane is opened;
   * errors are silently ignored because this is purely informational.
   */
  private _checkGitHubRelease(): void {
    const functionUrl = this.properties.functionUrl;
    const aadHttpClient = this._aadHttpClient;
    // No function configured or client not yet acquired — skip silently.
    if (!functionUrl || !aadHttpClient) return;

    const releaseCheckUrl =
      `https://${functionUrl.replace(/^https?:\/\//i, '').replace(/\/$/, '')}/api/getLatestRelease`;

    aadHttpClient
      .get(releaseCheckUrl, AadHttpClient.configurations.v1)
      .then((res): Promise<{ latestVersion: string | null; url: string | null } | null> =>
        res.ok
          ? (res.json() as Promise<{ latestVersion: string | null; url: string | null }>)
          : Promise.resolve(null)
      )
      .then((data) => {
        const latest = data?.latestVersion;
        const releaseUrl = data?.url;
        // Null values mean the timer hasn't completed its first run yet — don't cache;
        // the next pane open will try again.
        if (typeof latest !== 'string' || !latest) return;
        if (typeof releaseUrl !== 'string' || !releaseUrl) return;
        const current = this.manifest.version.split('.').slice(0, 3).join('.');
        const newer = this._isNewerVersion(latest, current);
        // Persist both outcomes so the next pane open (even after a view↔edit
        // switch) doesn't re-fetch within the cache TTL window.
        try {
          sessionStorage.setItem(
            this._githubCacheKey,
            JSON.stringify({
              version: newer ? latest : null,
              url: newer ? releaseUrl : null,
              ts: Date.now()
            })
          );
        } catch { /* sessionStorage unavailable — ignore */ }
        if (newer) {
          this._newVersionAvailable = { version: latest, url: releaseUrl };
          if (this.context.propertyPane.isPropertyPaneOpen()) {
            this.context.propertyPane.refresh();
          }
        }
      })
      .catch(() => { /* informational only — silent on failure */ });
  }

  private _isNewerVersion(candidate: string, current: string): boolean {
    const parse = (v: string): number[] => v.split('.').map(n => parseInt(n, 10) || 0);
    const a = parse(candidate);
    const b = parse(current);
    for (let i = 0; i < 3; i++) {
      if ((a[i] ?? 0) > (b[i] ?? 0)) return true;
      if ((a[i] ?? 0) < (b[i] ?? 0)) return false;
    }
    return false;
  }

  /**
   * Checks whether the web part's JavaScript bundle is accessible to B2B guest
   * users and whether the current site's Visitor group includes Everyone.
   *
   * Asset-source detection uses the SPFx manifest's `loaderConfig.internalModuleBaseUrls`
   * field which is populated at runtime by the SPFx loader and reliably indicates
   * where the bundle was actually fetched from.
   *
   * The Everyone Visitor check uses the SharePoint REST API (spHttpClient).
   * A 403 response on the Tenant App Catalog yields 'unknown' — gracefully handled.
   *
   * Results are stored in `_diagnosticsResult` and trigger a property pane refresh.
   */
  private async _runDiagnostics(): Promise<void> {
    // ── Step 1: Determine asset source (synchronous, no network) ─────────────
    const baseUrls: string[] = (this.manifest as unknown as {
      loaderConfig?: { internalModuleBaseUrls?: string[] };
    }).loaderConfig?.internalModuleBaseUrls ?? [];

    const legacyCtx = this.context.pageContext.legacyPageContext as Record<string, unknown>;
    const corporateCatalogUrl = (typeof legacyCtx?.corporateCatalogUrl === 'string'
      ? legacyCtx.corporateCatalogUrl
      : ''
    ).replace(/\/$/, '');
    const siteUrl = this.context.pageContext.site.absoluteUrl.replace(/\/$/, '');

    let assetSource: 'publicCdn' | 'privateCdn' | 'tenantCatalog' | 'siteCollectionCatalog' | 'unknown';
    if (baseUrls.some(u => u.includes('publiccdn.sharepointonline.com'))) {
      assetSource = 'publicCdn';
    } else if (baseUrls.some(u => u.includes('privatecdn.sharepointonline.com'))) {
      assetSource = 'privateCdn';
    } else if (corporateCatalogUrl && baseUrls.some(u => u.startsWith(corporateCatalogUrl))) {
      assetSource = 'tenantCatalog';
    } else if (baseUrls.some(u => u.startsWith(siteUrl))) {
      assetSource = 'siteCollectionCatalog';
    } else {
      assetSource = 'unknown';
    }

    // ── Step 2: Tenant App Catalog Everyone access (only when relevant) ───────
    let tenantCatalogEveryoneAccess: 'ok' | 'denied' | 'unknown' = 'unknown';
    if (assetSource === 'tenantCatalog' && corporateCatalogUrl) {
      try {
        const url = `${corporateCatalogUrl}/_api/web/roleassignments?$expand=Member&$select=Member/LoginName`;
        const resp = await this.context.spHttpClient.get(url, SPHttpClient.configurations.v1);
        if (resp.ok) {
          const data = await resp.json() as { value: Array<{ Member: { LoginName: string } }> };
          const hasEveryone = data.value.some(ra => ra.Member.LoginName === 'c:0(.s|true');
          tenantCatalogEveryoneAccess = hasEveryone ? 'ok' : 'denied';
        }
        // 403 or other non-ok = leave as 'unknown'
      } catch { /* network error — leave as 'unknown' */ }
    }

    // ── Step 3: Site Everyone access (always check) ───────────────────────────
    let siteEveryoneAccess: 'ok' | 'missing' | 'unknown' = 'unknown';
    try {
      const url = `${siteUrl}/_api/web/roleassignments?$expand=Member&$select=Member/LoginName`;
      const resp = await this.context.spHttpClient.get(url, SPHttpClient.configurations.v1);
      if (resp.ok) {
        const data = await resp.json() as { value: Array<{ Member: { LoginName: string } }> };
        const hasEveryone = data.value.some(ra => ra.Member.LoginName === 'c:0(.s|true');
        siteEveryoneAccess = hasEveryone ? 'ok' : 'missing';
      }
    } catch { /* network error — leave as 'unknown' */ }

    this._diagnosticsResult = { assetSource, tenantCatalogEveryoneAccess, siteEveryoneAccess };
    if (this.context.propertyPane.isPropertyPaneOpen()) {
      this.context.propertyPane.refresh();
    }
  }

  /**
   * Starts Graph and AAD HTTP client acquisition concurrently without blocking
   * the SPFx page lifecycle. Calls render() once both have settled so the React
   * component receives real clients in a single props update (avoids triggering
   * two separate data fetches if both clients resolve close together).
   * Both sub-promises catch their own errors and always resolve, so Promise.all
   * is safe to use here without needing Promise.allSettled (ES2020).
   */
  private _acquireClientsInBackground(): void {
    const clientId = this.properties.functionClientId;
    const aadPromise = clientId
      ? this.context.aadHttpClientFactory.getClient(clientId)
          .then(client => { this._aadHttpClient = client; })
          .catch(() => { /* AAD client unavailable — proxy path stays disabled */ })
      : Promise.resolve();

    // The trailing .catch() satisfies @typescript-eslint/no-floating-promises without the
    // void operator (which no-void disallows). In practice this catch is never reached.
    aadPromise.then(() => {
      this.render();
    }).catch(() => { /* unreachable */ });
  }

  protected onDispose(): void {
    this._themeProvider?.themeChangedEvent.remove(this, this._handleThemeChanged);
    this._unmountReactTree(this.domElement);
  }

  protected get dataVersion(): Version {
    return Version.parse('1.0');
  }

  protected onPropertyPaneOpened(): void {
    // Run the guest-accessibility diagnostics once per editing session.
    // Not gated on functionUrl — the check is about bundle delivery and site
    // permissions, which are independent of the proxy configuration.
    if (!this._diagnosticsCheckDone) {
      this._diagnosticsCheckDone = true;
      this._runDiagnostics().catch(() => { /* informational only — silent on failure */ });
    }

    if (this._githubCheckDone) return;
    // Only perform the release check when a function URL is configured.
    // Without a function the web part has no indirect path to GitHub, so we
    // skip the check entirely (no direct browser → GitHub API calls).
    // Leave _githubCheckDone = false so the check can still run later in the
    // same editing session if the admin configures a function URL.
    if (!this.properties.functionUrl) return;
    this._githubCheckDone = true;
    // Restore from sessionStorage cache so that view↔edit mode switches or page
    // saves (which re-instantiate the web part without a full browser reload)
    // don't trigger a redundant Function API call.
    try {
      const raw = sessionStorage.getItem(this._githubCacheKey);
      if (raw) {
        const { version, url, ts } = JSON.parse(raw) as
          { version: string | null; url: string | null; ts: number };
        if (Date.now() - ts < GuestSponsorInfoWebPart._GITHUB_CACHE_TTL) {
          if (version && url) {
            this._newVersionAvailable = { version, url };
            this.context.propertyPane.refresh();
          }
          return; // cache hit — skip fetch
        }
      }
    } catch { /* sessionStorage unavailable or corrupt — proceed with live fetch */ }
    this._checkGitHubRelease();
  }

  protected onPropertyPaneFieldChanged(propertyPath: string, oldValue: unknown, newValue: unknown): void {
    super.onPropertyPaneFieldChanged(propertyPath, oldValue, newValue);

    if (oldValue === newValue) return;
    if (
      propertyPath === 'functionUrl' ||
      propertyPath === 'functionClientId'
    ) {
      this._proxyStatus = 'checking';
      // Allow the release check to re-run after the function URL changes so the
      // update badge appears without requiring a full page reload.
      this._githubCheckDone = false;
      this.context.propertyPane.refresh();
    } else if (
      propertyPath === 'mockMode' ||
      propertyPath === 'cardLayout' ||
      propertyPath === 'showCity' ||
      propertyPath === 'showCountry' ||
      propertyPath === 'showWorkLocation' ||
      propertyPath === 'showStreetAddress' ||
      propertyPath === 'showPostalCode' ||
      propertyPath === 'showState' ||
      propertyPath === 'showManager' ||
      propertyPath === 'showTitle' ||
      propertyPath === 'mapProviderConfigMode'
    ) {
      this.context.propertyPane.refresh();
    }
  }

  private _getLocationDisplayHint(): string {
    if (!(strings as unknown as object | undefined)) return '';
    // Always return the location display hint, regardless of which address fields
    // are enabled. Admins need to understand how address fields work in general.
    return strings.LocationDisplayHintSeparateRows;
  }

  /** Thin wrapper so the file has exactly one `ReactDom.render` call site (ESLint pair-react-dom-render-unmount). */
  private _renderReactTree(tree: React.ReactElement, container: Element): void {
    ReactDom.render(tree, container);
  }

  /** Thin wrapper so the file has exactly one `ReactDom.unmountComponentAtNode` call site (ESLint pair-react-dom-render-unmount). */
  private _unmountReactTree(container: Element): void {
    ReactDom.unmountComponentAtNode(container);
  }

  /**
   * Renders a Fluent UI MessageBar info/warning box into a property pane custom
   * field container. Re-uses the same shared Griffel renderer and FluentProvider
   * setup as the proxy status field so styles are deduplicated correctly.
   */
  private _renderInfoBox(
    element: HTMLElement | undefined,
    key: string,
    text: string,
    intent: 'info' | 'warning' | 'success' | 'error' = 'info'
  ): void {
    if (!element) return;
    this._renderReactTree(
      React.createElement(IdPrefixProvider, { value: `gsi-pp-info-${key}-${this.context.instanceId}-` },
        React.createElement(RendererProvider, { renderer: griffelRenderer } as React.ComponentProps<typeof RendererProvider>,
          React.createElement(FluentProvider,
            {
              theme: this._theme
                ? createV9Theme(this._theme as unknown as Parameters<typeof createV9Theme>[0], this._theme.isInverted ? webDarkTheme : webLightTheme)
                : undefined,
              id: `gsi-pp-info-${key}-${this.context.instanceId}`,
            },
            React.createElement(MessageBar,
              { intent, style: { marginBottom: 8 } },
              React.createElement(MessageBarBody, null, text)
            )
          )
        )
      ),
      element
    );
  }

  /**
   * Renders a small styled sub-section divider (uppercase label + horizontal
   * rule) using SPFx theme colours for consistent appearance. DOM-based to
   * avoid the overhead of a full React tree for a purely presentational element.
   */
  private _renderSectionHeader(element: HTMLElement | undefined, label: string): void {
    if (!element) return;
    element.innerHTML = '';
    const divider = this._v9Theme.colorNeutralStroke2;
    const subtext = this._v9Theme.colorNeutralForeground3;

    const wrapper = document.createElement('div');
    wrapper.style.display = 'flex';
    wrapper.style.alignItems = 'center';
    wrapper.style.gap = '8px';
    wrapper.style.marginTop = '12px';
    wrapper.style.marginBottom = '4px';

    const labelEl = document.createElement('span');
    labelEl.textContent = label;
    labelEl.style.fontSize = '11px';
    labelEl.style.fontWeight = '600';
    labelEl.style.textTransform = 'uppercase';
    labelEl.style.letterSpacing = '0.05em';
    labelEl.style.color = subtext;
    labelEl.style.whiteSpace = 'nowrap';
    labelEl.style.flexShrink = '0';

    const line = document.createElement('div');
    line.style.flex = '1';
    line.style.height = '1px';
    line.style.backgroundColor = divider;

    wrapper.appendChild(labelEl);
    wrapper.appendChild(line);
    element.appendChild(wrapper);
  }

  /**
   * Renders a subtle helper description text (no MessageBar chrome) for
   * non-urgent explanatory copy in the property pane.
   */
  private _renderDescriptionText(element: HTMLElement | undefined, text: string): void {
    if (!element) return;
    element.innerHTML = '';

    const bodySubtext = this._v9Theme.colorNeutralForeground3;

    const description = document.createElement('div');
    description.textContent = text;
    description.style.fontSize = '12px';
    description.style.lineHeight = '1.45';
    description.style.color = bodySubtext;
    description.style.margin = '2px 0 8px 0';

    element.appendChild(description);
  }

  /** Factory: property pane custom field that renders a MessageBar info/warning box. */
  private _infoBoxField(
    key: string,
    text: string,
    intent: 'info' | 'warning' | 'success' | 'error' = 'info'
  ): IPropertyPaneField<unknown> {
    return PropertyPaneCustomField({
      key,
      onRender: (el: HTMLElement | undefined) => this._renderInfoBox(el, key, text, intent),
      onDispose: (el: HTMLElement | undefined) => { if (el) this._unmountReactTree(el); },
    }) as unknown as IPropertyPaneField<unknown>;
  }

  /** Factory: property pane custom field that renders subtle description text. */
  private _descriptionField(key: string, text: string): IPropertyPaneField<unknown> {
    return PropertyPaneCustomField({
      key,
      onRender: (el: HTMLElement | undefined) => this._renderDescriptionText(el, text),
      onDispose: (el: HTMLElement | undefined) => { if (el) el.innerHTML = ''; },
    }) as unknown as IPropertyPaneField<unknown>;
  }

  /** Factory: property pane custom field that renders a styled sub-section header. */
  private _sectionHeaderField(key: string, label: string): IPropertyPaneField<unknown> {
    return PropertyPaneCustomField({
      key,
      onRender: (el: HTMLElement | undefined) => this._renderSectionHeader(el, label),
      onDispose: (el: HTMLElement | undefined) => { if (el) el.innerHTML = ''; },
    }) as unknown as IPropertyPaneField<unknown>;
  }

  /** Factory: property pane custom field that renders a stable label + info icon row. */
  private _labelWithInlineInfoField(key: string, label: string, tooltipText: string): IPropertyPaneField<unknown> {
    return PropertyPaneCustomField({
      key,
      onRender: (el: HTMLElement | undefined) => this._renderLabelWithInlineInfo(el, key, label, tooltipText),
      onDispose: (el: HTMLElement | undefined) => { if (el) this._unmountReactTree(el); },
    }) as unknown as IPropertyPaneField<unknown>;
  }

  /** Renders a property-pane style label row with an inline info icon. */
  private _renderLabelWithInlineInfo(
    element: HTMLElement | undefined,
    key: string,
    label: string,
    tooltipText: string
  ): void {
    if (!element) return;
    element.innerHTML = '';

    const bodyText = this._v9Theme.colorNeutralForeground1;
    const bodySubtext = this._v9Theme.colorNeutralForeground3;

    const row = document.createElement('div');
    row.style.display = 'inline-flex';
    row.style.alignItems = 'center';
    row.style.gap = '6px';
    row.style.margin = '0 0 4px 0';

    const labelEl = document.createElement('span');
    labelEl.textContent = label;
    labelEl.style.fontSize = '14px';
    labelEl.style.fontWeight = '400';
    labelEl.style.color = bodyText;

    const iconHost = document.createElement('span');
    row.appendChild(labelEl);
    row.appendChild(iconHost);
    element.appendChild(row);

    this._renderReactTree(
      React.createElement(IdPrefixProvider, { value: `gsi-pp-label-inline-icon-${key}-${this.context.instanceId}-` },
        React.createElement(RendererProvider, { renderer: griffelRenderer } as React.ComponentProps<typeof RendererProvider>,
          React.createElement(FluentProvider,
            {
              theme: this._theme
                ? createV9Theme(this._theme as unknown as Parameters<typeof createV9Theme>[0], this._theme.isInverted ? webDarkTheme : webLightTheme)
                : undefined,
              id: `gsi-pp-label-inline-icon-${key}-${this.context.instanceId}`,
            },
            React.createElement(Tooltip,
            {
              content: tooltipText,
              relationship: 'label' as const,
            },
            React.createElement(InfoRegular, {
              style: {
                fontSize: '14px',
                color: bodySubtext,
                cursor: 'help',
              },
            })
          )
        )
      )
    ),
    iconHost
    );
  }

  private _renderAuthorSection(element?: HTMLElement): void {
    if (!element) return;

    element.innerHTML = '';

    // Read resolved colour tokens from the cached Fluent v9 theme.
    const t = this._v9Theme;
    const linkColor           = t.colorBrandForegroundLink;
    const bodySubtextColor    = t.colorNeutralForeground3;   // tertiary – used for labels/meta
    const themeLighterBg      = t.colorBrandBackground2;
    const themeLightBorder    = t.colorBrandStroke2;
    const themeDarkColor      = t.colorBrandForeground2;

    const createParagraph = (
      text: string,
      marginBottom: string = '8px',
      fontWeight: string = '400'
    ): HTMLParagraphElement => {
      const paragraph = document.createElement('p');
      paragraph.textContent = text;
      paragraph.style.margin = `0 0 ${marginBottom} 0`;
      paragraph.style.fontWeight = fontWeight;
      paragraph.style.lineHeight = '1.45';
      return paragraph;
    };

    const createCtaLink = (text: string, href: string): HTMLAnchorElement => {
      const ctaLink = document.createElement('a');
      ctaLink.textContent = `${text}\u00a0\u2192`;
      ctaLink.href = href;
      ctaLink.target = '_blank';
      ctaLink.rel = 'noopener noreferrer';
      ctaLink.style.display = 'inline-block';
      ctaLink.style.margin = '0 0 12px 0';
      ctaLink.style.padding = '3px 10px';
      ctaLink.style.color = linkColor;
      ctaLink.style.border = `1px solid ${linkColor}`;
      ctaLink.style.borderRadius = '4px';
      ctaLink.style.fontWeight = '600';
      ctaLink.style.fontSize = '13px';
      ctaLink.style.textDecoration = 'none';
      return ctaLink;
    };

    const createGlobeIcon = (): SVGSVGElement => {
      const svgNs = 'http://www.w3.org/2000/svg';
      const icon = document.createElementNS(svgNs, 'svg');
      icon.setAttribute('viewBox', '0 0 16 16');
      icon.setAttribute('width', '12');
      icon.setAttribute('height', '12');
      icon.setAttribute('aria-hidden', 'true');
      icon.style.verticalAlign = 'text-bottom';
      icon.style.marginRight = '4px';

      const path = document.createElementNS(svgNs, 'path');
      path.setAttribute(
        'd',
        'M8 0a8 8 0 1 0 0 16A8 8 0 0 0 8 0zM1.052 8.634a6.977 6.977 0 0 1-.04-' +
        '1.27c.24.073.49.135.75.184C1.86 9.222 2.38 10.74 3.14 11.9a6.98 6.98 0 0 1-' +
        '2.088-3.266zM8 14c-.68 0-1.47-.9-2.046-2.38-.17-.43-.31-.9-.425-1.389a14.517' +
        ' 14.517 0 0 0 2.47.209c.86 0 1.69-.073 2.472-.209a10.51 10.51 0 0 1-.426 1.389' +
        'C9.471 13.1 8.68 14 8 14zm-2.72-4.838A13.35 13.35 0 0 1 5.09 8c0-.388.02-.768' +
        '.058-1.137a15.65 15.65 0 0 0 2.853.267c1 0 1.963-.092 2.854-.267.038.369.058' +
        '.749.058 1.137 0 .405-.022.8-.063 1.175A15.63 15.63 0 0 1 8 9.42a15.63 15.63' +
        ' 0 0 1-2.72-.258zM12.86 11.9c.76-1.16 1.28-2.678 1.377-4.282.26-.049.51-.11' +
        '.75-.184a6.977 6.977 0 0 1-.04 1.27A6.98 6.98 0 0 1 12.86 11.9zM8 2c.68 0 1.47' +
        '.9 2.046 2.38.17.43.31.9.425 1.389A14.517 14.517 0 0 0 8 5.56a14.517 14.517 0' +
        ' 0 0-2.472.209A10.51 10.51 0 0 1 5.954 4.38C6.529 2.9 7.32 2 8 2zM3.14 4.1' +
        'A6.98 6.98 0 0 1 5.228 1.834c-.46.617-.854 1.393-1.16 2.265-.19.532-.341 1.1-' +
        '.452 1.693a7.16 7.16 0 0 1-.476-.184V5.6A6.99 6.99 0 0 1 3.14 4.1zm9.72 0a6.99' +
        ' 6.99 0 0 1 0 1.5v.008a7.16 7.16 0 0 1-.476.184 10.52 10.52 0 0 0-.452-1.693' +
        'c-.306-.872-.7-1.648-1.16-2.265A6.98 6.98 0 0 1 12.86 4.1z'
      );
      path.setAttribute('fill', 'currentColor');

      icon.appendChild(path);
      return icon;
    };

    // --- App info box: icon · name · meta · munich (centered) ---
    const appBox = document.createElement('div');
    appBox.style.display = 'flex';
    appBox.style.flexDirection = 'column';
    appBox.style.alignItems = 'center';
    appBox.style.textAlign = 'center';

    // App icon with link to the app homepage
    const appIconLink = document.createElement('a');
    appIconLink.href = 'https://guest-sponsor-info.workoho.cloud/?utm_source=guest-sponsor-info-webpart&utm_medium=sharepoint-webpart&utm_campaign=property-pane&utm_content=app-icon';
    appIconLink.target = '_blank';
    appIconLink.rel = 'noopener noreferrer';
    appIconLink.style.display = 'inline-block';
    appIconLink.style.marginBottom = '6px';
    appIconLink.style.textDecoration = 'none';

    const appIconImg = document.createElement('img');
    appIconImg.src = appIcon;
    appIconImg.alt = 'Guest Sponsor Info';
    appIconImg.style.display = 'block';
    appIconImg.style.width = '56px';
    appIconImg.style.height = '56px';
    appIconLink.appendChild(appIconImg);
    appBox.appendChild(appIconLink);

    // App name
    const appName = document.createElement('div');
    appName.textContent = 'Guest Sponsor Info';
    appName.style.fontWeight = '700';
    appName.style.fontSize = '15px';
    appName.style.marginBottom = '8px';
    appBox.appendChild(appName);

    // Meta line: source code · version [· update/new-release badges]
    const metaLine = document.createElement('div');
    metaLine.style.fontSize = '13px';
    metaLine.style.color = bodySubtextColor;
    metaLine.style.marginBottom = '6px';

    // semver is used below for the version link.
    const semver = this.manifest.version.split('.').slice(0, 3).join('.');

    const sourceLink = document.createElement('a');
    sourceLink.appendChild(createGlobeIcon());
    sourceLink.append(strings.AuthorSectionAppWebsiteLabel);
    sourceLink.href = 'https://guest-sponsor-info.workoho.cloud/?utm_source=guest-sponsor-info-webpart&utm_medium=sharepoint-webpart&utm_campaign=property-pane&utm_content=meta-website';
    sourceLink.target = '_blank';
    sourceLink.rel = 'noopener noreferrer';
    sourceLink.style.display = 'inline';
    sourceLink.style.color = linkColor;
    sourceLink.style.textDecoration = 'none';

    const separator = document.createElement('span');
    separator.textContent = ' · ';

    const versionLink = document.createElement('a');
    versionLink.textContent = `${strings.AuthorSectionVersionLabel} ${this.manifest.version}`;
    versionLink.href = `https://github.com/workoho/spfx-guest-sponsor-info/releases/tag/v${semver}`;
    versionLink.target = '_blank';
    versionLink.rel = 'noopener noreferrer';
    versionLink.style.color = linkColor;
    versionLink.style.textDecoration = 'none';

    metaLine.appendChild(sourceLink);
    metaLine.appendChild(separator);
    metaLine.appendChild(versionLink);

    if (this._versionMismatch) {
      const updateSep = document.createElement('span');
      updateSep.textContent = ' · ';
      const updateBadge = document.createElement('span');
      updateBadge.textContent = `\u2191\u00a0${strings.VersionMismatchTitle}`;
      updateBadge.title = strings.VersionMismatchMessage;
      updateBadge.style.display = 'inline-block';
      updateBadge.style.padding = '1px 7px';
      updateBadge.style.borderRadius = '4px';
      updateBadge.style.backgroundColor = themeLighterBg;
      updateBadge.style.color = themeDarkColor;
      updateBadge.style.border = `1px solid ${themeLightBorder}`;
      updateBadge.style.fontWeight = '700';
      updateBadge.style.fontSize = '12px';
      updateBadge.style.cursor = 'help';
      metaLine.appendChild(updateSep);
      metaLine.appendChild(updateBadge);
    }

    if (this._newVersionAvailable) {
      const newRelSep = document.createElement('span');
      newRelSep.textContent = ' \u00b7 ';
      const semverLatest = this._newVersionAvailable.version.split('.').slice(0, 3).join('.');
      const newRelBadge = document.createElement('a');
      newRelBadge.textContent = `\u2191\u00a0v${semverLatest}`;
      newRelBadge.href = this._newVersionAvailable.url + (this._newVersionAvailable.url.includes('?') ? '&' : '?') + 'utm_source=guest-sponsor-info-webpart&utm_medium=sharepoint-webpart&utm_campaign=property-pane&utm_content=update-badge';
      newRelBadge.target = '_blank';
      newRelBadge.rel = 'noopener noreferrer';
      newRelBadge.title = strings.NewReleaseAvailableLabel;
      newRelBadge.style.display = 'inline-block';
      newRelBadge.style.padding = '1px 7px';
      newRelBadge.style.borderRadius = '4px';
      newRelBadge.style.backgroundColor = t.colorPaletteGreenBackground1;
      newRelBadge.style.color = t.colorPaletteGreenForeground1;
      newRelBadge.style.border = `1px solid ${t.colorPaletteGreenBorder1}`;
      newRelBadge.style.fontWeight = '700';
      newRelBadge.style.fontSize = '12px';
      newRelBadge.style.textDecoration = 'none';
      metaLine.appendChild(newRelSep);
      metaLine.appendChild(newRelBadge);
    }

    appBox.appendChild(metaLine);

    const munichLine = document.createElement('span');
    munichLine.textContent = 'Built in Munich, World City with \u2764';
    munichLine.style.display = 'inline-block';
    munichLine.style.marginTop = '6px';
    munichLine.style.padding = '3px 10px';
    munichLine.style.borderRadius = '999px';
    munichLine.style.backgroundColor = themeLighterBg;
    munichLine.style.color = themeDarkColor;
    munichLine.style.fontWeight = '600';
    munichLine.style.fontSize = '12px';
    appBox.appendChild(munichLine);

    element.appendChild(appBox);

    // --- Divider ---
    const divider = document.createElement('div');
    divider.setAttribute('role', 'separator');
    divider.setAttribute('aria-hidden', 'true');
    divider.style.height = '1px';
    divider.style.background = `linear-gradient(to right, transparent, ${t.colorNeutralStroke1} 20%, ${t.colorNeutralStroke1} 80%, transparent)`;
    divider.style.margin = '20px 0';
    element.appendChild(divider);

    // --- Creator section ---
    const creatorSection = document.createElement('div');

    const createdByLabel = document.createElement('div');
    createdByLabel.textContent = strings.AuthorSectionCreatedBy;
    createdByLabel.style.fontSize = '13px';
    createdByLabel.style.color = t.colorNeutralForeground2;
    createdByLabel.style.textAlign = 'center';
    createdByLabel.style.marginBottom = '4px';
    creatorSection.appendChild(createdByLabel);

    const logoLink = document.createElement('a');
    logoLink.href = 'https://workoho.com/?utm_source=guest-sponsor-info-webpart&utm_medium=sharepoint-webpart&utm_campaign=property-pane&utm_content=author-logo';
    logoLink.target = '_blank';
    logoLink.rel = 'noopener noreferrer';
    logoLink.style.display = 'flex';
    logoLink.style.justifyContent = 'center';
    logoLink.style.marginBottom = '12px';
    logoLink.style.textDecoration = 'none';

    const logo = document.createElement('img');
    logo.src = workohoDefaultLogo;
    logo.alt = 'Workoho GmbH';
    logo.style.display = 'block';
    logo.style.width = 'min(150px, 50%)';
    logo.style.height = 'auto';
    logo.style.maxWidth = '150px';
    logo.style.aspectRatio = '4 / 1';
    logo.style.objectFit = 'contain';
    logo.style.opacity = '0.6';
    logoLink.appendChild(logo);
    creatorSection.appendChild(logoLink);

    const introPara = createParagraph(strings.AuthorSectionIntro, '8px', '700');
    introPara.style.color = t.colorNeutralForeground2;
    introPara.style.fontSize = '12px';
    creatorSection.appendChild(introPara);
    const consultingPara = createParagraph(strings.AuthorSectionConsultingText);
    consultingPara.style.color = t.colorNeutralForeground2;
    consultingPara.style.fontSize = '12px';
    creatorSection.appendChild(consultingPara);

    const ctaWrapper = document.createElement('div');
    ctaWrapper.style.textAlign = 'center';
    const ctaLink = createCtaLink(strings.AuthorSectionWebsiteLinkLabel, 'https://workoho.com/?utm_source=guest-sponsor-info-webpart&utm_medium=sharepoint-webpart&utm_campaign=property-pane&utm_content=author-cta');
    ctaLink.style.margin = '0';
    ctaWrapper.appendChild(ctaLink);
    creatorSection.appendChild(ctaWrapper);

    element.appendChild(creatorSection);
  }

  protected getPropertyPaneConfiguration(): IPropertyPaneConfiguration {
    // Guard: SPFx AMD locale bundles load asynchronously. If the property pane is opened
    // before the bundle resolves (can happen during rapid initialisation), return an empty
    // config so the framework doesn't crash. The pane will re-render once strings are ready.
    if (!(strings as unknown as object | undefined)) return { pages: [] };

    const showManager = this.properties.showManager ?? true;
    const mapProviderConfig = buildMapProviderConfig(this.properties);
    const mapProviderMode = mapProviderConfig.mode;
    const locationDisplayHint = this._getLocationDisplayHint();
    const hasAddressFieldsEnabled =
      (this.properties.showStreetAddress ?? false) ||
      (this.properties.showPostalCode ?? false) ||
      (this.properties.showCity ?? false) ||
      (this.properties.showState ?? false) ||
      (this.properties.showCountry ?? false);

    // Provider options for each section
    const windowsOptions: IPropertyPaneDropdownOption[] = [
      { key: 'none', text: strings.MapProviderNoneOption },
      { key: 'bing', text: strings.MapProviderBingOption },
      { key: 'google', text: strings.MapProviderGoogleOption },
      { key: 'apple', text: strings.MapProviderAppleOption },
      { key: 'openstreetmap', text: strings.MapProviderOpenStreetMapOption },
    ];
    const macosOptions: IPropertyPaneDropdownOption[] = [
      { key: 'none', text: strings.MapProviderNoneOption },
      { key: 'bing', text: strings.MapProviderBingOption },
      { key: 'google', text: strings.MapProviderGoogleOption },
      { key: 'apple', text: strings.MapProviderAppleOption },
      { key: 'openstreetmap', text: strings.MapProviderOpenStreetMapOption },
    ];

    const linuxOptions: IPropertyPaneDropdownOption[] = [
      { key: 'none', text: strings.MapProviderNoneOption },
      { key: 'bing', text: strings.MapProviderBingOption },
      { key: 'google', text: strings.MapProviderGoogleOption },
      { key: 'apple', text: strings.MapProviderAppleOption },
      { key: 'openstreetmap', text: strings.MapProviderOpenStreetMapOption },
    ];

    const allProviderOptions: IPropertyPaneDropdownOption[] = [
      { key: 'none', text: strings.MapProviderNoneOption },
      { key: 'bing', text: strings.MapProviderBingOption },
      { key: 'google', text: strings.MapProviderGoogleOption },
      { key: 'apple', text: strings.MapProviderAppleOption },
      { key: 'openstreetmap', text: strings.MapProviderOpenStreetMapOption },
    ];

    return {
      pages: [
        {
          displayGroupsAsAccordion: true,
          groups: [
            {
              groupName: strings.BasicGroupName,
              isCollapsed: false,
              groupFields: [
                this._sectionHeaderField('settingsLivePage', strings.PpSectionLivePage),
                this._labelWithInlineInfoField('maxSponsorCountLabelWithInfo', strings.MaxSponsorCountFieldLabel, strings.MaxSponsorCountFieldTooltip),
                PropertyPaneSlider('maxSponsorCount', {
                  label: '',
                  min: 1,
                  max: 5,
                  step: 1,
                  value: this.properties.maxSponsorCount ?? 2,
                }),
                this._sectionHeaderField('settingsDemoMode', strings.PpSectionDemoMode),
                this._labelWithInlineInfoField('mockModeLabelWithInfo', strings.MockModeFieldLabel, strings.MockModeFieldTooltip),
                PropertyPaneToggle('mockMode', {
                  label: ''
                }),
                PropertyPaneDropdown('mockSimulatedHint', {
                  label: strings.MockSimulatedHintFieldLabel,
                  options: [
                    { key: 'none', text: strings.MockSimulatedHintNoneOption },
                    { key: 'teamsAccessPending', text: strings.MockSimulatedHintTeamsAccessPendingOption },
                    { key: 'versionMismatch', text: strings.MockSimulatedHintVersionMismatchOption },
                    { key: 'sponsorUnavailable', text: strings.MockSimulatedHintSponsorUnavailableOption },
                    { key: 'noSponsors', text: strings.MockSimulatedHintNoSponsorsOption },
                  ],
                  selectedKey: this.properties.mockSimulatedHint ?? 'none',
                }),
              ]
            },
            {
              groupName: strings.SponsorEligibilityGroupName,
              isCollapsed: true,
              groupFields: [
                this._descriptionField('sponsorEligibilityGroupHint', strings.SponsorEligibilityGroupHint),
                PropertyPaneDropdown('sponsorFilter', {
                  label: strings.SponsorFilterFieldLabel,
                  options: [
                    { key: 'teams',    text: strings.SponsorFilterTeamsOption },
                    { key: 'exchange', text: strings.SponsorFilterExchangeOption },
                    { key: 'any',      text: strings.SponsorFilterAnyOption },
                  ],
                  selectedKey: this.properties.sponsorFilter ?? 'teams',
                }),
                PropertyPaneCheckbox('requireUserMailbox', {
                  text: strings.RequireUserMailboxFieldLabel,
                  checked: this.properties.requireUserMailbox ?? true,
                }),
              ]
            },
            {
              groupName: strings.GuestNotificationsGroupName,
              isCollapsed: true,
              groupFields: [
                this._descriptionField('guestNotificationsGroupHint', strings.GuestNotificationsGroupHint),
                PropertyPaneCheckbox('showTeamsAccessPendingHint', {
                  text: strings.ShowTeamsAccessPendingHintLabel,
                  checked: this.properties.showTeamsAccessPendingHint ?? true,
                }),
                PropertyPaneCheckbox('showVersionMismatchHint', {
                  text: strings.ShowVersionMismatchHintLabel,
                  checked: this.properties.showVersionMismatchHint ?? false,
                }),
                PropertyPaneCheckbox('showSponsorUnavailableHint', {
                  text: strings.ShowSponsorUnavailableHintLabel,
                  checked: this.properties.showSponsorUnavailableHint ?? true,
                }),
                PropertyPaneCheckbox('showNoSponsorsHint', {
                  text: strings.ShowNoSponsorsHintLabel,
                  checked: this.properties.showNoSponsorsHint ?? true,
                }),
                PropertyPaneHorizontalRule(),
                PropertyPaneCheckbox('useInformalAddress', {
                  text: strings.UseInformalAddressFieldLabel,
                }),
              ]
            },
            {
              // Card display: what fields appear on the sponsor card itself.
              groupName: strings.DisplayGroupName,
              isCollapsed: true,
              groupFields: [
                PropertyPaneToggle('showTitle', {
                  label: strings.ShowTitleFieldLabel,
                  checked: this.properties.showTitle ?? true,
                }),
                ...((this.properties.showTitle ?? true) ? [
                  PropertyPaneDropdown('titleSize', {
                    label: strings.TitleSizeFieldLabel,
                    options: [
                      { key: 'h2', text: strings.TitleSizeH2Option },
                      { key: 'h3', text: strings.TitleSizeH3Option },
                      { key: 'h4', text: strings.TitleSizeH4Option },
                      { key: 'normal', text: strings.TitleSizeNormalOption },
                    ],
                    selectedKey: this.properties.titleSize ?? 'h2',
                  }),
                ] : []),
                this._sectionHeaderField('displayCardLayout', strings.PpSectionCardLayout),
                PropertyPaneDropdown('cardLayout', {
                  label: strings.CardLayoutFieldLabel,
                  options: [
                    { key: 'auto', text: strings.CardLayoutAutoOption },
                    { key: 'full', text: strings.CardLayoutFullOption },
                    { key: 'compact', text: strings.CardLayoutCompactOption },
                  ],
                  selectedKey: this.properties.cardLayout ?? 'auto',
                }),
                ...((this.properties.cardLayout ?? 'auto') === 'auto' ? [
                  PropertyPaneSlider('cardLayoutAutoThreshold', {
                    label: strings.CardLayoutAutoThresholdFieldLabel,
                    min: 1,
                    max: 5,
                    step: 1,
                    value: this.properties.cardLayoutAutoThreshold ?? 3,
                  }),
                ] : []),
                this._sectionHeaderField('displayProfileFields', strings.PpSectionProfileFields),
                PropertyPaneToggle('showPresence', {
                  label: strings.ShowPresenceFieldLabel,
                  checked: this.properties.showPresence ?? true,
                }),
                PropertyPaneToggle('showSponsorPhoto', {
                  label: strings.ShowSponsorPhotoFieldLabel,
                  checked: this.properties.showSponsorPhoto ?? true,
                }),
                PropertyPaneCheckbox('showSponsorJobTitle', {
                  text: strings.ShowSponsorJobTitleFieldLabel,
                  checked: this.properties.showSponsorJobTitle ?? true,
                }),
                PropertyPaneCheckbox('showSponsorDepartment', {
                  text: strings.ShowSponsorDepartmentFieldLabel,
                }),
              ]
            },
            {
              groupName: strings.ContactInfoSection,
              isCollapsed: true,
              groupFields: [
                // Phone numbers
                this._sectionHeaderField('contactPhone', strings.PpSectionPhone),
                PropertyPaneToggle('showBusinessPhones', {
                  label: strings.ShowBusinessPhonesFieldLabel,
                  checked: this.properties.showBusinessPhones ?? true,
                }),
                PropertyPaneToggle('showMobilePhone', {
                  label: strings.ShowMobilePhoneFieldLabel,
                  checked: this.properties.showMobilePhone ?? true,
                }),
                // Work location (officeLocation)
                this._sectionHeaderField('contactWorkLocation', strings.PpSectionWorkLocation),
                PropertyPaneToggle('showWorkLocation', {
                  label: strings.ShowWorkLocationFieldLabel,
                  checked: this.properties.showWorkLocation ?? true,
                }),
                // Address fields (combined into a single clickable address row)
                this._sectionHeaderField('contactAddress', strings.PpSectionAddress),
                ...(locationDisplayHint
                  ? [this._descriptionField('locationDisplayHint', locationDisplayHint)]
                  : []),
                PropertyPaneCheckbox('showStreetAddress', {
                  text: strings.ShowStreetAddressFieldLabel,
                }),
                PropertyPaneCheckbox('showPostalCode', {
                  text: strings.ShowPostalCodeFieldLabel,
                }),
                PropertyPaneCheckbox('showCity', {
                  text: strings.ShowCityFieldLabel,
                }),
                PropertyPaneCheckbox('showState', {
                  text: strings.ShowStateFieldLabel,
                }),
                PropertyPaneCheckbox('showCountry', {
                  text: strings.ShowCountryFieldLabel,
                }),
                ...(hasAddressFieldsEnabled ? [
                  // Map settings are only relevant when at least one postal
                  // address field is shown in the contact section.
                  this._sectionHeaderField('contactMapLink', strings.PpSectionMapLink),
                  this._descriptionField('mapProviderModeHint', strings.MapProviderModeHint),
                  this._descriptionField('addressMapProviderHint', strings.AddressMapProviderHint),
                  PropertyPaneDropdown('mapProviderConfigMode', {
                    label: strings.ExternalMapProviderFieldLabel,
                    options: [
                      { key: 'auto', text: strings.MapProviderModeAutoLabel },
                      { key: 'manual', text: strings.MapProviderModeManualLabel },
                      { key: 'none', text: strings.MapProviderNoneOption },
                    ],
                    selectedKey: mapProviderMode,
                  }),
                  ...(mapProviderMode === 'auto' ? [
                    // Auto mode: show OS-specific selectors
                    PropertyPaneDropdown('mapProviderConfigIosProvider', {
                      label: strings.MapProviderIosLabel,
                      options: allProviderOptions,
                      selectedKey: mapProviderConfig.iosProvider ?? 'apple',
                    }),
                    PropertyPaneDropdown('mapProviderConfigAndroidProvider', {
                      label: strings.MapProviderAndroidLabel,
                      options: allProviderOptions,
                      selectedKey: mapProviderConfig.androidProvider ?? 'google',
                    }),
                    PropertyPaneDropdown('mapProviderConfigWindowsProvider', {
                      label: strings.MapProviderWindowsLabel,
                      options: windowsOptions,
                      selectedKey: mapProviderConfig.windowsProvider ?? 'bing',
                    }),
                    PropertyPaneDropdown('mapProviderConfigMacosProvider', {
                      label: strings.MapProviderMacOSLabel,
                      options: macosOptions,
                      selectedKey: mapProviderConfig.macosProvider ?? 'apple',
                    }),
                    PropertyPaneDropdown('mapProviderConfigLinuxProvider', {
                      label: strings.MapProviderLinuxLabel,
                      options: linuxOptions,
                      selectedKey: mapProviderConfig.linuxProvider ?? 'openstreetmap',
                    }),
                  ] : mapProviderMode === 'manual' ? [
                    // Manual mode: single provider for all
                    PropertyPaneDropdown('mapProviderConfigManualProvider', {
                      label: strings.ExternalMapProviderFieldLabel,
                      options: allProviderOptions,
                      selectedKey: mapProviderConfig.manualProvider ?? 'bing',
                    }),
                  ] : [
                    // None mode: all map links disabled — no sub-dropdowns needed
                  ]),
                  PropertyPaneHorizontalRule(),
                  // Azure Maps key: optional inline preview image
                  this._sectionHeaderField('contactMapPreview', strings.PpSectionMapPreview),
                  this._descriptionField('azureMapsPreviewHint', strings.AzureMapsPreviewHint),
                  PropertyPaneTextField('azureMapsSubscriptionKey', {
                    label: strings.AzureMapsSubscriptionKeyFieldLabel,
                  }),
                ] : []),
              ]
            },
            {
              groupName: strings.OrganizationSection,
              isCollapsed: true,
              groupFields: [
                PropertyPaneToggle('showManager', {
                  label: strings.ShowManagerFieldLabel,
                  checked: showManager,
                }),
                ...(showManager ? [
                  PropertyPaneCheckbox('showManagerJobTitle', {
                    text: strings.ShowManagerJobTitleFieldLabel,
                    checked: this.properties.showManagerJobTitle ?? true,
                  }),
                  PropertyPaneCheckbox('showManagerDepartment', {
                    text: strings.ShowManagerDepartmentFieldLabel,
                  }),
                  PropertyPaneCheckbox('showManagerPhoto', {
                    text: strings.ShowManagerPhotoFieldLabel,
                    checked: this.properties.showManagerPhoto ?? true,
                  }),
                ] : []),
              ]
            },
            {
              groupName: strings.FunctionGroupName,
              isCollapsed: true,
              groupFields: [
                PropertyPaneTextField('functionUrl', {
                  label: strings.FunctionUrlFieldLabel
                }),
                // Validation error — shown immediately below the URL field when
                // the stored value fails the format check.
                ...(this.properties.functionUrl && !isValidFunctionUrl(this.properties.functionUrl)
                  ? [this._infoBoxField('functionUrlValidationError', strings.InvalidUrlFormat, 'error')]
                  : []),
                // Inline hint — hoverable info icon that explains where to find
                // the Function App URL in the Azure portal.
                PropertyPaneCustomField({
                  key: 'functionUrlHintField',
                  onRender: (element: HTMLElement | undefined) => {
                    if (!element) return;
                    element.innerHTML = '';

                    const linkColor = this._v9Theme.colorBrandForegroundLink;
                    const subtextColor = this._v9Theme.colorNeutralForeground3;
                    const bgColor = this._v9Theme.colorNeutralBackground1;
                    const borderColor = this._v9Theme.colorNeutralStroke2;

                    const row = document.createElement('div');
                    row.style.position = 'relative';
                    row.style.display = 'inline-flex';
                    row.style.alignItems = 'center';
                    row.style.gap = '4px';
                    row.style.cursor = 'help';

                    const icon = document.createElement('span');
                    icon.textContent = '\u24d8';
                    icon.style.fontSize = '13px';
                    icon.style.color = linkColor;

                    const label = document.createElement('span');
                    label.textContent = strings.WelcomeDialogUrlHintLabel;
                    label.style.fontSize = '11px';
                    label.style.color = subtextColor;

                    const popup = document.createElement('div');
                    popup.style.display = 'none';
                    popup.style.position = 'absolute';
                    popup.style.left = '0';
                    popup.style.top = '100%';
                    popup.style.padding = '12px 10px 8px';
                    popup.style.backgroundColor = bgColor;
                    popup.style.border = `1px solid ${borderColor}`;
                    popup.style.borderRadius = '4px';
                    popup.style.boxShadow = this._v9Theme.shadow8;
                    popup.style.zIndex = '10';
                    popup.style.minWidth = '200px';
                    popup.style.maxWidth = '280px';
                    popup.style.lineHeight = '1.5';

                    const desc = document.createElement('div');
                    desc.style.fontSize = '11px';
                    desc.style.color = subtextColor;
                    desc.textContent = strings.WelcomeDialogUrlHintBody;
                    popup.appendChild(desc);

                    row.addEventListener('mouseenter', () => { popup.style.display = 'block'; });
                    row.addEventListener('mouseleave', () => { popup.style.display = 'none'; });

                    row.appendChild(icon);
                    row.appendChild(label);
                    row.appendChild(popup);
                    element.appendChild(row);
                  },
                  onDispose: (element: HTMLElement | undefined) => {
                    if (element) element.innerHTML = '';
                  },
                }) as unknown as IPropertyPaneField<unknown>,
                PropertyPaneTextField('functionClientId', {
                  label: strings.FunctionClientIdFieldLabel
                }),
                // Validation error — shown immediately below the Client ID field
                // when the stored value is not a valid GUID.
                ...(this.properties.functionClientId && !isValidGuid(this.properties.functionClientId)
                  ? [this._infoBoxField('functionClientIdValidationError', strings.InvalidGuidFormat, 'error')]
                  : []),
                // Inline Entra hint — hoverable info icon next to a short label
                // that reveals the exact App Registration name + a direct Entra
                // portal link, so admins can find the Client ID easily.
                PropertyPaneCustomField({
                  key: 'clientIdEntraHintField',
                  onRender: (element: HTMLElement | undefined) => {
                    if (!element) return;
                    element.innerHTML = '';

                    const linkColor = this._v9Theme.colorBrandForegroundLink;
                    const subtextColor = this._v9Theme.colorNeutralForeground3;
                    const bgColor = this._v9Theme.colorNeutralBackground1;
                    const borderColor = this._v9Theme.colorNeutralStroke2;

                    // Wrapper row: info icon + label
                    const row = document.createElement('div');
                    row.style.position = 'relative';
                    row.style.display = 'inline-flex';
                    row.style.alignItems = 'center';
                    row.style.gap = '4px';
                    row.style.cursor = 'help';

                    // ⓘ circle icon (Unicode)
                    const icon = document.createElement('span');
                    icon.textContent = '\u24d8';
                    icon.style.fontSize = '13px';
                    icon.style.color = linkColor;

                    // Short trigger label
                    const label = document.createElement('span');
                    label.textContent = strings.PpClientIdHintLabel;
                    label.style.fontSize = '11px';
                    label.style.color = subtextColor;

                    // Popup — hidden by default, shown on hover
                    const popup = document.createElement('div');
                    popup.style.display = 'none';
                    popup.style.position = 'absolute';
                    popup.style.left = '0';
                    popup.style.top = '100%';
                    popup.style.padding = '12px 10px 8px';
                    popup.style.backgroundColor = bgColor;
                    popup.style.border = `1px solid ${borderColor}`;
                    popup.style.borderRadius = '4px';
                    popup.style.boxShadow = this._v9Theme.shadow8;
                    popup.style.zIndex = '10';
                    popup.style.minWidth = '200px';
                    popup.style.maxWidth = '280px';
                    popup.style.lineHeight = '1.5';

                    // Popup content
                    const desc = document.createElement('div');
                    desc.style.fontSize = '11px';
                    desc.style.color = subtextColor;
                    desc.textContent = strings.PpClientIdHintBody;
                    popup.appendChild(desc);

                    const appName = document.createElement('div');
                    appName.style.fontSize = '11px';
                    appName.style.fontWeight = '600';
                    appName.style.color = subtextColor;
                    appName.style.marginTop = '4px';
                    appName.textContent = '“Guest Sponsor Info - SharePoint Web Part Auth”';
                    popup.appendChild(appName);

                    const entraLink = document.createElement('a');
                    entraLink.textContent = strings.FunctionEntraLinkLabel;
                    entraLink.href =
                      'https://entra.microsoft.com/#view/Microsoft_AAD_RegisteredApps/ApplicationsListBlade';
                    entraLink.target = '_blank';
                    entraLink.rel = 'noopener noreferrer';
                    entraLink.style.display = 'block';
                    entraLink.style.marginTop = '4px';
                    entraLink.style.fontSize = '11px';
                    entraLink.style.color = linkColor;
                    entraLink.style.textDecoration = 'none';
                    popup.appendChild(entraLink);

                    // Show/hide on hover
                    row.addEventListener('mouseenter', () => { popup.style.display = 'block'; });
                    row.addEventListener('mouseleave', () => { popup.style.display = 'none'; });

                    row.appendChild(icon);
                    row.appendChild(label);
                    row.appendChild(popup);
                    element.appendChild(row);
                  },
                  onDispose: (element: HTMLElement | undefined) => {
                    if (element) element.innerHTML = '';
                  },
                }) as unknown as IPropertyPaneField<unknown>,
                PropertyPaneCustomField({
                  key: 'functionSetupLinksField',
                  onRender: (element: HTMLElement | undefined) => {
                    if (!element) return;
                    element.innerHTML = '';

                    const linkColor = this._v9Theme.colorBrandForegroundLink;
                    const subtextColor = this._v9Theme.colorNeutralForeground3;

                    const setupGuideHref =
                      'https://guest-sponsor-info.workoho.cloud/setup';

                    const isConfigured = !!(this.properties.functionUrl && this.properties.functionClientId);
                    const isOk = isConfigured && this._proxyStatus === 'ok';
                    const isError = isConfigured && this._proxyStatus === 'error';

                    const wrapper = document.createElement('div');
                    wrapper.style.marginTop = '8px';
                    wrapper.style.display = 'flex';
                    wrapper.style.flexDirection = 'column';
                    wrapper.style.alignItems = 'flex-start';
                    wrapper.style.gap = '4px';

                    // Show the 3-step overview when unconfigured or on error
                    if (!isOk) {
                      const stepsIntro = document.createElement('div');
                      stepsIntro.style.fontSize = '12px';
                      stepsIntro.style.lineHeight = '1.45';
                      stepsIntro.style.color = subtextColor;
                      stepsIntro.textContent = strings.PpSetupIntro;
                      wrapper.appendChild(stepsIntro);

                      const steps = [
                        strings.WelcomeDialogSetupStep1Label,
                        strings.WelcomeDialogDeployToAzureLabel,
                        strings.WelcomeDialogSetupStep3Label,
                      ];
                      steps.forEach((text, i) => {
                        const step = document.createElement('div');
                        step.style.fontSize = '12px';
                        step.style.fontWeight = '600';
                        step.style.color = subtextColor;
                        step.textContent = `${i + 1}. ${text}`;
                        wrapper.appendChild(step);
                      });
                    }

                    // Re-run hint — shown only on error
                    if (isError) {
                      const rerunHint = document.createElement('div');
                      rerunHint.style.fontSize = '11px';
                      rerunHint.style.lineHeight = '1.5';
                      rerunHint.style.color = subtextColor;
                      rerunHint.style.marginTop = '2px';
                      rerunHint.textContent = strings.PpSetupRerunHint;
                      wrapper.appendChild(rerunHint);
                    }

                    // "View setup guide" link — always visible
                    const setupGuideLink = document.createElement('a');
                    setupGuideLink.textContent = strings.WelcomeDialogOptionApiDocsLabel;
                    setupGuideLink.href = setupGuideHref;
                    setupGuideLink.target = '_blank';
                    setupGuideLink.rel = 'noopener noreferrer';
                    setupGuideLink.style.fontSize = '12px';
                    setupGuideLink.style.color = linkColor;
                    setupGuideLink.style.textDecoration = 'none';
                    if (!isOk) setupGuideLink.style.marginTop = '2px';
                    wrapper.appendChild(setupGuideLink);

                    element.appendChild(wrapper);
                  },
                  onDispose: (element: HTMLElement | undefined) => {
                    if (element) element.innerHTML = '';
                  },
                }) as unknown as IPropertyPaneField<unknown>,
                ...(this.properties.functionUrl && this.properties.functionClientId ? [
                  PropertyPaneCustomField({
                    key: 'proxyStatusField',
                    onRender: (element: HTMLElement | undefined) => {
                      if (!element) return;
                      const status = this._proxyStatus;
                      const intent: 'success' | 'error' | 'info' = status === 'ok' ? 'success'
                        : status === 'error' ? 'error'
                        : 'info';
                      const text = status === 'ok' ? strings.ProxyStatusOk
                        : status === 'error' ? strings.ProxyStatusError
                        : strings.ProxyStatusChecking;
                      this._renderReactTree(
                        React.createElement(IdPrefixProvider, { value: `gsi-pp-${this.context.instanceId}-` },
                          React.createElement(RendererProvider, { renderer: griffelRenderer } as React.ComponentProps<typeof RendererProvider>,
                            React.createElement(FluentProvider,
                              { theme: this._theme ? createV9Theme(this._theme as unknown as Parameters<typeof createV9Theme>[0], this._theme.isInverted ? webDarkTheme : webLightTheme) : undefined, id: `gsi-pp-${this.context.instanceId}` },
                              React.createElement(MessageBar,
                                { intent, style: { marginTop: 8 } },
                                React.createElement(MessageBarBody, null, text)
                              )
                            )
                          )
                        ),
                        element
                      );
                    },
                    onDispose: (element: HTMLElement | undefined) => {
                      if (element) this._unmountReactTree(element);
                    },
                  }) as unknown as IPropertyPaneField<unknown>
                ] : [])
              ]
            },
            {
              groupName: strings.DiagnosticsGroupName,
              isCollapsed: true,
              groupFields: [
                this._descriptionField('diagnosticsGroupHint', strings.DiagnosticsGroupHint),
                // ── Check 1: Asset source / bundle delivery ─────────────────
                PropertyPaneCustomField({
                  key: 'diagAssetSourceField',
                  onRender: (element: HTMLElement | undefined) => {
                    if (!element) return;
                    const diag = this._diagnosticsResult;
                    let text: string;
                    let intent: 'info' | 'success' | 'warning' | 'error' = 'info';
                    if (!diag) {
                      text = strings.DiagnosticsChecking;
                      intent = 'info';
                    } else {
                      switch (diag.assetSource) {
                        case 'publicCdn':
                          text = strings.DiagPublicCdnOk;
                          intent = 'success';
                          break;
                        case 'privateCdn':
                          text = strings.DiagPrivateCdnInfo;
                          intent = 'info';
                          break;
                        case 'siteCollectionCatalog':
                          text = strings.DiagSiteColCatalogOk;
                          intent = 'success';
                          break;
                        case 'tenantCatalog':
                          if (diag.tenantCatalogEveryoneAccess === 'ok') {
                            text = strings.DiagTenantCatalogEveryoneOk;
                            intent = 'success';
                          } else if (diag.tenantCatalogEveryoneAccess === 'denied') {
                            text = strings.DiagTenantCatalogEveryoneDenied;
                            intent = 'error';
                          } else {
                            text = strings.DiagTenantCatalogEveryoneUnknown;
                            intent = 'warning';
                          }
                          break;
                        default:
                          text = strings.DiagAssetSourceUnknown;
                          intent = 'warning';
                      }
                    }
                    this._renderInfoBox(element, 'diagAssetSource', text, intent);
                  },
                  onDispose: (element: HTMLElement | undefined) => {
                    if (element) this._unmountReactTree(element);
                  },
                }) as unknown as IPropertyPaneField<unknown>,
                // ── Check 2: Site Everyone access ───────────────────────────
                PropertyPaneCustomField({
                  key: 'diagSiteEveryoneField',
                  onRender: (element: HTMLElement | undefined) => {
                    if (!element) return;
                    const diag = this._diagnosticsResult;
                    let text: string;
                    let intent: 'info' | 'success' | 'warning' | 'error' = 'info';
                    if (!diag) {
                      // Still loading — render nothing; diagAssetSourceField shows the spinner
                      element.innerHTML = '';
                      return;
                    }
                    switch (diag.siteEveryoneAccess) {
                      case 'ok':
                        text = strings.DiagSiteEveryoneOk;
                        intent = 'success';
                        break;
                      case 'missing':
                        text = strings.DiagSiteEveryoneMissing;
                        intent = 'warning';
                        break;
                      default:
                        text = strings.DiagSiteEveryoneUnknown;
                        intent = 'info';
                    }
                    this._renderInfoBox(element, 'diagSiteEveryone', text, intent);
                  },
                  onDispose: (element: HTMLElement | undefined) => {
                    if (element) this._unmountReactTree(element);
                  },
                }) as unknown as IPropertyPaneField<unknown>,
              ]
            },
            {
              groupName: strings.AuthorSectionGroupName,
              isCollapsed: false,
              groupFields: [
                PropertyPaneCustomField({
                  key: 'authorSectionCustomField',
                  onRender: (element: HTMLElement | undefined) => {
                    this._renderAuthorSection(element);
                  },
                  onDispose: (element: HTMLElement | undefined) => {
                    if (element) {
                      element.innerHTML = '';
                    }
                  },
                }) as unknown as IPropertyPaneField<unknown>,
              ]
            }
          ]
        }
      ]
    };
  }
}

