import * as React from 'react';
import * as ReactDom from 'react-dom';
import { Version } from '@microsoft/sp-core-library';
import {
  type IPropertyPaneConfiguration,
  type IPropertyPaneDropdownOption,
  type IPropertyPaneField,
  PropertyPaneHorizontalRule,
  PropertyPaneLabel,
  PropertyPaneTextField,
  PropertyPaneCheckbox,
  PropertyPaneDropdown,
  PropertyPaneSlider,
} from '@microsoft/sp-property-pane';
import { PropertyPaneCustomField } from '@microsoft/sp-property-pane/lib/propertyPaneFields/propertyPaneCustomField/PropertyPaneCustomField';
import { BaseClientSideWebPart } from '@microsoft/sp-webpart-base';
import { MSGraphClientV3, AadHttpClient } from '@microsoft/sp-http';
import { ThemeProvider, IReadonlyTheme, ThemeChangedEventArgs } from '@microsoft/sp-component-base';
import { FluentProvider, MessageBar, MessageBarBody, webLightTheme, webDarkTheme } from '@fluentui/react-components';
import { createV9Theme } from '@fluentui/react-migration-v8-v9';

import * as strings from 'GuestSponsorInfoWebPartStrings';
import GuestSponsorInfo from './components/GuestSponsorInfo';
import { IGuestSponsorInfoProps } from './components/IGuestSponsorInfoProps';
import workohoDefaultLogo from './assets/workoho-default-logo.svg';

export interface IGuestSponsorInfoWebPartProps {
  title: string;
  mockMode: boolean;
  /**
   * Notification to simulate in demo mode.
   * Replaces the former mockTeamsUnavailable boolean with a dropdown selection.
   * Default: 'none'.
   */
  /** Maximum number of sponsors shown to visitors on the live page (1–5). Default: 2. */
  maxSponsorCount: number;
  /** Number of mock sponsor cards to show in demo mode (1–5). Default: 2. */
  mockSponsorCount: number;
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
  /** Azure Maps subscription key used for inline map preview. */
  azureMapsSubscriptionKey: string;
  /** External map provider used for fallback links. */
  externalMapProvider: 'bing' | 'google' | 'apple' | 'openstreetmap' | 'here' | 'none';
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
}

export default class GuestSponsorInfoWebPart extends BaseClientSideWebPart<IGuestSponsorInfoWebPartProps> {

  private _graphClient: MSGraphClientV3 | undefined;
  private _aadHttpClient: AadHttpClient | undefined;
  private _proxyStatus: 'checking' | 'ok' | 'error' = 'checking';
  private _versionMismatch = false;
  private _themeProvider: ThemeProvider | undefined;
  private _theme: IReadonlyTheme | undefined;

  public render(): void {
    const element: React.ReactElement<IGuestSponsorInfoProps> = React.createElement(
      GuestSponsorInfo,
      {
        loginName: this.context.pageContext.user.loginName,
        isExternalGuestUser: this.context.pageContext.user.isExternalGuestUser,
        displayMode: this.displayMode,
        graphClient: this._graphClient, // undefined until onInit resolves
        title: this.properties.title,
        mockMode: this.properties.mockMode ?? false,
        maxSponsorCount: this.properties.maxSponsorCount ?? 2,
        mockSponsorCount: this.properties.mockSponsorCount ?? 2,
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
        azureMapsSubscriptionKey: this.properties.azureMapsSubscriptionKey || undefined,
        externalMapProvider: this.properties.externalMapProvider ?? 'bing',
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
        fluentProviderId: `gsi-${this.context.instanceId}`,
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
        theme: this._theme,
      }
    );

    ReactDom.render(element, this.domElement);
  }

  protected async onInit(): Promise<void> {
    await super.onInit();
    // Consume the SPFx ThemeProvider service so the FluentProvider can receive the
    // host site's v8-style theme and convert it to a v9 theme via createV9Theme.
    this._themeProvider = this.context.serviceScope.consume(ThemeProvider.serviceKey);
    this._theme = this._themeProvider?.tryGetTheme();
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
    this.render();
  };

  /**
   * Starts Graph and AAD HTTP client acquisition concurrently without blocking
   * the SPFx page lifecycle. Calls render() once both have settled so the React
   * component receives real clients in a single props update (avoids triggering
   * two separate data fetches if both clients resolve close together).
   * Both sub-promises catch their own errors and always resolve, so Promise.all
   * is safe to use here without needing Promise.allSettled (ES2020).
   */
  private _acquireClientsInBackground(): void {
    const graphPromise = this.context.msGraphClientFactory.getClient('3')
      .then(client => { this._graphClient = client; })
      .catch(() => { /* Graph unavailable — component renders in placeholder state */ });

    const clientId = this.properties.functionClientId;
    const aadPromise = clientId
      ? this.context.aadHttpClientFactory.getClient(clientId)
          .then(client => { this._aadHttpClient = client; })
          .catch(() => { /* AAD client unavailable — falls back to direct Graph path */ })
      : Promise.resolve();

    // Both sub-promises have their own .catch() and always resolve — Promise.all is safe here.
    // The trailing .catch() satisfies @typescript-eslint/no-floating-promises without the
    // void operator (which no-void disallows). In practice this catch is never reached.
    Promise.all([graphPromise, aadPromise]).then(() => {
      this.render();
    }).catch(() => { /* sub-promises never reject; unreachable */ });
  }

  protected onDispose(): void {
    this._themeProvider?.themeChangedEvent.remove(this, this._handleThemeChanged);
    ReactDom.unmountComponentAtNode(this.domElement);
  }

  protected get dataVersion(): Version {
    return Version.parse('1.0');
  }

  protected onPropertyPaneFieldChanged(propertyPath: string, oldValue: unknown, newValue: unknown): void {
    super.onPropertyPaneFieldChanged(propertyPath, oldValue, newValue);

    if (oldValue === newValue) return;
    if (
      propertyPath === 'functionUrl' ||
      propertyPath === 'functionClientId'
    ) {
      this._proxyStatus = 'checking';
      this.context.propertyPane.refresh();
    } else if (
      propertyPath === 'mockMode' ||
      propertyPath === 'cardLayout' ||
      propertyPath === 'showCity' ||
      propertyPath === 'showCountry' ||
      propertyPath === 'showWorkLocation' ||
      propertyPath === 'showManager'
    ) {
      this.context.propertyPane.refresh();
    }
  }

  private _getLocationDisplayHint(): string {
    if (!(strings as unknown as object | undefined)) return '';
    const showCity = this.properties.showCity ?? false;
    const showCountry = this.properties.showCountry ?? false;
    const showWorkLocation = this.properties.showWorkLocation ?? true;
    const showStreet = this.properties.showStreetAddress ?? false;
    const showPostal = this.properties.showPostalCode ?? false;
    const showState = this.properties.showState ?? false;
    if (showCity || showCountry || showWorkLocation || showStreet || showPostal || showState) return strings.LocationDisplayHintSeparateRows;
    return strings.LocationDisplayHintHidden;
  }

  private _renderAuthorSection(element?: HTMLElement): void {
    if (!element) return;

    element.innerHTML = '';

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
      ctaLink.style.border = '1px solid currentColor';
      ctaLink.style.borderRadius = '4px';
      ctaLink.style.fontWeight = '600';
      ctaLink.style.fontSize = '12px';
      ctaLink.style.textDecoration = 'none';
      return ctaLink;
    };

    const createGitHubIcon = (): SVGSVGElement => {
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
        'M8 0C3.58 0 0 3.58 0 8c0 3.54 2.29 6.53 5.47 7.59.4.07.55-.17.55-.38' +
        ' 0-.19-.01-.82-.01-1.49-2.01.37-2.53-.49-2.69-.94-.09-.23-.48-.94-.82-1.13' +
        '-.28-.15-.68-.52-.01-.53.63-.01 1.08.58 1.23.82.72 1.21 1.87.87 2.33.66' +
        '.07-.52.28-.87.51-1.07-1.78-.2-3.64-.89-3.64-3.95 0-.87.31-1.59.82-2.15' +
        '-.08-.2-.36-1.02.08-2.12 0 0 .67-.21 2.2.82a7.54 7.54 0 012-.27c.68 0 1.36.09' +
        ' 2 .27 1.53-1.04 2.2-.82 2.2-.82.44 1.1.16 1.92.08 2.12.51.56.82 1.27.82 2.15' +
        ' 0 3.07-1.87 3.75-3.65 3.95.29.25.54.73.54 1.48 0 1.07-.01 1.93-.01 2.2' +
        ' 0 .21.15.46.55.38A8.013 8.013 0 0016 8c0-4.42-3.58-8-8-8z'
      );
      path.setAttribute('fill', 'currentColor');

      icon.appendChild(path);
      return icon;
    };

    const logoLink = document.createElement('a');
    logoLink.href = 'https://workoho.com/';
    logoLink.target = '_blank';
    logoLink.rel = 'noopener noreferrer';
    logoLink.style.display = 'block';
    logoLink.style.width = '100%';
    logoLink.style.margin = '0 0 12px 0';
    logoLink.style.textDecoration = 'none';

    const logo = document.createElement('img');
    logo.src = workohoDefaultLogo;
    logo.alt = 'Workoho GmbH logo';
    logo.style.display = 'block';
    logo.style.width = 'min(150px, 50%)';
    logo.style.height = 'auto';
    logo.style.maxWidth = '150px';
    logo.style.aspectRatio = '4 / 1';
    logo.style.objectFit = 'contain';
    logoLink.appendChild(logo);
    element.appendChild(logoLink);

    element.appendChild(createParagraph(strings.AuthorSectionIntro, '8px', '700'));
    element.appendChild(createParagraph(strings.AuthorSectionConsultingText));
    element.appendChild(createCtaLink(strings.AuthorSectionWebsiteLinkLabel, 'https://workoho.com/'));

    const partnerLine = document.createElement('p');
    partnerLine.style.margin = '0 0 8px 0';
    partnerLine.style.fontWeight = '700';
    partnerLine.style.lineHeight = '1.45';
    partnerLine.append(`${strings.AuthorSectionPartnerPrefix} `);

    const easyLifeLink = document.createElement('a');
    easyLifeLink.textContent = strings.AuthorSectionPartnerLinkLabel;
    easyLifeLink.href = 'https://easylife365.cloud/';
    easyLifeLink.target = '_blank';
    easyLifeLink.rel = 'noopener noreferrer';
    easyLifeLink.style.fontWeight = '500';
    easyLifeLink.style.textDecoration = 'none';

    partnerLine.appendChild(easyLifeLink);
    partnerLine.append(` ${strings.AuthorSectionPartnerSuffix}`);

    const easyLifeBox = document.createElement('div');
    easyLifeBox.style.backgroundColor = '#f5f5f5';
    easyLifeBox.style.border = '1px solid #e1e1e1';
    easyLifeBox.style.borderRadius = '6px';
    easyLifeBox.style.padding = '10px 12px';
    easyLifeBox.style.margin = '0';

    easyLifeBox.appendChild(partnerLine);
    easyLifeBox.appendChild(createParagraph(strings.AuthorSectionPartnerTagline, '0'));
    element.appendChild(easyLifeBox);

    const footer = document.createElement('div');
    footer.style.marginTop = '10px';
    footer.style.fontSize = '12px';
    footer.style.color = '#616161';

    const metaLine = document.createElement('div');

    const sourceLink = document.createElement('a');
    sourceLink.appendChild(createGitHubIcon());
    sourceLink.append(strings.AuthorSectionSourceCodeLabel);
    sourceLink.href = 'https://github.com/workoho/spfx-guest-sponsor-info';
    sourceLink.target = '_blank';
    sourceLink.rel = 'noopener noreferrer';
    sourceLink.style.display = 'inline';
    sourceLink.style.textDecoration = 'none';

    const separator = document.createElement('span');
    separator.textContent = ' · ';

    const semver = this.manifest.version.split('.').slice(0, 3).join('.');
    const versionLink = document.createElement('a');
    versionLink.textContent = `${strings.AuthorSectionVersionLabel}: ${this.manifest.version}`;
    versionLink.href = `https://github.com/workoho/spfx-guest-sponsor-info/releases/tag/v${semver}`;
    versionLink.target = '_blank';
    versionLink.rel = 'noopener noreferrer';
    versionLink.style.textDecoration = 'none';

    const munichLine = document.createElement('span');
    munichLine.textContent = 'Built in Munich, World City with ❤';
    munichLine.style.display = 'inline-block';
    munichLine.style.marginTop = '8px';
    munichLine.style.padding = '2px 8px';
    munichLine.style.borderRadius = '999px';
    munichLine.style.backgroundColor = '#eef6ff';
    munichLine.style.color = '#24527a';
    munichLine.style.fontWeight = '600';
    munichLine.style.fontSize = '11px';

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
      updateBadge.style.backgroundColor = '#eff6fc';
      updateBadge.style.color = '#005a9e';
      updateBadge.style.border = '1px solid #c7e0f4';
      updateBadge.style.fontWeight = '700';
      updateBadge.style.fontSize = '11px';
      updateBadge.style.cursor = 'help';
      metaLine.appendChild(updateSep);
      metaLine.appendChild(updateBadge);
    }

    footer.appendChild(metaLine);
    footer.appendChild(document.createElement('br'));
    footer.appendChild(munichLine);
    element.appendChild(footer);
  }

  protected getPropertyPaneConfiguration(): IPropertyPaneConfiguration {
    // Guard: SPFx AMD locale bundles load asynchronously. If the property pane is opened
    // before the bundle resolves (can happen during rapid initialisation), return an empty
    // config so the framework doesn't crash. The pane will re-render once strings are ready.
    if (!(strings as unknown as object | undefined)) return { pages: [] };

    const showManager = this.properties.showManager ?? true;
    const externalMapProvider = this.properties.externalMapProvider ?? 'bing';
    const mapProviderOptions: IPropertyPaneDropdownOption[] = [
      { key: 'none', text: strings.MapProviderNoneOption },
      { key: 'bing', text: strings.MapProviderBingOption },
      { key: 'google', text: strings.MapProviderGoogleOption },
      { key: 'apple', text: strings.MapProviderAppleOption },
      { key: 'openstreetmap', text: strings.MapProviderOpenStreetMapOption },
      { key: 'here', text: strings.MapProviderHereOption },
    ];

    return {
      pages: [
        {
          displayGroupsAsAccordion: true,
          header: {
            description: strings.PropertyPaneDescription
          },
          groups: [
            {
              groupName: strings.BasicGroupName,
              isCollapsed: false,
              groupFields: [
                PropertyPaneTextField('title', {
                  label: strings.TitleFieldLabel
                }),
                PropertyPaneHorizontalRule(),
                PropertyPaneSlider('maxSponsorCount', {
                  label: strings.MaxSponsorCountFieldLabel,
                  min: 1,
                  max: 5,
                  step: 1,
                  value: this.properties.maxSponsorCount ?? 2,
                }),
                PropertyPaneHorizontalRule(),
                PropertyPaneSlider('mockSponsorCount', {
                  label: strings.MockSponsorCountFieldLabel,
                  min: 1,
                  max: 5,
                  step: 1,
                  value: this.properties.mockSponsorCount ?? 2,
                }),
                PropertyPaneCheckbox('mockMode', {
                  text: strings.MockModeFieldLabel
                }),
                PropertyPaneHorizontalRule(),
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
                })
              ]
            },
            {
              groupName: strings.GuestNotificationsGroupName,
              isCollapsed: true,
              groupFields: [
                PropertyPaneLabel('guestNotificationsGroupHint', {
                  text: strings.GuestNotificationsGroupHint,
                }),
                PropertyPaneHorizontalRule(),
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
                PropertyPaneHorizontalRule(),
                PropertyPaneCheckbox('showPresence', {
                  text: strings.ShowPresenceFieldLabel,
                  checked: this.properties.showPresence ?? true,
                }),
                PropertyPaneCheckbox('showSponsorPhoto', {
                  text: strings.ShowSponsorPhotoFieldLabel,
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
                PropertyPaneCheckbox('showBusinessPhones', {
                  text: strings.ShowBusinessPhonesFieldLabel,
                  checked: this.properties.showBusinessPhones ?? true,
                }),
                PropertyPaneCheckbox('showMobilePhone', {
                  text: strings.ShowMobilePhoneFieldLabel,
                  checked: this.properties.showMobilePhone ?? true,
                }),
                PropertyPaneHorizontalRule(),
                // Work location (officeLocation)
                PropertyPaneCheckbox('showWorkLocation', {
                  text: strings.ShowWorkLocationFieldLabel,
                  checked: this.properties.showWorkLocation ?? true,
                }),
                PropertyPaneHorizontalRule(),
                // Address fields (combined into a single clickable address row)
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
                PropertyPaneLabel('locationDisplayHint', {
                  text: this._getLocationDisplayHint(),
                }),
                PropertyPaneHorizontalRule(),
                // Map settings
                PropertyPaneDropdown('externalMapProvider', {
                  label: strings.ExternalMapProviderFieldLabel,
                  options: mapProviderOptions,
                  selectedKey: externalMapProvider,
                }),
                PropertyPaneLabel('addressMapProviderHint', {
                  text: strings.AddressMapProviderHint,
                }),
                PropertyPaneTextField('azureMapsSubscriptionKey', {
                  label: strings.AzureMapsSubscriptionKeyFieldLabel,
                }),
              ]
            },
            {
              groupName: strings.OrganizationSection,
              isCollapsed: true,
              groupFields: [
                PropertyPaneCheckbox('showManager', {
                  text: strings.ShowManagerFieldLabel,
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
                ] : [
                  PropertyPaneLabel('managerOptionsDisabledHint', {
                    text: strings.ManagerOptionsDisabledHint,
                  }),
                ]),
              ]
            },
            {
              groupName: strings.FunctionGroupName,
              isCollapsed: true,
              groupFields: [
                PropertyPaneTextField('functionUrl', {
                  label: strings.FunctionUrlFieldLabel
                }),
                PropertyPaneTextField('functionClientId', {
                  label: strings.FunctionClientIdFieldLabel
                }),
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
                      ReactDom.render(
                        React.createElement(FluentProvider,
                          { theme: this._theme ? createV9Theme(this._theme as unknown as Parameters<typeof createV9Theme>[0], this._theme.isInverted ? webDarkTheme : webLightTheme) : undefined, id: `gsi-pp-${this.context.instanceId}` },
                          React.createElement(MessageBar,
                            { intent, style: { marginTop: 8 } },
                            React.createElement(MessageBarBody, null, text)
                          )
                        ),
                        element
                      );
                    },
                    onDispose: (element: HTMLElement | undefined) => {
                      if (element) ReactDom.unmountComponentAtNode(element);
                    },
                  }) as unknown as IPropertyPaneField<unknown>
                ] : [])
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

