import * as React from 'react';
import * as ReactDom from 'react-dom';
import { Version } from '@microsoft/sp-core-library';
import {
  type IPropertyPaneConfiguration,
  type IPropertyPaneDropdownOption,
  PropertyPaneHorizontalRule,
  PropertyPaneLabel,
  PropertyPaneTextField,
  PropertyPaneCheckbox,
  PropertyPaneDropdown,
} from '@microsoft/sp-property-pane';
import { BaseClientSideWebPart } from '@microsoft/sp-webpart-base';
import { MSGraphClientV3, AadHttpClient } from '@microsoft/sp-http';
import { initializeIcons } from '@fluentui/react';

import * as strings from 'GuestSponsorInfoWebPartStrings';
import GuestSponsorInfo from './components/GuestSponsorInfo';
import { IGuestSponsorInfoProps } from './components/IGuestSponsorInfoProps';

export interface IGuestSponsorInfoWebPartProps {
  title: string;
  mockMode: boolean;
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
  /** Enable structured address fields (street/postal code/state). Default: false. */
  showStructuredAddress: boolean;
  /** Show the sponsor's street address. Default: false. */
  showStreetAddress: boolean;
  /** Show the sponsor's postal code. Default: false. */
  showPostalCode: boolean;
  /** Show the sponsor's state or province. Default: false. */
  showState: boolean;
  /** Enable inline map preview for address data. Default: false. */
  showAddressMap: boolean;
  /** Azure Maps subscription key used for inline map preview. */
  azureMapsSubscriptionKey: string;
  /** External map provider used for fallback links. */
  externalMapProvider: 'bing' | 'google' | 'apple' | 'openstreetmap' | 'here';
  /** Show the manager section in the contact card. Default: true. */
  showManager: boolean;
  /** Show the presence status indicator and label. Default: true. */
  showPresence: boolean;
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
        hostTenantId: this.context.pageContext.aadInfo.tenantId.toString(),
        functionUrl: this.properties.functionUrl
          ? `https://${this.properties.functionUrl.replace(/^https?:\/\//i, '').replace(/\/$/, '')}/api/getGuestSponsors`
          : undefined,
        presenceUrl: this.properties.functionUrl
          ? `https://${this.properties.functionUrl.replace(/^https?:\/\//i, '').replace(/\/$/, '')}/api/getPresence`
          : undefined,
        functionClientId: this.properties.functionClientId || undefined,
        aadHttpClient: this._aadHttpClient,
        showBusinessPhones: this.properties.showBusinessPhones ?? true,
        showMobilePhone: this.properties.showMobilePhone ?? true,
        showWorkLocation: this.properties.showWorkLocation ?? true,
        showCity: this.properties.showCity ?? false,
        showCountry: this.properties.showCountry ?? false,
        showStructuredAddress: this.properties.showStructuredAddress ?? false,
        showStreetAddress: this.properties.showStreetAddress ?? false,
        showPostalCode: this.properties.showPostalCode ?? false,
        showState: this.properties.showState ?? false,
        showAddressMap: this.properties.showAddressMap ?? false,
        azureMapsSubscriptionKey: this.properties.azureMapsSubscriptionKey || undefined,
        externalMapProvider: this.properties.externalMapProvider ?? 'bing',
        showManager: this.properties.showManager ?? true,
        showPresence: this.properties.showPresence ?? true,
        showSponsorJobTitle: this.properties.showSponsorJobTitle ?? true,
        showManagerJobTitle: this.properties.showManagerJobTitle ?? true,
        showSponsorDepartment: this.properties.showSponsorDepartment ?? false,
        showManagerDepartment: this.properties.showManagerDepartment ?? false,
        useInformalAddress: this.properties.useInformalAddress ?? false,
        clientVersion: this.manifest.version,
      }
    );

    ReactDom.render(element, this.domElement);
  }

  protected async onInit(): Promise<void> {
    await super.onInit();
    // Register the Fluent UI MDL2 icon font. Passing disableWarnings suppresses
    // the "icons re-registered" console warning that occurs when multiple web parts
    // or the SharePoint page itself have already called initializeIcons().
    initializeIcons(undefined, { disableWarnings: true });
    // Acquire Graph and AAD clients in the background — do NOT await them here.
    // SPFx awaits the onInit() Promise before rendering any web part on the page,
    // so blocking here would delay the entire page. Instead we resolve immediately
    // and re-render once clients are ready. The React component already shows a
    // shimmer while clients are undefined, so there is no visible flash.
    this._acquireClientsInBackground();
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
    ReactDom.unmountComponentAtNode(this.domElement);
  }

  protected get dataVersion(): Version {
    return Version.parse('1.0');
  }

  protected onPropertyPaneFieldChanged(propertyPath: string, oldValue: unknown, newValue: unknown): void {
    super.onPropertyPaneFieldChanged(propertyPath, oldValue, newValue);

    if (oldValue === newValue) return;
    if (
      propertyPath === 'showCity' ||
      propertyPath === 'showCountry' ||
      propertyPath === 'showWorkLocation' ||
      propertyPath === 'showStructuredAddress' ||
      propertyPath === 'showAddressMap' ||
      propertyPath === 'showManager'
    ) {
      this.context.propertyPane.refresh();
    }
  }

  private _getLocationDisplayHint(): string {
    const showCity = this.properties.showCity ?? false;
    const showCountry = this.properties.showCountry ?? false;
    const showWorkLocation = this.properties.showWorkLocation ?? true;
    const showStructuredAddress = this.properties.showStructuredAddress ?? false;
    if (showStructuredAddress) return strings.LocationDisplayHintAddressEnabled;
    if (showCity || showCountry || showWorkLocation) return strings.LocationDisplayHintSeparateRows;
    return strings.LocationDisplayHintHidden;
  }

  protected getPropertyPaneConfiguration(): IPropertyPaneConfiguration {
    const showManager = this.properties.showManager ?? true;
    const showStructuredAddress = this.properties.showStructuredAddress ?? false;
    const showAddressMap = this.properties.showAddressMap ?? false;
    const externalMapProvider = this.properties.externalMapProvider ?? 'bing';
    const mapProviderOptions: IPropertyPaneDropdownOption[] = [
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
              groupFields: [
                PropertyPaneTextField('title', {
                  label: strings.TitleFieldLabel
                }),
                PropertyPaneCheckbox('mockMode', {
                  text: strings.MockModeFieldLabel
                })
              ]
            },
            {
              // Card display: what fields appear on the sponsor card itself.
              groupName: strings.DisplayGroupName,
              groupFields: [
                PropertyPaneCheckbox('showPresence', {
                  text: strings.ShowPresenceFieldLabel,
                  checked: this.properties.showPresence ?? true,
                }),
                PropertyPaneCheckbox('showSponsorJobTitle', {
                  text: strings.ShowSponsorJobTitleFieldLabel,
                  checked: this.properties.showSponsorJobTitle ?? true,
                }),
                PropertyPaneCheckbox('showSponsorDepartment', {
                  text: strings.ShowSponsorDepartmentFieldLabel,
                }),
                PropertyPaneHorizontalRule(),
                PropertyPaneCheckbox('useInformalAddress', {
                  text: strings.UseInformalAddressFieldLabel,
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
                // Simple location fields
                PropertyPaneCheckbox('showCity', {
                  text: strings.ShowCityFieldLabel,
                }),
                PropertyPaneCheckbox('showCountry', {
                  text: strings.ShowCountryFieldLabel,
                }),
                PropertyPaneCheckbox('showWorkLocation', {
                  text: strings.ShowWorkLocationFieldLabel,
                  checked: this.properties.showWorkLocation ?? true,
                }),
                PropertyPaneLabel('locationDisplayHint', {
                  text: this._getLocationDisplayHint(),
                }),
                PropertyPaneHorizontalRule(),
                // Structured address (street/ZIP/state)
                PropertyPaneCheckbox('showStructuredAddress', {
                  text: strings.ShowStructuredAddressFieldLabel,
                  checked: showStructuredAddress,
                }),
                ...(showStructuredAddress ? [
                  PropertyPaneCheckbox('showStreetAddress', {
                    text: strings.ShowStreetAddressFieldLabel,
                    checked: this.properties.showStreetAddress ?? false,
                  }),
                  PropertyPaneCheckbox('showPostalCode', {
                    text: strings.ShowPostalCodeFieldLabel,
                    checked: this.properties.showPostalCode ?? false,
                  }),
                  PropertyPaneCheckbox('showState', {
                    text: strings.ShowStateFieldLabel,
                    checked: this.properties.showState ?? false,
                  }),
                ] : []),
                PropertyPaneHorizontalRule(),
                // Map preview
                PropertyPaneCheckbox('showAddressMap', {
                  text: strings.ShowAddressMapFieldLabel,
                  checked: showAddressMap,
                }),
                ...(showAddressMap ? [
                  PropertyPaneTextField('azureMapsSubscriptionKey', {
                    label: strings.AzureMapsSubscriptionKeyFieldLabel,
                  }),
                  PropertyPaneDropdown('externalMapProvider', {
                    label: strings.ExternalMapProviderFieldLabel,
                    options: mapProviderOptions,
                    selectedKey: externalMapProvider,
                  }),
                  PropertyPaneLabel('mapProviderHint', {
                    text: strings.AddressMapFallbackHint,
                  }),
                ] : []),
              ]
            },
            {
              groupName: strings.OrganizationSection,
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
                })
              ]
            }
          ]
        }
      ]
    };
  }
}

