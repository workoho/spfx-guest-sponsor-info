import * as React from 'react';
import * as ReactDom from 'react-dom';
import { Version } from '@microsoft/sp-core-library';
import {
  type IPropertyPaneConfiguration,
  PropertyPaneTextField,
  PropertyPaneCheckbox,
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
  /** Show the manager section in the contact card. Default: true. */
  showManager: boolean;
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
        functionUrl: this.properties.functionUrl || undefined,
        functionClientId: this.properties.functionClientId || undefined,
        aadHttpClient: this._aadHttpClient,
        showBusinessPhones: this.properties.showBusinessPhones ?? true,
        showMobilePhone: this.properties.showMobilePhone ?? true,
        showWorkLocation: this.properties.showWorkLocation ?? true,
        showManager: this.properties.showManager ?? true,
        useInformalAddress: this.properties.useInformalAddress ?? false,
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

  protected getPropertyPaneConfiguration(): IPropertyPaneConfiguration {
    return {
      pages: [
        {
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
              groupName: strings.DisplayGroupName,
              groupFields: [
                PropertyPaneCheckbox('showBusinessPhones', {
                  text: strings.ShowBusinessPhonesFieldLabel,
                  checked: this.properties.showBusinessPhones ?? true,
                }),
                PropertyPaneCheckbox('showMobilePhone', {
                  text: strings.ShowMobilePhoneFieldLabel,
                  checked: this.properties.showMobilePhone ?? true,
                }),
                PropertyPaneCheckbox('showWorkLocation', {
                  text: strings.ShowWorkLocationFieldLabel,
                  checked: this.properties.showWorkLocation ?? true,
                }),
                PropertyPaneCheckbox('showManager', {
                  text: strings.ShowManagerFieldLabel,
                  checked: this.properties.showManager ?? true,
                }),
                PropertyPaneCheckbox('useInformalAddress', {
                  text: strings.UseInformalAddressFieldLabel,
                }),
              ]
            },
            {
              groupName: strings.FunctionGroupName,
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

