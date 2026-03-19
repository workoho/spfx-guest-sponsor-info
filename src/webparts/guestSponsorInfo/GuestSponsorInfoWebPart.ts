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
    try {
      this._graphClient = await this.context.msGraphClientFactory.getClient('3');
    } catch {
      // Graph client unavailable – the web part will still render in edit-mode placeholder state.
    }
    if (this.properties.functionClientId) {
      try {
        this._aadHttpClient = await this.context.aadHttpClientFactory.getClient(
          this.properties.functionClientId
        );
      } catch {
        // AAD HTTP client unavailable – will fall back to direct Graph path.
      }
    }
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

