import * as React from 'react';
import * as ReactDom from 'react-dom';
import { Version } from '@microsoft/sp-core-library';
import {
  type IPropertyPaneConfiguration,
  PropertyPaneTextField,
  PropertyPaneCheckbox,
} from '@microsoft/sp-property-pane';
import { BaseClientSideWebPart } from '@microsoft/sp-webpart-base';
import { MSGraphClientV3 } from '@microsoft/sp-http';
import { initializeIcons } from '@fluentui/react';

import * as strings from 'GuestSponsorInfoWebPartStrings';
import GuestSponsorInfo from './components/GuestSponsorInfo';
import { IGuestSponsorInfoProps } from './components/IGuestSponsorInfoProps';

export interface IGuestSponsorInfoWebPartProps {
  title: string;
  mockMode: boolean;
}

export default class GuestSponsorInfoWebPart extends BaseClientSideWebPart<IGuestSponsorInfoWebPartProps> {

  private _graphClient: MSGraphClientV3 | undefined;

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
      }
    );

    ReactDom.render(element, this.domElement);
  }

  protected async onInit(): Promise<void> {
    await super.onInit();
    // Register the Fluent UI MDL2 icon font. This is safe to call multiple times;
    // the implementation skips re-registration if already initialised.
    initializeIcons();
    try {
      this._graphClient = await this.context.msGraphClientFactory.getClient('3');
    } catch {
      // Graph client unavailable – the web part will still render in edit-mode placeholder state.
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
            }
          ]
        }
      ]
    };
  }
}

