import * as React from 'react';
import * as ReactDom from 'react-dom';
import { DisplayMode, Version } from '@microsoft/sp-core-library';
import {
  type IPropertyPaneConfiguration,
  PropertyPaneTextField
} from '@microsoft/sp-property-pane';
import { BaseClientSideWebPart } from '@microsoft/sp-webpart-base';

import * as strings from 'FTLAppWebPartStrings';
import FTLApp from './components/FTLApp';
import { IFTLAppProps } from './components/IFTLAppProps';

export interface IFTLAppWebPartProps {
  description: string;
}

export default class FTLAppWebPart extends BaseClientSideWebPart<IFTLAppWebPartProps> {
  protected onInit(): Promise<void> {
    return super.onInit();
  }

  protected onDisplayModeChanged(oldDisplayMode: DisplayMode): void {
    super.onDisplayModeChanged(oldDisplayMode);
    this.render();
  }

  public render(): void {
    const element: React.ReactElement<IFTLAppProps> = React.createElement(
      FTLApp,
      {
        description: this.properties.description,
        isServedFromLocalhost: this.context.isServedFromLocalhost,
        isFullScreen: this.displayMode === DisplayMode.Read,
        spHttpClient: this.context.spHttpClient,
        webUrl: this.context.pageContext.web.absoluteUrl
      }
    );

    ReactDom.render(element, this.domElement);
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
                PropertyPaneTextField('description', {
                  label: strings.DescriptionFieldLabel
                })
              ]
            }
          ]
        }
      ]
    };
  }
}
