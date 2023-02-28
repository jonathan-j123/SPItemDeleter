import * as React from 'react';
import * as ReactDom from 'react-dom';
import { Version } from '@microsoft/sp-core-library';
import {
  IPropertyPaneConfiguration,
  PropertyPaneTextField
} from '@microsoft/sp-property-pane';
import { BaseClientSideWebPart } from '@microsoft/sp-webpart-base';

import * as strings from 'ProductivityAppWebPartStrings';
import ProductivityApp from './components/ProductivityApp';
import { IProductivityAppProps } from './components/IProductivityAppProps';

import { getSP } from './pnpjsConfig';


export interface IProductivityAppWebPartProps {
  description: string;
}

export default class ProductivityAppWebPart extends BaseClientSideWebPart<IProductivityAppWebPartProps> {

  private _isDarkTheme: boolean = false;
  private _environmentMessage: string = '';

  public render(): void {

      const element: React.ReactElement<IProductivityAppProps> = React.createElement(
        ProductivityApp,
        {
          description: this.properties.description,
          isDarkTheme: this._isDarkTheme,
          environmentMessage: this._environmentMessage,
          hasTeamsContext: !!this.context.sdks.microsoftTeams,
          userDisplayName: this.context.pageContext.user.displayName
        }
      );
  
      ReactDom.render(element, this.domElement);
    
  }

public async onInit(): Promise<void> {

  await super.onInit();
  getSP(this.context);
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
