import * as React from 'react';
import * as ReactDom from 'react-dom';
import { Version } from '@microsoft/sp-core-library';
import {
  IPropertyPaneConfiguration,
  PropertyPaneTextField
} from '@microsoft/sp-property-pane';
import { BaseClientSideWebPart, WebPartContext } from '@microsoft/sp-webpart-base';

import * as strings from 'MuhimbiExportHtmlContentToPdfWebPartStrings';
import MuhimbiExportHtmlContentToPdf from './components/MuhimbiExportHtmlContentToPdf';
import { IMuhimbiExportHtmlContentToPdfProps } from './components/IMuhimbiExportHtmlContentToPdfProps';
import { HttpClient } from "@microsoft/sp-http";
import { ConvertFileService } from '../../services/ConvertFileService';

export interface IMuhimbiExportHtmlContentToPdfWebPartProps {
  description: string;
  context: WebPartContext;
  httpClient: HttpClient;
  apiKey: string;
  apiUrl: string;
}

export default class MuhimbiExportHtmlContentToPdfWebPart extends BaseClientSideWebPart<IMuhimbiExportHtmlContentToPdfWebPartProps> {

  private _services: ConvertFileService = null;

  public onInit(): Promise<void> {
    return super.onInit().then(_ => {
      this._services = new ConvertFileService(this.context);
    });
  }

  public render(): void {
    const element: React.ReactElement<IMuhimbiExportHtmlContentToPdfProps> = React.createElement(
      MuhimbiExportHtmlContentToPdf,
      {
        description: this.properties.description,
        apiKey: this.properties.apiKey ? this.properties.apiKey : "8878a0cb-371f-4ff3-b223-af122514bd2a",
        apiUrl: this.properties.apiUrl ? this.properties.apiUrl : "https://api.muhimbi.com/api/v1/operations/convert", //"https://api.muhimbi.com/api/v1/operations/convert_html",
        context: this.context,
        httpClient: this.properties.httpClient
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
                }),
                PropertyPaneTextField('apiKey', {
                  label: "API Key"
                }),
                PropertyPaneTextField('apiUrl', {
                  label: "API Url"
                })
              ]
            }
          ]
        }
      ]
    };
  }
}
