import * as React from 'react';
import * as ReactDom from 'react-dom';
import { Version } from '@microsoft/sp-core-library';
import {
  IPropertyPaneConfiguration,
  PropertyPaneTextField
} from '@microsoft/sp-property-pane';
import { BaseClientSideWebPart } from '@microsoft/sp-webpart-base';
import { SPComponentLoader } from "@microsoft/sp-loader";
import { sp } from "@pnp/sp/presets/all";

import * as strings from 'SpfxBoilerplateCodeWebPartStrings';
import SpfxBoilerplateCode from './components/SpfxBoilerplateCode';
import { ISpfxBoilerplateCodeProps } from './components/ISpfxBoilerplateCodeProps';
import "jquery";
import "bootstrap";
let cssURL =
  "https://cdn.jsdelivr.net/npm/bootstrap@4.5.3/dist/css/bootstrap.min.css"; 
SPComponentLoader.loadCss(cssURL);

import { initializeIcons } from "@uifabric/icons";
initializeIcons();

export interface ISpfxBoilerplateCodeWebPartProps {
  description: string;
}

export default class SpfxBoilerplateCodeWebPart extends BaseClientSideWebPart<ISpfxBoilerplateCodeWebPartProps> {

  public render(): void {
    const element: React.ReactElement<ISpfxBoilerplateCodeProps> = React.createElement(
      SpfxBoilerplateCode,
      {
        siteUrl: this.context.pageContext.web.absoluteUrl,
        spHttpClient: this.context.spHttpClient,
        context: this.context
      }
    );

    ReactDom.render(element, this.domElement);
  }
  protected onInit(): Promise<void> {
    sp.setup({
      spfxContext: this.context,
    });
    return super.onInit();
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
