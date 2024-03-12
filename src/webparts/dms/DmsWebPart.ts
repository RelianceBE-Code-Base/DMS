import * as React from 'react';
import * as ReactDom from 'react-dom';
import { Version } from '@microsoft/sp-core-library';
import {
  IPropertyPaneConfiguration,
  PropertyPaneTextField
} from '@microsoft/sp-property-pane';
import { BaseClientSideWebPart } from '@microsoft/sp-webpart-base';
import { IReadonlyTheme } from '@microsoft/sp-component-base';

import { SPFI } from '@pnp/sp';
import { IDocumentLibraryInformation } from "@pnp/sp/sites";
import "@pnp/sp/sites";
import { useEffect, useState } from 'react';
//import { getSP } from '../../../pnpjsConfig';


import "@pnp/sp/webs";
import "@pnp/sp/folders";
import { IFileInfo, IFolderInfo, Web } from "@pnp/sp/presets/all";
import { folderFromAbsolutePath } from "@pnp/sp/folders";
import { folderFromPath } from "@pnp/sp/folders";
import { folderFromServerRelativePath } from "@pnp/sp/folders";
import "@pnp/sp/webs";
import "@pnp/sp/presets/all";


import * as strings from 'DmsWebPartStrings';
import Dms from './components/Dms';
import { IDmsProps } from './components/IDmsProps';

import TreeView from './components/modules/Treeview';
//import HomeScreen from './components/modules/HomeScreen';

export interface IDmsWebPartProps {
  description: string;
 
}

export default class DmsWebPart extends BaseClientSideWebPart<IDmsWebPartProps> {

  private _isDarkTheme: boolean = false;
  private _environmentMessage: string = '';

  public render(): void {
    const element: React.ReactElement<IDmsProps> = React.createElement(
      Dms,
      {
        description: this.properties.description,
        // isDarkTheme: this._isDarkTheme,
        // environmentMessage: this._environmentMessage,
        // hasTeamsContext: !!this.context.sdks.microsoftTeams,
        userDisplayName: this.context.pageContext.user.displayName,
        context: this.context,
        siteUrl: this.context.pageContext.web.absoluteUrl,
        folderUrl: this.context.pageContext.site.serverRelativeUrl

      }
    );

    ReactDom.render(element, this.domElement);
  }

  protected onInit(): Promise<void> {
    this._environmentMessage = this._getEnvironmentMessage();

    return super.onInit();
  }

  private _getEnvironmentMessage(): string {
    if (!!this.context.sdks.microsoftTeams) { // running in Teams
      return this.context.isServedFromLocalhost ? strings.AppLocalEnvironmentTeams : strings.AppTeamsTabEnvironment;
    }

    return this.context.isServedFromLocalhost ? strings.AppLocalEnvironmentSharePoint : strings.AppSharePointEnvironment;
  }

  protected onThemeChanged(currentTheme: IReadonlyTheme | undefined): void {
    if (!currentTheme) {
      return;
    }

    this._isDarkTheme = !!currentTheme.isInverted;
    const {
      semanticColors
    } = currentTheme;

    if (semanticColors) {
      this.domElement.style.setProperty('--bodyText', semanticColors.bodyText || null);
      this.domElement.style.setProperty('--link', semanticColors.link || null);
      this.domElement.style.setProperty('--linkHovered', semanticColors.linkHovered || null);
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
