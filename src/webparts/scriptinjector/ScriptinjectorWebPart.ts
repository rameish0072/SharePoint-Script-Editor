import * as React from 'react';
import * as ReactDom from 'react-dom';
import { Version } from '@microsoft/sp-core-library';
import {
  type IPropertyPaneConfiguration,
  PropertyPaneTextField
} from '@microsoft/sp-property-pane';
import { BaseClientSideWebPart } from '@microsoft/sp-webpart-base';
import { IReadonlyTheme } from '@microsoft/sp-component-base';

import * as strings from 'ScriptinjectorWebPartStrings';
import Scriptinjector from './components/Scriptinjector';
import { IScriptinjectorProps } from './components/IScriptinjectorProps';

export interface IScriptinjectorWebPartProps {
  htmlInput:    string;
  htmlFileName: string;
}

export default class ScriptinjectorWebPart extends BaseClientSideWebPart<IScriptinjectorWebPartProps> {

  private _isDarkTheme: boolean = false;

  public render(): void {
    const element: React.ReactElement<IScriptinjectorProps> = React.createElement(
      Scriptinjector,
      {
        htmlInput:       this.properties.htmlInput,
        htmlFileName:    this.properties.htmlFileName,
        isDarkTheme:     this._isDarkTheme,
        hasTeamsContext: !!this.context.sdks.microsoftTeams,
        instanceId:      this.instanceId,
        displayMode:     this.displayMode,
      }
    );
    ReactDom.render(element, this.domElement);
  }

  protected onInit(): Promise<void> {
    return this._getEnvironmentMessage().then(() => { });
  }

  private _getEnvironmentMessage(): Promise<string> {
    if (!!this.context.sdks.microsoftTeams) {
      return this.context.sdks.microsoftTeams.teamsJs.app.getContext()
        .then(context => {
          let environmentMessage: string = '';
          switch (context.app.host.name) {
            case 'Office':
              environmentMessage = this.context.isServedFromLocalhost ? strings.AppLocalEnvironmentOffice : strings.AppOfficeEnvironment;
              break;
            case 'Outlook':
              environmentMessage = this.context.isServedFromLocalhost ? strings.AppLocalEnvironmentOutlook : strings.AppOutlookEnvironment;
              break;
            case 'Teams':
            case 'TeamsModern':
              environmentMessage = this.context.isServedFromLocalhost ? strings.AppLocalEnvironmentTeams : strings.AppTeamsTabEnvironment;
              break;
            default:
              environmentMessage = strings.UnknownEnvironment;
          }
          return environmentMessage;
        });
    }
    return Promise.resolve(this.context.isServedFromLocalhost ? strings.AppLocalEnvironmentSharePoint : strings.AppSharePointEnvironment);
  }

  protected onThemeChanged(currentTheme: IReadonlyTheme | undefined): void {
    if (!currentTheme) return;
    this._isDarkTheme = !!currentTheme.isInverted;
    const { semanticColors } = currentTheme;
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
          header: { description: strings.PropertyPaneDescription },
          groups: [
            {
              groupName: 'Inline HTML',
              groupFields: [
                PropertyPaneTextField('htmlInput', {
                  label: 'HTML Code',
                  description: 'Paste HTML code here. Rendered directly on page. CSS and JS inside will execute.',
                  multiline: true,
                  rows: 10,
                  placeholder: `<div>Hello from ScriptInjector!</div>`,
                })
              ]
            },
            {
              groupName: 'HTML File',
              groupFields: [
                PropertyPaneTextField('htmlFileName', {
                  label: 'HTML File URL',
                  description: 'Paste full URL of HTML file from Site Assets. CSS, JS and body will all be extracted and executed automatically.',
                  placeholder: 'https://tenant.sharepoint.com/sites/mysite/SiteAssets/myfile.html',
                })
              ]
            }
          ]
        }
      ]
    };
  }
}