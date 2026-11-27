import * as React from 'react';
import { createRoot, type Root } from 'react-dom/client';
import { Version } from '@microsoft/sp-core-library';
import {
  type IPropertyPaneConfiguration,
  PropertyPaneTextField
} from '@microsoft/sp-property-pane';
import { BaseClientSideWebPart } from '@microsoft/sp-webpart-base';
import { IReadonlyTheme } from '@microsoft/sp-component-base';

import * as strings from 'HelloWorldWebPartStrings';
import App from './components/App';
import { IAppProps } from './components/App';
import { FluentProvider, FluentProviderProps, IdPrefixProvider, teamsDarkTheme, teamsLightTheme, Theme, webDarkTheme, webLightTheme } from '@fluentui/react-components';
import { createV9Theme } from '@fluentui/react-migration-v8-v9';
import { initSP } from '../../utils/sharepoint';

/**
 * Web part Property Pane Props
 */
export interface IHelloWorldWebPartProps {
  description: string;
}

export enum AppMode {
  SharePoint, SharePointLocal, Teams, TeamsLocal, Office, OfficeLocal, Outlook, OutlookLocal
}

export default class HelloWorldWebPart extends BaseClientSideWebPart<IHelloWorldWebPartProps> {

  private _isDarkTheme: boolean = false;
  private _appMode: AppMode = AppMode.SharePoint;
  private _theme: Theme = webLightTheme;
  private _root: Root | undefined;

  public render(): void {
    const element: React.ReactElement<IAppProps> = React.createElement(
      App,
      {
        description: this.properties.description
      }
    );

    //wrap the component with the Fluent UI 9 Provider.
    const fluentElement: React.ReactElement<FluentProviderProps> = React.createElement(
      FluentProvider,
      {
        theme: this._appMode === AppMode.Teams || this._appMode === AppMode.TeamsLocal ?
          this._isDarkTheme ? teamsDarkTheme : teamsLightTheme :
          this._appMode === AppMode.SharePoint || this._appMode === AppMode.SharePointLocal ?
            this._isDarkTheme ? webDarkTheme : this._theme :
            this._isDarkTheme ? webDarkTheme : webLightTheme,
        applyStylesToPortals: true
      },
      element
    );

    //wrap Fluent UI 9 Provider to IdPrefixProvider
    // https://react.fluentui.dev/?path=/docs/concepts-developer-advanced-configuration--docs#idprefixprovider
    const idPrefixElement: React.ReactElement = React.createElement(IdPrefixProvider, {
      value: 'hello-world' // change to unique id
    }, fluentElement);

    if (!this._root) {
      this._root = createRoot(this.domElement);
    }
    this._root.render(idPrefixElement);
  }

  protected async onInit(): Promise<void> {
    const _l = this.context.isServedFromLocalhost;
    if (!!this.context.sdks.microsoftTeams) {
      const teamsContext = await this.context.sdks.microsoftTeams.teamsJs.app.getContext();
      switch (teamsContext.app.host.name.toLowerCase()) {
        case 'teams': this._appMode = _l ? AppMode.TeamsLocal : AppMode.Teams; break;
        case 'teamsmodern': this._appMode = _l ? AppMode.TeamsLocal : AppMode.Teams; break;
        case 'office': this._appMode = _l ? AppMode.OfficeLocal : AppMode.Office; break;
        case 'outlook': this._appMode = _l ? AppMode.OutlookLocal : AppMode.Outlook; break;
        default: this._appMode = _l ? AppMode.SharePointLocal : AppMode.SharePoint;
      }
    } else this._appMode = _l ? AppMode.SharePointLocal : AppMode.SharePoint;
    initSP(this.context);
    return super.onInit();
  }

  protected onThemeChanged(currentTheme: IReadonlyTheme | undefined): void {
    if (!currentTheme) return;
    this._isDarkTheme = !!currentTheme.isInverted;
    //if the app mode is sharepoint, adjust the fluent ui 9 web light theme to use the sharepoint theme color, teams/dark mode should be fine on default
    if (this._appMode === AppMode.SharePoint || this._appMode === AppMode.SharePointLocal) {
      this._theme = createV9Theme(currentTheme as Parameters<typeof createV9Theme>[0], webLightTheme);
    }
  }

  protected onDispose(): void {
    this._root?.unmount();
    this._root = undefined;
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
