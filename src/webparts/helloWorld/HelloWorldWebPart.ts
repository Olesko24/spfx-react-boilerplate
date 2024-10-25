import * as React from 'react';
import * as ReactDom from 'react-dom';
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
import "../../styles/dist/tailwind.css";

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
  private _environmentMessage: string = '';
  private _appMode: AppMode = AppMode.SharePoint;
  private _theme: Theme = webLightTheme;

  public render(): void {
    const element: React.ReactElement<IAppProps> = React.createElement(
      App,
      {
        description: this.properties.description,
        isDarkTheme: this._isDarkTheme,
        environmentMessage: this._environmentMessage,
        hasTeamsContext: !!this.context.sdks.microsoftTeams,
        userDisplayName: this.context.pageContext.user.displayName
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

    ReactDom.render(idPrefixElement, this.domElement);
  }

  protected async onInit(): Promise<void> {
    const _l = this.context.isServedFromLocalhost;
    if (!!this.context.sdks.microsoftTeams) {
      const teamsContext = await this.context.sdks.microsoftTeams.teamsJs.app.getContext();
      switch (teamsContext.app.host.name.toLowerCase()) {
        case 'teams': this._appMode = _l ? AppMode.TeamsLocal : AppMode.Teams; break;
        case 'office': this._appMode = _l ? AppMode.OfficeLocal : AppMode.Office; break;
        case 'outlook': this._appMode = _l ? AppMode.OutlookLocal : AppMode.Outlook; break;
        default: throw new Error('Unknown host');
      }
    } else this._appMode = _l ? AppMode.SharePointLocal : AppMode.SharePoint;
    await initSP(this.context);
    return super.onInit();
  }

  protected onThemeChanged(currentTheme: IReadonlyTheme | undefined): void {
    if (!currentTheme) return;
    this._isDarkTheme = !!currentTheme.isInverted;
    //if the app mode is sharepoint, adjust the fluent ui 9 web light theme to use the sharepoint theme color, teams/dark mode should be fine on default
    if (this._appMode === AppMode.SharePoint || this._appMode === AppMode.SharePointLocal) {
      this._theme = createV9Theme(currentTheme as undefined, webLightTheme);
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
