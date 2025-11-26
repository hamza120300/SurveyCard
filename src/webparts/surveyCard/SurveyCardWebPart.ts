import * as React from 'react';
import * as ReactDom from 'react-dom';
import { Version } from '@microsoft/sp-core-library';
import {
  type IPropertyPaneConfiguration,
  PropertyPaneTextField
} from '@microsoft/sp-property-pane';
import { BaseClientSideWebPart } from '@microsoft/sp-webpart-base';
import { IReadonlyTheme } from '@microsoft/sp-component-base';

import * as strings from 'SurveyCardWebPartStrings';
import SurveyCard from './components/SurveyCard';
import { ISurveyCardProps } from './components/ISurveyCardProps';

export interface ISurveyCardWebPartProps {
  description: string;
  activeSurveysURL: string;
}

export default class SurveyCardWebPart extends BaseClientSideWebPart<ISurveyCardWebPartProps> {

  //private _isDarkTheme: boolean = false;
  //private _environmentMessage: string = '';

  
public render(): void {
  const element: React.ReactElement<ISurveyCardProps> = React.createElement(
    SurveyCard,
    {
      spHttpClient: this.context.spHttpClient,
      siteUrl: this.context.pageContext.web.absoluteUrl,
      activeSurveysURL: this.properties.activeSurveysURL // use user input here
    }
  );

  ReactDom.render(element, this.domElement);
}

  protected onInit(): Promise<void> {
    return this._getEnvironmentMessage().then(message => {
    //  this._environmentMessage = message;
    });
  }



  private _getEnvironmentMessage(): Promise<string> {
    if (!!this.context.sdks.microsoftTeams) { // running in Teams, office.com or Outlook
      return this.context.sdks.microsoftTeams.teamsJs.app.getContext()
        .then(context => {
          let environmentMessage: string = '';
          switch (context.app.host.name) {
            case 'Office': // running in Office
              environmentMessage = this.context.isServedFromLocalhost ? strings.AppLocalEnvironmentOffice : strings.AppOfficeEnvironment;
              break;
            case 'Outlook': // running in Outlook
              environmentMessage = this.context.isServedFromLocalhost ? strings.AppLocalEnvironmentOutlook : strings.AppOutlookEnvironment;
              break;
            case 'Teams': // running in Teams
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
    if (!currentTheme) {
      return;
    }

   //this._isDarkTheme = !!currentTheme.isInverted;
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
        header: { description: "Survey Card Settings" },
        groups: [
          {
            groupName: "Settings",
            groupFields: [
              PropertyPaneTextField('activeSurveysURL', {
                label: "Active Surveys Page URL",
                placeholder: "/SitePages/ActiveSurveys.aspx"
              })
            ]
          }
        ]
      }
    ]
  };
}

}
