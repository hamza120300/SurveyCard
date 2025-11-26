import { SPHttpClient } from "@microsoft/sp-http";

export interface ISurveyCardProps {
  // description: string;
  // isDarkTheme: boolean;
  // environmentMessage: string;
  // hasTeamsContext: boolean;
  // userDisplayName: string;

  spHttpClient: SPHttpClient;
  siteUrl: string;
  activeSurveysURL: string;

}
