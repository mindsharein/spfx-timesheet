import { WebPartContext } from "@microsoft/sp-webpart-base";

export interface ITimeSheetProps {
  wpContext: WebPartContext;
  description: string;
  isDarkTheme: boolean;
  environmentMessage: string;
  hasTeamsContext: boolean;
  userDisplayName: string;
}
