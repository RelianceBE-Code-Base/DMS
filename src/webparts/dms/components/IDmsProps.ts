import { WebPartContext } from "@microsoft/sp-webpart-base";
import { PageContext } from "@microsoft/sp-page-context";
import { IDmsWebPartProps } from "../DmsWebPart";
import { ServiceScope } from "@microsoft/sp-core-library";



export interface IDmsProps {
  description: string;
  // isDarkTheme: boolean;
  // environmentMessage: string;
  // hasTeamsContext: boolean;
  userDisplayName: string;
  context: WebPartContext;
  siteUrl: string,
  folderUrl: string
  //serviceScope: ServiceScope;

  // siteID: string;
  // channelName: string;
}
