import { SPHttpClient, MSGraphClientFactory } from '@microsoft/sp-http';

export interface IManagerDashboardProps {
  siteUrl: string;
  userDisplayName: string;
  spHttpClient: SPHttpClient;
  msGraphClientFactory: MSGraphClientFactory;
  /** Access code for the CRM module, set per site in the web part property pane. */
  crmAccessCode: string;
}
