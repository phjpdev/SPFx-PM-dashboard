import { SPHttpClient, MSGraphClientFactory } from '@microsoft/sp-http';

export interface IManagerDashboardProps {
  siteUrl: string;
  userDisplayName: string;
  spHttpClient: SPHttpClient;
  msGraphClientFactory: MSGraphClientFactory;
}
