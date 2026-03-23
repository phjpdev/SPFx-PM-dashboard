import { SPHttpClient } from '@microsoft/sp-http';

export interface IStaffDashboardProps {
  siteUrl: string;
  userDisplayName: string;
  spHttpClient: SPHttpClient;
}
