import type { SPHttpClient } from '@microsoft/sp-http';

export interface IFTLAppProps {
  description: string;
  isServedFromLocalhost: boolean;
  isFullScreen: boolean;
  userDisplayName?: string;
  userEmail?: string;
  userLoginName?: string;
  spHttpClient: Pick<SPHttpClient, 'get' | 'post'>;
  webUrl: string;
}
