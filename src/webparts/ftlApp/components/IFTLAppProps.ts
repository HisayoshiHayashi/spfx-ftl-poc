import type { SPHttpClient } from '@microsoft/sp-http';

export interface IFTLAppProps {
  description: string;
  isServedFromLocalhost: boolean;
  isFullScreen: boolean;
  spHttpClient: Pick<SPHttpClient, 'get' | 'post'>;
  webUrl: string;
}
