import { SPHttpClient } from "@microsoft/sp-http";

export interface IHelloWorldProps {
  description: string;
  spHttpClient: SPHttpClient;
  siteUrl: string;
}
