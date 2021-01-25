import { SPHttpClient } from "@microsoft/sp-http";
export interface ISpfxBoilerplateCodeProps {
  siteUrl: string;
  spHttpClient: SPHttpClient;
  context: any | null;
}
