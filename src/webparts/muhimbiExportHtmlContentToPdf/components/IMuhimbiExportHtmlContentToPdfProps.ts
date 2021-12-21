import { WebPartContext } from "@microsoft/sp-webpart-base";
import { HttpClient } from "@microsoft/sp-http";

export interface IMuhimbiExportHtmlContentToPdfProps {
  description: string;
  context: WebPartContext;
  httpClient: HttpClient;
  apiKey: string;
  apiUrl: string;
}
