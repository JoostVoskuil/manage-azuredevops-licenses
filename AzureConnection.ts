import * as rm from 'typed-rest-client/RestClient';
import { IHeaders, IHttpClientResponse, IRequestOptions } from "typed-rest-client/Interfaces";
import qs from 'qs';
import { Logger } from './Logger';
import { Settings } from './Settings';

export class AzureConnection {
  private restClient: rm.RestClient;
  private static instance: AzureConnection;

  /**
   * Private constructpr
   * @param { string } endPointUrl The base url of the azure connection
   * @param { string } bearerToken optional bearerToken needed to communicate to Microsoft Graph
   */
  private constructor(endPointUrl: string, bearerToken?: string) {
    const userAgent = 'Azure DevOps';
    const requestHeaders: IHeaders = {};
    requestHeaders['Content-Type'] = 'application/x-www-form-urlencoded';
    if (bearerToken !== undefined) requestHeaders['Authorization'] = `Bearer ${bearerToken}`;

    const requestOptions: IRequestOptions = {
      socketTimeout: 100000,
      allowRetries: true,
      maxRetries: 3,
      headers: requestHeaders,
      ignoreSslError: false
    };

    this.restClient = new rm.RestClient(userAgent, endPointUrl, undefined, requestOptions);
  }

  /**
   * Static initializer for the Azure AAD Connection
   * @param { string } AADGraphEndPointUrl The endpoint url of the azure connection
   * @param { string } AADGraphDirectoryId Directory Id (Tenant) of the Azure AAD
   * @param { string } AADGraphApplicationId Application Id of the registered App in AAD
   * @param { string } AADGraphApplicationToken Secret Token
   * @return { AzureConnection } returns the connection
   */
  static async Initialize(AADGraphEndPointUrl: string, AADGraphDirectoryId: string, AADGraphApplicationId: string, AADGraphApplicationToken: string): Promise<AzureConnection> {
    AzureConnection.instance = new AzureConnection(AADGraphEndPointUrl);
    const postData = {
      client_id: AADGraphApplicationId,
      scope: "https://graph.microsoft.com/.default",
      client_secret: AADGraphApplicationToken,
      grant_type: 'client_credentials'
    };

    const url = `${AADGraphEndPointUrl}/${AADGraphDirectoryId}/oauth2/v2.0/token`
    const response = await AzureConnection.instance.restClient.client.post(url, qs.stringify(postData));
    const body = JSON.parse(await (response.readBody()));
    const msGraphBearerToken = body.access_token;
    AzureConnection.instance = new AzureConnection(AADGraphEndPointUrl, msGraphBearerToken);
    return AzureConnection.instance;
  }

  /**
   * HTTP Get Request
   * @param { string } url Url to get
   * @return { IHttpClientResponse } the response
   */
  static async HttpGet(url: string): Promise<IHttpClientResponse> | undefined {
    try {
      const result = await AzureConnection.instance.restClient.client.get(url);
      return result;
    }
    catch (err) {
      Logger.log(`Error connecting to Azure Active Directory. Error: ${err.message}`);
    }
  }

  /**
 * HTTP Delete Request
 * @param { string } url Url to get
 * @return { IHttpClientResponse } the response
 */
  static async HttpDelete(url: string): Promise<IHttpClientResponse> | undefined {
    if (!Settings.getConfiguration().disableAPIOperations) {
      try {
        const result = await AzureConnection.instance.restClient.client.del(url);
        return result;
      }
      catch (err) {
        Logger.log(`Error connecting to Azure Active Directory. Error: ${err.message}`);
      }
    }
    else {
      Logger.log(`'disableAPIOperations' is set to true, so API operations are disabled`);
    }
  }
}
