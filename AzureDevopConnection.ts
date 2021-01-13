import * as rm from 'typed-rest-client/RestClient';
import { IHeaders, IRequestOptions } from 'typed-rest-client/Interfaces';
import { Settings } from './Settings';
import { Logger } from './Logger';

export class AzureDevOpsConnection {
  private restClient: rm.RestClient;
  private static instance: AzureDevOpsConnection;

  /**
   * Private constructpr
   * @param { string } organisationUrl The url of the organisation
   * @param { string } token PAT token
   */
  private constructor(organisationUrl: string, token: string) {
    const userAgent = 'Azure DevOps';
    const requestHeaders: IHeaders = {};
    requestHeaders['Content-Type'] = 'application/json';
    requestHeaders['Authorization'] = `Basic ${Buffer.from(`PAT:${token}`).toString('base64')}`
    const requestOptions: IRequestOptions = {
      socketTimeout: 100000,
      allowRetries: true,
      maxRetries: 10,
      headers: requestHeaders,
    };

    this.restClient = new rm.RestClient(userAgent, organisationUrl, undefined, requestOptions);
  }
  /**
   * Static initializer for the Azure DevOps Connection
   * @param { string } organisationUrl The url of the organisation
   * @param { string } personalAccessToken Full Access (Project Collection Administrator) PAT
   * @return { AzureDevOpsConnection } returns the connection
   */
  static Initialize(organisationUrl: string, personalAccessToken: string): AzureDevOpsConnection {
    AzureDevOpsConnection.instance = new AzureDevOpsConnection(organisationUrl, personalAccessToken);
    return AzureDevOpsConnection.instance;
  }

  /**
   * Creates a resource
   * @param { string } resource resource Url
   * @param { any } createObject Object to be created
   * @return { rm.IRestResponse<T> } returns the RestResponse
   */
  /* eslint-disable @typescript-eslint/no-explicit-any, @typescript-eslint/explicit-module-boundary-types */
  static async create<T>(resource: string, createObject: any): Promise<rm.IRestResponse<T>> {
    if (!Settings.getConfiguration().disableAPIOperations) {
      try {
        const result = await AzureDevOpsConnection.instance.restClient.create<T>(resource, createObject);
        return result;
      }
      catch (err) {
        Logger.log(`Error sending create to ${resource} to Azure DevOps API. Error: ${err.message}`);
      }
    }
    else {
      Logger.log(`'disableAPIOperations' is set to true, so API operations are disabled`);
    }
  }

  /**
   * Replaces a resource
   * @param { string } resource resource Url
   * @param { any } createObject Object to be created
   * @return { rm.IRestResponse<T> } returns the RestResponse
   */
  /* eslint-disable @typescript-eslint/no-explicit-any, @typescript-eslint/explicit-module-boundary-types */
  static async replace<T>(resource: string, createObject: any): Promise<rm.IRestResponse<T>> {
    if (!Settings.getConfiguration().disableAPIOperations) {
      try {
        const result = await AzureDevOpsConnection.instance.restClient.replace<T>(resource, createObject);
        return result;
      }
      catch (err) {
        Logger.log(`Error sending replace to ${resource} to Azure DevOps API. Error: ${err.message}`);
      }
    }
    else {
      Logger.log(`'disableAPIOperations' is set to true, so API operations are disabled`);
    }
  }

  /**
   * Updates a resource
   * @param { string } resource resource Url
   * @param { any } createObject Object to be created
   * @return { rm.IRestResponse<T> } returns the RestResponse
   */
  /* eslint-disable @typescript-eslint/no-explicit-any, @typescript-eslint/explicit-module-boundary-types */
  static async update<T>(resource: string, createObject: any): Promise<rm.IRestResponse<T>> {
    if (!Settings.getConfiguration().disableAPIOperations) {

      const requestHeaders: IHeaders = {};
      requestHeaders['Content-Type'] = 'application/json-patch+json';

      const requestOptions: rm.IRequestOptions = {
        additionalHeaders: requestHeaders,
      };

      try {
        const result = await AzureDevOpsConnection.instance.restClient.update<T>(resource, createObject, requestOptions);
        return result;
      }
      catch (err) {
        Logger.log(`Error sending update to ${resource} to Azure DevOps API. Error: ${err.message}`);
      }
    }
    else {
      Logger.log(`'disableAPIOperations' is set to true, so API operations are disabled`);
    }
  }

  /**
   * Gets a resource
   * @param { string } resource resource Url
   * @return { rm.IRestResponse<T> } returns the RestResponse
   */
  static async get<T>(resource: string): Promise<rm.IRestResponse<T>> {
    try {
      const result = await AzureDevOpsConnection.instance.restClient.get<T>(resource);
      if (result.result === null) throw new Error("Not found");
      return result;
    }
    catch (err) {
      Logger.log(`Error sending get to ${resource} to Azure DevOps API. Error: ${err.message}`);
    }
  }

  /**
   * Deletes a resource
   * @param { string } resource resource Url
   * @return { rm.IRestResponse<T> } returns the RestResponse
   */
  static async del<T>(resource: string): Promise<rm.IRestResponse<T>> {
    if (!Settings.getConfiguration().disableAPIOperations) {
      try {
        const result = await AzureDevOpsConnection.instance.restClient.del<T>(resource);
        return result;
      }
      catch (err) {
        Logger.log(`Error sending delete to ${resource} to Azure DevOps API. Error: ${err.message}`);
      }
    }
    else {
      Logger.log(`'disableAPIOperations' is set to true, so API operations are disabled`);
    }
  }
}
