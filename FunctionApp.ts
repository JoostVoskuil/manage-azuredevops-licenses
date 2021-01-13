import { AzureConnection } from "./AzureConnection";
import { AzureDevOpsConnection } from "./AzureDevopConnection";
import { IManageLicenseConfiguration, Settings } from "./Settings";
import { AzureDevOpsOrganisation } from "./AzureDevOpsOrganisation";
import { AzureFunction, Context } from "@azure/functions"
import { Logger, LoggerType } from "./Logger";

/* eslint-disable @typescript-eslint/no-explicit-any */
const AzureDevOpsLicensesTrigger: AzureFunction = async function (context: Context): Promise<void> {
    Logger.Initialize(LoggerType.FunctionContext, context);

    // Get configuration from environmental vars
    const environmentConfiguration: IManageLicenseConfiguration = {
        numberOfDaysNotLoggedInToBecomeStakeholder: process.env['numberOfDaysNotLoggedInToBecomeStakeholder'] ? Number(process.env['numberOfDaysNotLoggedInToBecomeStakeholder']): undefined,
        numberOfDaysNotLoggedInForDeletionOfUser: process.env['numberOfDaysNotLoggedInForDeletionOfUser'] ? Number(process.env['numberOfDaysNotLoggedInForDeletionOfUser']):undefined,
        numberOfDaysToWaitAfterUserIsCreated: process.env['numberOfDaysToWaitAfterUserIsCreated'] ? Number(process.env['numberOfDaysToWaitAfterUserIsCreated']):undefined,
        AADGraphApplicationId: process.env['AADGraphApplicationId'],
        AADGraphDirectoryId: process.env['AADGraphDirectoryId'],
        AADGraphOAuthEndPoint: process.env['AADGraphOAuthEndPoint'],
        excludedWordsInUserNames: process.env['excludedWordsInUserNames'],
        azureDevOpsApiUrl: process.env['azureDevOpsApiUrl'],
        azureDevOpsVSAexBaseUrl: process.env['azureDevOpsVSAexBaseUrl'],
        azureDevOpsOrganisationName: process.env['azureDevOpsOrganisationName'],
        disableAPIOperations: process.env['disableAPIOperations'] ? Boolean(process.env['disableAPIOperations']): false,
    };

    if (!environmentConfiguration.azureDevOpsOrganisationName ||
        !environmentConfiguration.AADGraphDirectoryId ||
        !environmentConfiguration.AADGraphApplicationId) {
        context.log.error('azureDevOpsOrganisationName or AADGraphDirectoryId or AADGraphApplicationId not set');
        return;
    }

    Settings.InitializeConfiguration(environmentConfiguration);
    // Read secrets
    const azureDevOpsPersonalAccessToken = process.env['azureDevOpsPersonalAccessToken'];
    const AADGraphApplicationToken = process.env['AADGraphApplicationToken'];

    if (!azureDevOpsPersonalAccessToken || !AADGraphApplicationToken) {
        context.log.error('azureDevOpsPersonalAccessToken or AADGraphApplicationToken not set');
        return;
    }

    // Process Organisation
    await processOrganisation(azureDevOpsPersonalAccessToken, AADGraphApplicationToken);
};

/**
 * Processes the licenses in the organisation
 * @param { string } azureDevOpsPersonalAccessToken PAT Token for Azure DevOps (requres Project Collection Administration rights)
 * @param { string } AADGraphApplicationToken Secret of Azure AAD App Registration
 */
async function processOrganisation(azureDevOpsPersonalAccessToken: string, AADGraphApplicationToken: string): Promise<void> {
    const azureDevOpsApiUrl = Settings.getConfiguration().azureDevOpsApiUrl;
    const oauthEndPoint = Settings.getConfiguration().AADGraphOAuthEndPoint;
    const msGraphDirectoryId = Settings.getConfiguration().AADGraphDirectoryId;
    const msGraphApplicationId = Settings.getConfiguration().AADGraphApplicationId;

    AzureConnection.Initialize(oauthEndPoint, msGraphDirectoryId, msGraphApplicationId, AADGraphApplicationToken);
    AzureDevOpsConnection.Initialize(azureDevOpsApiUrl, azureDevOpsPersonalAccessToken);
    const azureDevOpsOrganisation: AzureDevOpsOrganisation = await AzureDevOpsOrganisation.CreateAsync(Settings.getConfiguration().azureDevOpsOrganisationName);
    await azureDevOpsOrganisation.processUserAccounts();
}

export default AzureDevOpsLicensesTrigger;