import commandLineArgs from 'command-line-args';
import commandLineUsage from 'command-line-usage';

import { AzureConnection } from "./AzureConnection";
import { AzureDevOpsConnection } from "./AzureDevopConnection";
import { IManageLicenseConfiguration, Settings } from "./Settings";
import { AzureDevOpsOrganization as AzureDevOpsOrganization } from "./AzureDevOpsOrganization";
import { Logger, LoggerType } from './Logger';

// Read arguments from console
const optionDefinitions =
    [
        { name: "help", alias: "h", type: Boolean, description: "Display this usage guide." },
        { name: "disableAPIOperations", alias: "t", type: Boolean, description: "Testing disables the real API Patch/Update/Put/delete operations." },
        { name: "azureDevOpsOrganizationName", alias: "o", type: String, description: "The Azure DevOps Organization." },
        { name: "personalAccessToken", alias: "p", type: String, description: "The Personal Access Token." },
        { name: "AADGraphDirectoryId", alias: "d", type: String, description: "The MS Graph Directory Id." },
        { name: "AADGraphApplicationId", alias: "a", type: String, description: "The MS Graph Application Id." },
        { name: "AADGraphApplicationToken", alias: "s", type: String, description: "The MS Graph Token to access Azure Active Directory." },
        { name: "deleteAADUsers", alias: "r", type: String, description: "Delete AAD users when not logged in recently." },
    ];

// check if options are valid
const options = commandLineArgs(optionDefinitions)
const valid = options.help ||
    (
        options.azureDevOpsOrganizationName &&
        options.personalAccessToken &&
        options.AADGraphDirectoryId &&
        options.AADGraphApplicationId &&
        options.AADGraphApplicationToken
    )

// Display help message
if (options.help || !valid) {
    const usage = commandLineUsage([
        {
            header: "Manage Azure DevOps Licenses",
            content: "Manage the Group Entitlements and Licenses of the Azure DevOps organization."
        },
        {
            optionList: optionDefinitions
        }
    ]);
    Logger.log(usage);
}
else {
    // Sets the configuration
    const parameterConfiguration: IManageLicenseConfiguration = {
        disableAPIOperations: options.disableAPIOperations ? undefined : false,
        deleteAADUsers: options.deleteAADUsers ? undefined : false,
        AADGraphApplicationId: options.AADGraphApplicationId,
        AADGraphDirectoryId: options.AADGraphDirectoryId,
        azureDevOpsOrganizationName: options.azureDevOpsOrganizationName
    };

    // Initializes configuration
    Logger.Initialize(LoggerType.Console);
    Settings.InitializeConfiguration(parameterConfiguration);
    processOrganization(options.personalAccessToken, options.AADGraphApplicationToken);
}

/**
 * Processes the licenses in the organization
 * @param { string } azureDevOpsPersonalAccessToken PAT Token for Azure DevOps (requres Project Collection Administration rights)
 * @param { string } AADGraphApplicationToken Secret of Azure AAD App Registration
 */
async function processOrganization(azureDevOpsPersonalAccessToken: string, AADGraphApplicationToken: string): Promise<void> {
    const azureDevOpsApiUrl = Settings.getConfiguration().azureDevOpsApiUrl;
    const oauthEndPoint = Settings.getConfiguration().AADGraphOAuthEndPoint;
    const msGraphDirectoryId = Settings.getConfiguration().AADGraphDirectoryId;
    const msGraphApplicationId = Settings.getConfiguration().AADGraphApplicationId;

    AzureConnection.Initialize(oauthEndPoint, msGraphDirectoryId, msGraphApplicationId, AADGraphApplicationToken);
    AzureDevOpsConnection.Initialize(azureDevOpsApiUrl, azureDevOpsPersonalAccessToken);
    const azureDevOpsOrganization: AzureDevOpsOrganization = await AzureDevOpsOrganization.CreateAsync(Settings.getConfiguration().azureDevOpsOrganizationName);
    await azureDevOpsOrganization.processUserAccounts();
}