import commandLineArgs from 'command-line-args';
import commandLineUsage from 'command-line-usage';

import { AzureConnection } from "./AzureConnection";
import { AzureDevOpsConnection } from "./AzureDevopConnection";
import { IManageLicenseConfiguration, Settings } from "./Settings";
import { AzureDevOpsOrganisation } from "./AzureDevOpsOrganisation";
import { Logger, LoggerType } from './Logger';

// Read arguments from console
const optionDefinitions =
    [
        { name: "help", alias: "h", type: Boolean, description: "Display this usage guide." },
        { name: "disableAPIOperations", alias: "t", type: Boolean, description: "Testing disables the real API Patch/Update/Put/delete operations." },
        { name: "azureDevOpsOrganisationName", alias: "o", type: String, description: "The Azure DevOps Organisation." },
        { name: "personalAccessToken", alias: "p", type: String, description: "The Personal Access Token." },
        { name: "AADGraphDirectoryId", alias: "d", type: String, description: "The MS Graph Directory Id." },
        { name: "AADGraphApplicationId", alias: "a", type: String, description: "The MS Graph Application Id." },
        { name: "AADGraphApplicationToken", alias: "s", type: String, description: "The MS Graph Token to access Azure Active Directory." },
    ];

// check if options are valid
const options = commandLineArgs(optionDefinitions)
const valid = options.help ||
    (
        options.azureDevOpsOrganisationName &&
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
            content: "Manage the Group Entitlements and Licenses of the Azure DevOps organisation."
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
        AADGraphApplicationId: options.AADGraphApplicationId,
        AADGraphDirectoryId: options.AADGraphDirectoryId,
        azureDevOpsOrganisationName: options.azureDevOpsOrganisationName
    };

    // Initializes configuration
    Logger.Initialize(LoggerType.Console);
    Settings.InitializeConfiguration(parameterConfiguration);
    processOrganisation(options.personalAccessToken, options.AADGraphApplicationToken);
}

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