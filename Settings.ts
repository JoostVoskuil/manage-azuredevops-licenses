import { addDays } from "date-fns";
import { Logger } from "./Logger";

export class Settings {
    private _configuration: IManageLicenseConfiguration;
    private static instance: Settings;

    /**
    * Private constructpr
    * @param { IManageLicenseConfiguration } configuration The specified configration to hold
    */
    private constructor(configuration: IManageLicenseConfiguration) {
        this._configuration = configuration;
    }

    /**
    * Static getter for the configuration
    * @return { IManageLicenseConfiguration } returns the configuration
    */
    public static getConfiguration(): IManageLicenseConfiguration {
        if (Settings.instance) {
            return Settings.instance._configuration;
        }
        throw Error("Configuration was not initialized");
    }

    /**
     * Static initializer for the configuration
     * @param { IManageLicenseConfiguration } givenConfiguration The provided (input) configration to hold
     * @return { Settings } returns the connection
     */
    public static InitializeConfiguration(givenConfiguration: IManageLicenseConfiguration): Settings {
        /* eslint-disable @typescript-eslint/no-var-requires */
        const savedConfiguration: IManageLicenseConfiguration = require('./settings.json');
        // merge configuration
        const configuration = Settings.mergeConfiguration(givenConfiguration, savedConfiguration);
        configuration.azureDevOpsApiUrl = `${configuration.azureDevOpsVSAexBaseUrl}/${configuration.azureDevOpsOrganisationName}/`;
        Settings.instance = new Settings(configuration);
        configuration.dateDeleteBeforeLastAccessDate = addDays(new Date(), -1 * Settings.instance._configuration.numberOfDaysNotLoggedInForDeletionOfUser);
        configuration.dateMakeStakeholderBeforeLastAccessDate = addDays(new Date(), -1 * Settings.instance._configuration.numberOfDaysNotLoggedInToBecomeStakeholder);
        configuration.dateProcessAfterCreatedOnDate = addDays(new Date(), -1 * Settings.instance._configuration.numberOfDaysToWaitAfterUserIsCreated);

        Logger.log(`Delete users that have not logged in before: '${configuration.dateDeleteBeforeLastAccessDate.toISOString()}'`);
        Logger.log(`Make users stakeholder that have not logged before: '${configuration.dateMakeStakeholderBeforeLastAccessDate.toISOString()}'`);
        Logger.log(`Do not delete or make users stakeholder when creation date is after: '${configuration.dateProcessAfterCreatedOnDate.toISOString()}'`);

        return Settings.instance;
    }

    /**
     * Merges two configuration inputs
     * @param { IManageLicenseConfiguration } leadingConfiguration The provided (input) configration to hold
     * @param { IManageLicenseConfiguration } subConfiguration the default configuration
     */
    private static mergeConfiguration(leadingConfiguration: IManageLicenseConfiguration, subConfiguration: IManageLicenseConfiguration): IManageLicenseConfiguration {
        for (const key in subConfiguration) {
            if (leadingConfiguration[key] === undefined) {
                leadingConfiguration[key] = subConfiguration[key];
            }
        }
        return leadingConfiguration
    }
}

// The Interface
export interface IManageLicenseConfiguration {
    dateDeleteBeforeLastAccessDate?: Date;
    dateMakeStakeholderBeforeLastAccessDate?: Date;
    dateProcessAfterCreatedOnDate?: Date;
    numberOfDaysNotLoggedInToBecomeStakeholder?: number;
    numberOfDaysNotLoggedInForDeletionOfUser?: number;
    numberOfDaysToWaitAfterUserIsCreated?: number;
    groupEntitlementPrefix?: string;
    groupEntitlementPostfix?: string;
    AADGraphApplicationId?: string;
    AADGraphDirectoryId?: string;
    AADGraphOAuthEndPoint?: string;
    excludedWordsInUserNames?: string;
    azureDevOpsApiUrl?: string;
    azureDevOpsVSAexBaseUrl?: string;
    azureDevOpsOrganisationName?: string;
    disableAPIOperations?: boolean;
}