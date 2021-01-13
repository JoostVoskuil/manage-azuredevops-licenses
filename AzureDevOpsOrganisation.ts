import { plainToClass } from "class-transformer";
import 'reflect-metadata';
import { AzureDevOpsConnection } from "./AzureDevopConnection";
import { GroupEntitlement } from "./GroupEntitlement";
import { IGroupEntitlementResponse, IUserEntitlementResponse } from "./IAPIResponses";
import { License, LicenseRule } from "./LicenseRule";
import { Settings } from "./Settings";
import { UserEntitlement } from "./UserEntitlement";
import { Group } from "./Group";
import { Logger } from "./Logger";

export class AzureDevOpsOrganisation {
    private groupEntitlements: GroupEntitlement[];
    private userEntitlements: UserEntitlement[];
    private organisationName: string;

    /**
     * Async constructor. Fetches all the groups and user entitlements
     * @param { string } organisationName The name of the organisation
     * @return { AzureDevOpsOrganisation } this object
     */
    public static async CreateAsync(organisationName: string): Promise<AzureDevOpsOrganisation> {
        const thisOrganisation = new AzureDevOpsOrganisation();
        thisOrganisation.organisationName = organisationName;

        await thisOrganisation.getAllGroupEntitlements();
        await thisOrganisation.getAllUserEntitlements();
        await thisOrganisation.createNeededGroupEntitlements();
        return thisOrganisation;
    }

    /**
     * Processes all the userEntitlements
     * @return { AzureDevOpsOrganisation } this object
     */
    public async processUserAccounts(): Promise<this> {
        if (!this.userEntitlements) return this;
        for (const userEntitlement of this.userEntitlements) {
            Logger.log(`Process '${userEntitlement.user.displayName}' with lastAccessDate '${userEntitlement.lastAccessedDate}' and license '${userEntitlement.accessLevel.licenseDisplayName}'`);

            // Delete user from the organization if the user did not login in the last period or user is not active in AAD
            if (await userEntitlement.shouldUserBeDeleted() === true) {
                await userEntitlement.delete();
                continue;
            }

            // Make sure that the user is allocated to a grouprule and remove direct assignment
            if (await userEntitlement.shouldUserBeAddedToAGroupEntitlement() === true) {
                await userEntitlement.addUserToAGroupEntitlementBasedOnLicense(this.groupEntitlements);
            }

            // Make user Stakeholder if it should be made stakeholder
            if (await userEntitlement.shouldUserBeStakeholder() === true) {
                await userEntitlement.removeFromAllLicenseGroupEntitlements()
                await userEntitlement.addToGroupEntitlement(this.groupEntitlements, License.Stakeholder);
                await userEntitlement.removeDirectAssignment();
            }
        }
        await this.reevaluateGroupEntitlementRules();
        return this;
    }

    /**
     * Creates all the needed Group Entitlements for License management
     * @return { AzureDevOpsOrganisation } this object
     */
    private async createNeededGroupEntitlements(): Promise<this> {
        for (const licenseDisplayName of Object.values(License)) {
            // groupnames cannot contain + symbols
            const groupEntitlementDisplayName = (`${Settings.getConfiguration().groupEntitlementPrefix}${licenseDisplayName}${Settings.getConfiguration().groupEntitlementPostfix}`).replace('+', '-');
            if (!(this.groupEntitlements.find(x => x.group.displayName === groupEntitlementDisplayName))) {
                await this.addGroupEntitlement(groupEntitlementDisplayName, licenseDisplayName as License);
            }
        }
        await this.getAllGroupEntitlements();
        return this;
    }

    /**
     * Gets all the user Entitlements of this organisation
     * @return { AzureDevOpsOrganisation } this object
     */
    private async getAllUserEntitlements(): Promise<this> {
        const excludedWords = Settings.getConfiguration().excludedWordsInUserNames;
        // Use old version of the API. Newer versions don't support top/continuationtoken
        const response = await AzureDevOpsConnection.get<IUserEntitlementResponse>(`_apis/userentitlements?api-version=4.1-preview.1&top=10000&select=Grouprules`);
        const userEntitlements = response.result.value.filter(x => !this.containsAny(x.user.displayName, excludedWords));
        Logger.log(`Fetched ${userEntitlements.length} users from Azure DevOps Organisation '${this.organisationName}'`);
        this.userEntitlements = plainToClass(UserEntitlement, userEntitlements);
        return this;
    }

    /**
     * Gets all the Group Entitlements of this organisation
     * @return { AzureDevOpsOrganisation } this object
     */
    private async getAllGroupEntitlements(): Promise<this> {
        const response = await AzureDevOpsConnection.get<IGroupEntitlementResponse>(`_apis/groupentitlements?api-version=6.0-preview.1`);
        this.groupEntitlements = plainToClass(GroupEntitlement, response.result.value);
        Logger.log(`Fetched ${response.result.value.length} Group Entitlements from Azure DevOps Organisation ${this.organisationName}`);
        return this;
    }

    /**
     * Adds a new Group Entitlement
     * @param { string } groupEntitlementDisplayName Display name of the group
     * @param { License } license The License
     * @return { AzureDevOpsOrganisation } this object
     */
    private async addGroupEntitlement(groupEntitlementDisplayName: string, license: License): Promise<this> {
        let licensingSource = '1';
        let accountLicenseType = '0';
        let msdnLicenseType = '0';

        switch (license) {
            case License.Basic: {
                accountLicenseType = '2';
                break;
            }
            case License.Basic_TestPlans: {
                accountLicenseType = '4';
                break;
            }
            case License.Stakeholder: {
                accountLicenseType = '5';
                break;
            }
            case License.VisualStudioSubscriber: {
                licensingSource = '2';
                msdnLicenseType = '1';
                break;
            }
        }

        const group: Group = {
            displayName: groupEntitlementDisplayName,
            origin: 'vsts',
            subjectKind: 'group'
        };
        const licenseRule: LicenseRule = {
            licensingSource: licensingSource,
            accountLicenseType: accountLicenseType,
            msdnLicenseType: msdnLicenseType,
            licenseDisplayName: license.toString(),
            status: '0',
            statusMessage: '',
            assignmentSource: '1'
        };

        const groupEntitlement = new GroupEntitlement() 
            .withGroup(group)
            .withLicenseRule(licenseRule);

        await groupEntitlement.create();
        Logger.log(`Added GroupEntitlement '${groupEntitlementDisplayName}'.`);
        return this;
    }

    /**
     * Re-evaluates the group entitlements.
     * Note: uses an undocumented API
     * @return { AzureDevOpsOrganisation } this object
     */
    private async reevaluateGroupEntitlementRules(): Promise<this> {
        Logger.log(`Re-evaluating Rules.`);
        await AzureDevOpsConnection.create(`_apis/MEMInternal/GroupEntitlementUserApplication?ruleOption=0&api-version=5.0-preview.1`, undefined);
        Logger.log(`Re-evaluated Rules.`);
        return this;
    }

    /** Returns true if any of the given words from the dictionary are part of the searchstring
     * @param searchString 
     * @param dictionary 
     * @return { boolean } true if found, false if not found
     */
    private containsAny(searchString, dictionary) {
        const dictionaryArray = dictionary.split(',');
        for (let i = 0; i != dictionaryArray.length; i++) {
            const substring = dictionaryArray[i];
            if (searchString.indexOf(substring) != - 1) {
                return true;
            }
        }
        return false;
    }
}
