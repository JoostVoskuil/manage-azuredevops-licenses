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
import { GraphUser } from "./GraphUser";

export class AzureDevOpsOrganization {
    private groupEntitlements: GroupEntitlement[];
    private userEntitlements: UserEntitlement[];
    private organizationName: string;

    /**
     * Async constructor. Fetches all the groups and user entitlements
     * @param { string } organizationName The name of the organization
     * @return { AzureDevOpsOrganization } this object
     */
    public static async CreateAsync(organizationName: string): Promise<AzureDevOpsOrganization> {
        const thisorganization = new AzureDevOpsOrganization();
        thisorganization.organizationName = organizationName;

        await thisorganization.getAllGroupEntitlements();
        await thisorganization.getAllUserEntitlements();
        await thisorganization.createNeededGroupEntitlements();
        return thisorganization;
    }

    /**
     * Processes all the userEntitlements
     * @return { AzureDevOpsOrganization } this object
     */
    public async processUserAccounts(): Promise<this> {
        if (!this.userEntitlements) return this;
        for (const userEntitlement of this.userEntitlements) {
            Logger.log(`Process '${userEntitlement.user.displayName}' with Azure DevOps lastAccessDate '${userEntitlement.lastAccessedDate}' and license '${userEntitlement.accessLevel.licenseDisplayName}'`);

            // Get AAD User
            const aadUser = await new GraphUser().getAADGraphUser(userEntitlement.user.principalName);
            // Delete user from Azure DevOps when user not active in AAD
            if (await aadUser.isActiveInAAD() === false) {
                await userEntitlement.deleteFromAzureDevOps();
                continue;
            }
            // Delete user from AAD and Azure DevOps when user did not login in the last period
            else if (Settings.getConfiguration().deleteAADUsers && await aadUser.shouldBeDeletedFromAAD() === true) {
                await aadUser.deleteFromAAD();
                await userEntitlement.deleteFromAzureDevOps();
                continue;          
            }
            // Delete user from the organization if the user did not login in the last period
            else if (await userEntitlement.shouldUserBeDeletedFromAzureDevOps() === true) {
                await userEntitlement.deleteFromAzureDevOps();
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
     * @return { AzureDevOpsOrganization } this object
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
     * Gets all the user Entitlements of this organization
     * @return { AzureDevOpsOrganization } this object
     */
    private async getAllUserEntitlements(): Promise<this> {
        const excludedWords = Settings.getConfiguration().excludedWordsInUserNames;
        const excludedUPNs = Settings.getConfiguration().excludedUPNs;
        // Use old version of the API. Newer versions don't support top/continuationtoken
        const response = await AzureDevOpsConnection.get<IUserEntitlementResponse>(`_apis/userentitlements?api-version=4.1-preview.1&top=10000&select=Grouprules`);
        // filter excluded words
        let userEntitlements = response.result.value.filter(x => !this.containsAny(x.user.displayName, excludedWords));
        // filter excluded upn's
        userEntitlements = response.result.value.filter(x => !this.containsAny(x.user.displayName, excludedUPNs));
        Logger.log(`Fetched ${userEntitlements.length} users from Azure DevOps organization '${this.organizationName}'`);
        this.userEntitlements = plainToClass(UserEntitlement, userEntitlements);
        return this;
    }

    /**
     * Gets all the Group Entitlements of this organization
     * @return { AzureDevOpsOrganization } this object
     */
    private async getAllGroupEntitlements(): Promise<this> {
        const response = await AzureDevOpsConnection.get<IGroupEntitlementResponse>(`_apis/groupentitlements?api-version=6.0-preview.1`);
        this.groupEntitlements = plainToClass(GroupEntitlement, response.result.value);
        Logger.log(`Fetched ${response.result.value.length} Group Entitlements from Azure DevOps organization ${this.organizationName}`);
        return this;
    }

    /**
     * Adds a new Group Entitlement
     * @param { string } groupEntitlementDisplayName Display name of the group
     * @param { License } license The License
     * @return { AzureDevOpsOrganization } this object
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
     * @return { AzureDevOpsOrganization } this object
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
    private containsAny(searchString: string, dictionary: string): boolean {
        const dictionaryArray = dictionary.split(',');
        for (let i = 0; i != dictionaryArray.length; i++) {
            const substring: string = dictionaryArray[i];
            if (searchString.toLowerCase().indexOf(substring.toLowerCase()) != - 1) {
                return true;
            }
        }
        return false;
    }
}