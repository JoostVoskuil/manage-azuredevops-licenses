import { plainToClass } from "class-transformer";
import 'reflect-metadata';
import { AzureDevOpsConnection } from "./AzureDevopConnection";
import { GroupEntitlement } from "./GroupEntitlement";
import { License, LicenseRule } from "./LicenseRule";
import { Settings } from "./Settings";
import { IUserEntitlementResponse } from "./IAPIResponses";
import { Logger } from "./Logger";

export class UserEntitlement {
    public id: string;
    public user: IUser;
    public accessLevel: LicenseRule;
    public lastAccessedDate?: Date;
    public dateCreated: Date;
    public GroupAssignments?: GroupEntitlement[];

    /**
     * Adds the User Entitlement to a Group Entitlement
     * @param { GroupEntitlement[] } availableGroupEntitlements the organization Group Entitlements
     * @param { License } license the license where the user needs to be made member of 
     * @return { UserEntitlement } this
     */
    public async addToGroupEntitlement(availableGroupEntitlements: GroupEntitlement[], license: License): Promise<this> {
        const groupEntitlementDisplayName = `${Settings.getConfiguration().groupEntitlementPrefix}${license.toString()}${Settings.getConfiguration().groupEntitlementPostfix}`
        const groupEntitlement = availableGroupEntitlements.find(x => x.group.displayName === groupEntitlementDisplayName);
        groupEntitlement.addMember(this);
        await this.refreshGroupEntitlements();
        Logger.log(`Added User '${this.user.displayName}' to group entitlement '${groupEntitlement.group.displayName}'.`);
        return this;
    }

    /**
     * Removes the Direct license assignment
     * @return { UserEntitlement } this
     */
    public async removeDirectAssignment(): Promise<this> {
        if (!this.accessLevel.licenseDisplayName.includes('Visual Studio')) {
            await AzureDevOpsConnection.create(`_apis/MEMInternal/RemoveExplicitAssignment?ruleOption=0&api-version=5.0-preview.1`, [this.id]);
        }

        Logger.log(`Removed Direct License Assignment for User '${this.user.displayName}'.`);
        return this;
    }

    /**
     * Deletes the user entitlement
     * @return { undefined } undefined
     */
    public async deleteFromAzureDevOps(): Promise<undefined> {
        await AzureDevOpsConnection.del<IUserEntitlementResponse>(`_apis/userentitlements/${this.id}?api-version=6.1-preview.3`);

        Logger.log(`Deleted User '${this.user.displayName}' from organization.`);
        return undefined;
    }

    /**
     * Removes the user from all the group entitlement (of this application)
     * @return { UserEntitlement } this
     */
    public async removeFromAllLicenseGroupEntitlements(): Promise<this> {
        for (const licenseDisplayName of Object.values(License)) {
            // groupnames cannot contain + symbols
            const groupEntitlementDisplayName = (`${Settings.getConfiguration().groupEntitlementPrefix}${licenseDisplayName}${Settings.getConfiguration().groupEntitlementPostfix}`).replace('+', '-');
            const groupEntitlement = this.GroupAssignments.find(x => x.group.displayName === groupEntitlementDisplayName);
            if (groupEntitlement) {
                groupEntitlement.removeMember(this);
            }
        }
        this.refreshGroupEntitlements();
        return this;
    }

    /**
     * Refreshes the Group Entitlements for this user
     * @return { UserEntitlement } this
     */
    private async refreshGroupEntitlements(): Promise<this> {
        const myself = await AzureDevOpsConnection.get<UserEntitlement>(`_apis/userentitlements/${this.id}?api-version=6.1-preview.3`);
        this.GroupAssignments = plainToClass(GroupEntitlement, myself.result.GroupAssignments);
        return this;
    }

    /**
     * Checks if user should be added to a group entitlement
     * @return { boolean } true if use has a direct assignment
     */
    public async shouldUserBeAddedToAGroupEntitlement(): Promise<boolean> {
        if (this.accessLevel.assignmentSource !== 'groupRule') {
            return true;
        }
        return false;
    }

    /**
     * Checks if user should be made a stakeholder
     * @return { boolean } true if this user should become a stakeholder
     */
    public async shouldUserBeStakeholder(): Promise<boolean> {
        // Make Stakeholder when user did not login recently
        if (this.lastAccessedDate < Settings.getConfiguration().dateMakeStakeholderBeforeLastAccessDate &&
            this.dateCreated < Settings.getConfiguration().dateProcessAfterCreatedOnDate &&
            this.accessLevel.licenseDisplayName !== 'Stakeholder' &&
            !this.accessLevel.licenseDisplayName.includes('Visual Studio')) {
            return true;
        }
        return false;
    }

    /**
     * Adds user to the correct group Entitlement
     * @param { GroupEntitlement[] } groupEntitlements the group entitlements of this organization (provided by this application)
     * @return { UserEntitlement } this
     */
    public async addUserToAGroupEntitlementBasedOnLicense(groupEntitlements: GroupEntitlement[]): Promise<this> {
        if (this.accessLevel.licenseDisplayName === License.Basic) {
            await this.addToGroupEntitlement(groupEntitlements, License.Basic);
            await this.removeDirectAssignment();
        }
        else if (this.accessLevel.licenseDisplayName === License.Stakeholder) {
            await this.addToGroupEntitlement(groupEntitlements, License.Stakeholder);
            await this.removeDirectAssignment();
        }
        else if (this.accessLevel.licenseDisplayName === License.Basic_TestPlans) {
            await this.addToGroupEntitlement(groupEntitlements, License.Basic_TestPlans);
            await this.removeDirectAssignment();
        }
        // Process special group for visual studio users, do not remove direct license
        else if (this.accessLevel.licenseDisplayName.includes('Visual Studio')) {
            await this.addToGroupEntitlement(groupEntitlements, License.VisualStudioSubscriber);
        }
        return this;
    }

    /**
     * Checks if user should be deleted
     * @return { boolean } true if this user should be deleted from Azure DevOps
     */
    public async shouldUserBeDeletedFromAzureDevOps(): Promise<boolean> {
        // Delete user from the organization if the user did not login in the last period.
        if (this.lastAccessedDate < Settings.getConfiguration().dateDeleteBeforeLastAccessDate &&
            this.dateCreated < Settings.getConfiguration().dateProcessAfterCreatedOnDate) {
            return true;
        }
        return false;
    }
}

// The User interface comming from Azure DevOps
interface IUser {
    principalName: string;
    displayName: string;
    originId: string;
    descriptor: string;
}