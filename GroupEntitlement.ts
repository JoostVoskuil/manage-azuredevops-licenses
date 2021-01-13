import { AzureDevOpsConnection } from "./AzureDevopConnection";
import { LicenseRule } from "./LicenseRule";
import { UserEntitlement } from "./UserEntitlement";
import { Group } from "./Group";
import { Logger } from "./Logger";

export class GroupEntitlement {
    public id?: string;
    public group?: Group;
    public licenseRule: LicenseRule;
    public lastExecuted?: string;
    public status?: number;

    /**
     * adds a group
     * @param { Group } group The Group to be created
     * @return { GroupEntitlement } this Group Entitlement
     */
    public withGroup(group: Group): this {
        this.group = group;
        return this;
    }

    /**
     * adds a license Rule
     * @param { LicenseRule } licenseRule The license Rule
     * @return { GroupEntitlement } this Group Entitlement
     */
    public withLicenseRule(licenseRule: LicenseRule): this {
        this.licenseRule = licenseRule;
        return this;
    }

    /**
     * adds a member to this Group Entitlement
     * @param { UserEntitlement } userEntitlement The User Entitlement that needs to be added to this Group Entitlement
     * @return { GroupEntitlement } this Group Entitlement
     */
    public async addMember(userEntitlement: UserEntitlement): Promise<this> {
        await AzureDevOpsConnection.replace(`_apis/GroupEntitlements/${this.id}/members/${userEntitlement.id}?api-version=6.0-preview.1`, undefined);
        Logger.log(`Added User '${userEntitlement.user.displayName}' to group '${this.group.displayName}`);
        return this;
    }

    /**
     * removes a member from this Group Entitlement
     * @param { UserEntitlement } userEntitlement The User Entitlement that needs to be removed from this Group Entitlement
     * @return { GroupEntitlement } this Group Entitlement
     */
    public async removeMember(userEntitlement: UserEntitlement): Promise<this> {
        await AzureDevOpsConnection.del(`_apis/GroupEntitlements/${this.id}/members/${userEntitlement.id}?api-version=6.0-preview.1`);
        Logger.log(`Removed User '${userEntitlement.user.displayName}' from group '${this.group.displayName}`);
        return this;
    }

    /**
     * Creates this Group Entitlement in Azure DevOps
     * @return { GroupEntitlement } this Group Entitlement
     */
    public async create(): Promise<this> {
        const apiResult = await AzureDevOpsConnection.create<GroupEntitlement>(`_apis/groupentitlements?api-version=6.1-preview.1`, this);
        this.id = apiResult.result.id;
        this.lastExecuted = apiResult.result.lastExecuted;
        this.licenseRule = apiResult.result.licenseRule;
        this.group = apiResult.result.group;
        this.status = apiResult.result.status;
        return this;
    }
}