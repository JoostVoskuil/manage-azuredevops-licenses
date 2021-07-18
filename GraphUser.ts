import { AzureConnection } from "./AzureConnection";
import { Settings } from "./Settings";
import { Logger } from "./Logger";

export class GraphUser {
    public id?: string;
    public deletedDateTime?: string;
    public accountEnabled?: boolean;
    public userPrincipalName?: string;
    public createdDateTime?: Date;
    public SignInActivity?: {
        LastSignInDateTime: Date;
    }
    error?: {
        code?: string;
        message?: string;
    }

    public async getAADGraphUser(principalName: string): Promise<GraphUser> {
        // use beta to query signinactivity
        const url = `https://graph.microsoft.com/beta/users/${principalName}?$select=id,accountEnabled,deletedDateTime,userPrincipalName,signInActivity,createdDateTime`;
        const response = await AzureConnection.HttpGet(url);
        const graphUser = JSON.parse(await (response.readBody()));
        this.id = graphUser.id;
        this.deletedDateTime = graphUser.deletedDateTime;
        this.accountEnabled = graphUser.accountEnabled;
        this.userPrincipalName = graphUser.userPrincipalName;
        this.createdDateTime = graphUser.createdDateTime;
        this.SignInActivity = graphUser.SignInActivity;
        this.error = graphUser.error;
        return this;
    }
    /**
     * Checks if user is active in the AAD
     * @return { boolean } true if this user is active
     */
    public async isActiveInAAD(): Promise<boolean> {
        if (this.error && this.error.code === 'Request_ResourceNotFound') {
            Logger.log(`User '${this.userPrincipalName}' cannot be found in AAD.`);
            return false;
        }
        else if (this.accountEnabled === false) {
            Logger.log(`User '${this.userPrincipalName}' account is inactive.`);
            return false;
        }
        else if (this.deletedDateTime !== null) {
            Logger.log(`User '${this.userPrincipalName}' is deleted on ${this.deletedDateTime}.`);
            return false;
        }
        return true;
    }

    public async shouldBeDeletedFromAAD(): Promise<boolean> {
        if (this.SignInActivity.LastSignInDateTime < Settings.getConfiguration().dateDeleteBeforeLastAccessDate &&
            this.createdDateTime < Settings.getConfiguration().dateProcessAfterCreatedOnDate) {
            return true;
        }
        return false;
    }


    /**
     * Deletes an user in the AAD
     * @return { boolean } true if this user is active
     */
    public async deleteFromAAD(): Promise<void> {
        const url = `https://graph.microsoft.com/v1.0/users/${this.userPrincipalName}`;
        await AzureConnection.HttpDelete(url);

        Logger.log(`Deleted User '${this.userPrincipalName}' from AAD because LastSignInDateTime is ${this.SignInActivity.LastSignInDateTime}.`);
    }
}
