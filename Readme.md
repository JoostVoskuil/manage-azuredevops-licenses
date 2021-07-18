# Manage Azure DevOps User Licenses

This little utility manages the Azure DevOps User Licenses. Most organizations use Azure Active Directory (AAD) groups to manage the permissions of users within projects.

When a new AAD user enters the Azure DevOps organization for the first time, the user is created within the Azure DevOps organization and the user gets a license assigned. This is the so called 'User Entitlement'.

However, there are two problems with the link between Azure DevOps and Azure Active Directory:

- When the user is deleted (or made inactive) in the AAD, the user is still active in Azure DevOps. The user cannot login anymore through the web interface but the personal access tokens are still active. 
- When an user is not using Azure DevOps, the user might occupy a paid license;

To save money and to lower security risks, this little utility is created. It creates four so called 'Group Entitlements' and make users member of those Group Entitlements.

If you have created your own Group Entitlements for setting access to Projects, don't worry. The utility is compatible with that. Group rule types are ranked in this order: Subscriber > Basic + Test Plans > Basic > Stakeholder, so the highest rank wins.

## How to use

- Run 'npm run build'. You can find the application in the 'dist' directory.
- You can run the application in two ways:
  1. Run as a console application: 'node ConsoleApp.js --help'
  2. Run as an Azure Functions (included azure devops pipeline)

I recommend to run this utility first in test mode with the disableAPIOperations is set to true. This disables the Create, Update, Delete operations.

## Configuration

### Setup Azure DevOps

The default license is set in the 'Billing' screen in the 'Organizational Settings'. I recommend that the value of 'Default access level for new users' is set to 'Basic'.

### Setup Azure Active Directory

You need to register an App into your Azure Active Directory configuration

1. Login to Azure and goto 'Azure Active Directory';
2. Goto 'App Registrations';
3. Create a new Registrations
   1. Choose your single or Multi tenant
   2. Leave redirect url empty
4. Goto your App registration 'Overview'
   1. Store the 'Application (client) ID'. This is your 'AADGraphApplicationId'
   2. Store the 'Directory (tenant) ID'. This is your 'AADGraphDirectoryId'
5. Goto your App registration 'API Permissions'
   1. Choose 'Add permission' and then 'Microsoft Graph'
   2. Choose 'Application permissions;
      1. Select 'User.Read.All and then 'Add Permission'
      2. If you want to delete users from AAD after they did not sign in for a certain period, also include AuditLogs.Read.All and Organization.Read.All
   3. Click on 'Grant admin consent..'
6. Goto your App registration 'Certificates & Secrets'
   1. Create a new 'Client Secret'
   2. Store the Value. This is your 'AADGraphApplicationToken'

### Configuration of this utility

The default settings are managed in the settings.json file. These are the 'Application settings'

| Parameter | Explaination | Default value |
| ------------- |:-------------:| -----:|
| numberOfDaysNotLoggedInForDeletionOfUser | if the user did not login in x days, the user is deleted in Azure DevOps | 186 |
| numberOfDaysNotLoggedInToBecomeStakeholder | if the user did not login in x days, the user is made stakeholder      |   $93 |
| numberOfDaysToWaitAfterUserIsCreated | Grace period for accounts that are created before action | 31 |
| groupEntitlementPrefix | Prefix for the Group Entitlements to be created by this utilty | All users -  |
| groupEntitlementPostfix | Postfix for the Group Entitlements to be created by this utilty | License |
| azureDevOpsVSAexBaseUrl | The Azure DevOps VSAEX api url | <https://vsaex.dev.azure.com> |
| disableAPIOperations | When set to true, there are no API change operations (Create, Update, Delete) | true |
| AADGraphOAuthEndPoint | The Azure AAD Url | <https://login.microsoftonline.com> |
| excludedWordsInUPNs | Do not process UPN's with the following terms (comma seperated). Is used for service accounts | svc,service|
| excludedUPNs | Do not process the following UPN's (comma seperated). | empty |

The following settings are the 'User settings':

| Parameter | Explaination |
| ------------- |-------------|
| azureDevOpsOrganizationName | The organization name |
| personalAccessToken | The PAT that is used for Azure DevOps. Requires Project Collection Administrator permissions |
| AADGraphDirectoryId | The AAD Directory ID (Tenant) |
| AADGraphApplicationId | The registered AAD Application ID for this utility |
| AADGraphApplicationToken | The Token for the AAD Application for this utility |

#### Console Application

The user settings needed to be specified in the console application. Run 'Node ConsoleApp.js --help' to see the options.

#### Azure Function Configuration

Runs on a Node.Js Function App:

1. Create a Function App in Azure with settings:
   - Publish = Code
   - Runtime Stack = Node.js
   - Runtime Version = 12 LTS
   - OS = Linux
   - Plantype = Consumption (Serverless)