// The License Rule Class
export class LicenseRule {
    public accountLicenseType: string;
    public assignmentSource: string;
    public licenseDisplayName: string;
    public licensingSource: string;
    public msdnLicenseType: string;
    public status: string;
    public statusMessage: string;
}

// Enum for all the different licenses
export enum License {
    Basic = 'Basic',
    Basic_TestPlans = 'Basic + Test Plans',
    Stakeholder = 'Stakeholder',
    VisualStudioSubscriber = 'Visual Studio Subscriber'
}