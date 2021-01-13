import { GroupEntitlement } from "./GroupEntitlement";
import { UserEntitlement } from "./UserEntitlement";

// API Response
export interface IGroupEntitlementResponse {
    count: number;
    value: GroupEntitlement[]
}

// API Response
export interface IUserEntitlementResponse {
    count: number;
    value: UserEntitlement[]
}
