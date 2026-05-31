export interface IUserRiskInfo {
    id?: number | string;
    displayName: string;
    email?: string;
    loginName?: string;
    isPossibleExternal: boolean;
    detectionReasons: string[];
    source?: string;
    groupName?: string;
    roleHint?: string;
}
