export interface IGroupSummary {
    id?: number | string;
    name: string;
    userCount?: number;
    possibleExternalUserCount?: number;
    isEmpty?: boolean;
    notes?: string[];
}
