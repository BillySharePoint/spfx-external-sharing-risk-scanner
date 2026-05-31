export interface IPermissionIndicator {
    scope: 'Site' | 'Library' | 'Folder' | 'Unknown';
    title: string;
    url?: string;
    hasUniquePermissions?: boolean;
    status: 'Ok' | 'Review' | 'Unknown' | 'Error';
    notes?: string[];
}
