import { ISiteSummary } from './ISiteSummary';
import { IUserRiskInfo } from './IUserRiskInfo';
import { IGroupSummary } from './IGroupSummary';
import { IPermissionIndicator } from './IPermissionIndicator';
import { IRiskScore } from './IRiskScore';
import { IScannerError } from './IScannerError';

export interface IScanResult {
    siteSummary: ISiteSummary;
    users: IUserRiskInfo[];
    groups: IGroupSummary[];
    permissionIndicators: IPermissionIndicator[];
    riskScore: IRiskScore;
    warnings: string[];
    errors: IScannerError[];
}
