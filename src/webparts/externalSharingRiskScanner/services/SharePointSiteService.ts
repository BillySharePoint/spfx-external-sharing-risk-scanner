import { SPFI } from '@pnp/sp';
import '@pnp/sp/webs';
import '@pnp/sp/site-users/web';

import { ISiteSummary } from '../models/ISiteSummary';
import { IPermissionIndicator } from '../models/IPermissionIndicator';
import { getFormattedScanTime } from '../utils/dateUtils';
import { LoggerService } from './LoggerService';

const SOURCE = 'SharePointSiteService';

export class SharePointSiteService {
    private _sp: SPFI;

    constructor(sp: SPFI) {
        this._sp = sp;
    }

    /**
     * Retrieves basic site summary information including current user details.
     */
    public async getSiteSummary(): Promise<ISiteSummary> {
        LoggerService.info(SOURCE, 'Fetching site summary...');

        try {
            const [webInfo, currentUser] = await Promise.all([
                this._sp.web.select(
                    'Title',
                    'Url',
                    'Id'
                )(),
                this._sp.web.currentUser.select(
                    'Title',
                    'Email',
                    'LoginName',
                    'IsSiteAdmin'
                )(),
            ]);

            const permissionLevel = currentUser.IsSiteAdmin
                ? 'Site Collection Administrator'
                : 'Site Member (permissions vary)';

            const summary: ISiteSummary = {
                title: webInfo.Title,
                url: webInfo.Url,
                webId: webInfo.Id,
                currentUserDisplayName: currentUser.Title,
                currentUserEmail: currentUser.Email,
                currentUserPermissionLevel: permissionLevel,
                scannedAt: getFormattedScanTime(),
            };

            LoggerService.info(SOURCE, `Site summary retrieved for: ${summary.title}`);
            return summary;
        } catch (error) {
            LoggerService.error(SOURCE, 'Failed to fetch site summary', error);
            throw error;
        }
    }

    /**
     * Checks whether the current web has unique (broken) role assignments.
     */
    public async getCurrentWebPermissionIndicator(): Promise<IPermissionIndicator> {
        LoggerService.info(SOURCE, 'Checking web permission inheritance...');

        try {
            // eslint-disable-next-line @typescript-eslint/no-explicit-any
            const webInfo: any = await this._sp.web.select(
                'Title',
                'Url',
                'HasUniqueRoleAssignments'
            )();

            const hasUnique = webInfo.HasUniqueRoleAssignments;

            return {
                scope: 'Site',
                title: webInfo.Title || 'Current Site',
                url: webInfo.Url,
                hasUniquePermissions: hasUnique,
                status: hasUnique ? 'Review' : 'Ok',
                notes: hasUnique
                    ? ['This site has unique permissions (not inheriting from parent).']
                    : ['This site inherits permissions from its parent.'],
            };
        } catch (error) {
            LoggerService.error(SOURCE, 'Failed to check web permissions', error);
            return {
                scope: 'Site',
                title: 'Current Site',
                hasUniquePermissions: undefined,
                status: 'Error',
                notes: ['Unable to determine permission inheritance status.'],
            };
        }
    }
}
