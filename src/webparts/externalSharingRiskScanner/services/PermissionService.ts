import { SPFI } from '@pnp/sp';
import '@pnp/sp/webs';
import '@pnp/sp/lists';
import '@pnp/sp/security/list';

import { IPermissionIndicator } from '../models/IPermissionIndicator';
import { MAX_LIBRARIES_TO_SCAN } from '../constants/RiskConstants';
import { LoggerService } from './LoggerService';

const SOURCE = 'PermissionService';

export class PermissionService {
    private _sp: SPFI;

    constructor(sp: SPFI) {
        this._sp = sp;
    }

    /**
     * Gets permission indicators for the current site and its document libraries.
     * Combines site-level and library-level checks.
     */
    public async getPermissionIndicators(): Promise<IPermissionIndicator[]> {
        LoggerService.info(SOURCE, 'Gathering permission indicators...');

        const indicators: IPermissionIndicator[] = [];

        // Get site-level permission indicator
        try {
            const siteIndicator = await this.getSitePermissionIndicator();
            indicators.push(siteIndicator);
        } catch (error) {
            LoggerService.error(SOURCE, 'Failed to check site permissions', error);
            indicators.push({
                scope: 'Site',
                title: 'Current Site',
                status: 'Error',
                notes: ['Unable to check site permission inheritance.'],
            });
        }

        // Get library-level permission indicators
        try {
            const libraryIndicators = await this.getLibraryPermissionIndicators();
            indicators.push(...libraryIndicators);
        } catch (error) {
            LoggerService.error(SOURCE, 'Failed to check library permissions', error);
            indicators.push({
                scope: 'Library',
                title: 'Document Libraries',
                status: 'Error',
                notes: ['Unable to check library permission inheritance.'],
            });
        }

        return indicators;
    }

    /**
     * Checks the site/web level for unique role assignments.
     */
    private async getSitePermissionIndicator(): Promise<IPermissionIndicator> {
        // eslint-disable-next-line @typescript-eslint/no-explicit-any
        const webInfo: any = await this._sp.web.select(
            'Title',
            'Url',
            'HasUniqueRoleAssignments'
        )();

        return {
            scope: 'Site',
            title: webInfo.Title || 'Current Site',
            url: webInfo.Url,
            hasUniquePermissions: webInfo.HasUniqueRoleAssignments,
            status: webInfo.HasUniqueRoleAssignments ? 'Review' : 'Ok',
            notes: webInfo.HasUniqueRoleAssignments
                ? ['Site has unique permissions — not inheriting from parent site.']
                : ['Site inherits permissions from its parent.'],
        };
    }

    /**
     * Checks document libraries (BaseTemplate 101, non-hidden) for unique permissions.
     * Limited to MAX_LIBRARIES_TO_SCAN to prevent performance issues.
     */
    public async getLibraryPermissionIndicators(): Promise<IPermissionIndicator[]> {
        LoggerService.info(SOURCE, 'Checking document library permissions...');

        try {
            // Get non-hidden document libraries
            // eslint-disable-next-line @typescript-eslint/no-explicit-any
            const lists: any[] = await this._sp.web.lists
                .filter('BaseTemplate eq 101 and Hidden eq false')
                .select('Id', 'Title', 'RootFolder/ServerRelativeUrl', 'HasUniqueRoleAssignments')
                .expand('RootFolder')
                .top(MAX_LIBRARIES_TO_SCAN)();

            LoggerService.info(SOURCE, `Found ${lists.length} document libraries to check`);

            const indicators: IPermissionIndicator[] = [];

            for (const list of lists) {
                indicators.push({
                    scope: 'Library',
                    title: list.Title,
                    url: list.RootFolder?.ServerRelativeUrl,
                    hasUniquePermissions: list.HasUniqueRoleAssignments,
                    status: list.HasUniqueRoleAssignments ? 'Review' : 'Ok',
                    notes: list.HasUniqueRoleAssignments
                        ? ['This library has unique permissions — access may differ from site defaults.']
                        : undefined,
                });
            }

            return indicators;
        } catch (error) {
            LoggerService.error(SOURCE, 'Failed to enumerate libraries', error);
            throw error;
        }
    }
}
