import { SPFI } from '@pnp/sp';
import '@pnp/sp/site-groups/web';
import '@pnp/sp/site-users/web';

import { IGroupSummary } from '../models/IGroupSummary';
import { IUserRiskInfo } from '../models/IUserRiskInfo';
import { detectExternalUser } from '../utils/externalUserUtils';
import { isSystemAccount } from '../constants/ExternalUserPatterns';
import { LoggerService } from './LoggerService';

const SOURCE = 'SharePointGroupService';

export class SharePointGroupService {
    private _sp: SPFI;

    constructor(sp: SPFI) {
        this._sp = sp;
    }

    /**
     * Retrieves SharePoint groups for the current site with user counts and external user indicators.
     *
     * @param internalDomains - Array of lowercase internal domain strings
     * @returns Array of group summaries
     */
    public async getSiteGroups(internalDomains: string[]): Promise<IGroupSummary[]> {
        LoggerService.info(SOURCE, 'Fetching site groups...');

        try {
            const groups = await this._sp.web.siteGroups
                .select('Id', 'Title', 'OwnerTitle', 'Description')();

            LoggerService.info(SOURCE, `Retrieved ${groups.length} site groups`);

            const results: IGroupSummary[] = [];

            for (const group of groups) {
                try {
                    const users = await this.getUsersForGroupById(group.Id, internalDomains);
                    const externalCount = users.filter((u) => u.isPossibleExternal).length;

                    results.push({
                        id: group.Id,
                        name: group.Title,
                        userCount: users.length,
                        possibleExternalUserCount: externalCount,
                        isEmpty: users.length === 0,
                        notes: externalCount > 0
                            ? [`${externalCount} possible external user(s) detected`]
                            : undefined,
                    });
                } catch (groupError) {
                    // Permission denied for this group's users - still include the group
                    LoggerService.warn(
                        SOURCE,
                        `Unable to read members of group: ${group.Title}`,
                        groupError
                    );
                    results.push({
                        id: group.Id,
                        name: group.Title,
                        userCount: undefined,
                        possibleExternalUserCount: undefined,
                        isEmpty: undefined,
                        notes: ['Unable to read group members (insufficient permissions)'],
                    });
                }
            }

            return results;
        } catch (error) {
            LoggerService.error(SOURCE, 'Failed to fetch site groups', error);
            throw error;
        }
    }

    /**
     * Retrieves users for a specific group by ID and evaluates external indicators.
     *
     * @param groupId - SharePoint group ID
     * @param internalDomains - Array of lowercase internal domain strings
     * @returns Array of user risk info for the group
     */
    public async getUsersForGroup(
        groupId: number,
        internalDomains: string[]
    ): Promise<IUserRiskInfo[]> {
        return this.getUsersForGroupById(groupId, internalDomains);
    }

    private async getUsersForGroupById(
        groupId: number,
        internalDomains: string[]
    ): Promise<IUserRiskInfo[]> {
        const users = await this._sp.web.siteGroups
            .getById(groupId)
            .users.select('Id', 'Title', 'Email', 'LoginName', 'PrincipalType')();

        const results: IUserRiskInfo[] = [];

        for (const user of users) {
            // Skip non-user principals
            if (user.PrincipalType !== 1) {
                continue;
            }

            // Skip system accounts
            if (isSystemAccount(user.LoginName || '')) {
                continue;
            }

            const detection = detectExternalUser(
                user.LoginName || '',
                user.Email || undefined,
                user.Title || '',
                internalDomains
            );

            results.push({
                id: user.Id,
                displayName: user.Title || '(No display name)',
                email: user.Email || undefined,
                loginName: user.LoginName || undefined,
                isPossibleExternal: detection.isPossibleExternal,
                detectionReasons: detection.reasons,
                source: 'SharePoint Group',
                groupName: undefined, // Will be set by caller if needed
            });
        }

        return results;
    }
}
