import { SPFI } from '@pnp/sp';
import '@pnp/sp/site-users/web';

import { IUserRiskInfo } from '../models/IUserRiskInfo';
import { detectExternalUser } from '../utils/externalUserUtils';
import { isSystemAccount } from '../constants/ExternalUserPatterns';
import { LoggerService } from './LoggerService';

const SOURCE = 'SharePointUserService';

export class SharePointUserService {
    private _sp: SPFI;

    constructor(sp: SPFI) {
        this._sp = sp;
    }

    /**
     * Retrieves all visible site users and evaluates each for external user indicators.
     *
     * @param internalDomains - Array of lowercase internal domain strings
     * @returns Array of user risk info objects
     */
    public async getSiteUsers(internalDomains: string[]): Promise<IUserRiskInfo[]> {
        LoggerService.info(SOURCE, 'Fetching site users...');

        try {
            const users = await this._sp.web.siteUsers
                .select(
                    'Id',
                    'Title',
                    'Email',
                    'LoginName',
                    'PrincipalType',
                    'IsSiteAdmin'
                )();

            LoggerService.info(SOURCE, `Retrieved ${users.length} site users`);

            const results: IUserRiskInfo[] = [];

            for (const user of users) {
                // Skip non-user principal types (e.g., SharePoint groups, security groups)
                // PrincipalType 1 = User, 4 = Security Group, 8 = SharePoint Group
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

                const roleHint = user.IsSiteAdmin ? 'Site Collection Administrator' : undefined;

                results.push({
                    id: user.Id,
                    displayName: user.Title || '(No display name)',
                    email: user.Email || undefined,
                    loginName: user.LoginName || undefined,
                    isPossibleExternal: detection.isPossibleExternal,
                    detectionReasons: detection.reasons,
                    source: 'Site Users',
                    roleHint,
                });
            }

            const externalCount = results.filter((u) => u.isPossibleExternal).length;
            LoggerService.info(
                SOURCE,
                `Processed ${results.length} users, ${externalCount} flagged as possible external`
            );

            return results;
        } catch (error) {
            LoggerService.error(SOURCE, 'Failed to fetch site users', error);
            throw error;
        }
    }

    /**
     * Utility method to check if a single user appears to be external.
     */
    public isPossibleExternalUser(
        loginName: string,
        email: string | undefined,
        displayName: string,
        internalDomains: string[]
    ): boolean {
        const result = detectExternalUser(loginName, email, displayName, internalDomains);
        return result.isPossibleExternal;
    }
}
