import { IScanResult } from '../models/IScanResult';
import { IRiskScore, RiskLabel } from '../models/IRiskScore';
import {
    RISK_WEIGHTS,
    RISK_THRESHOLDS,
    MAX_RISK_SCORE,
    MIN_RISK_SCORE,
    EXTERNAL_USER_THRESHOLDS,
} from '../constants/RiskConstants';
import { LoggerService } from './LoggerService';

const SOURCE = 'RiskScoringService';

export class RiskScoringService {
    /**
     * Calculates a deterministic risk score based on scan results.
     * The scoring logic is transparent and easy to explain.
     *
     * @param scanPartial - Partial scan result data to score
     * @returns Calculated risk score with reasons and recommended actions
     */
    public calculateRiskScore(scanPartial: Partial<IScanResult>): IRiskScore {
        LoggerService.info(SOURCE, 'Calculating risk score...');

        let score = MIN_RISK_SCORE;
        const reasons: string[] = [];
        const recommendedActions: string[] = [];

        // Check for data retrieval errors
        if (scanPartial.errors && scanPartial.errors.length > 0) {
            score += RISK_WEIGHTS.DATA_RETRIEVAL_FAILURE;
            reasons.push(
                `${scanPartial.errors.length} error(s) occurred during scan — some data may be incomplete`
            );
            recommendedActions.push(
                'Review scan errors and ensure you have sufficient permissions to access site data'
            );
        }

        // Evaluate external users
        if (scanPartial.users) {
            const externalUsers = scanPartial.users.filter((u) => u.isPossibleExternal);

            if (externalUsers.length > 0) {
                score += RISK_WEIGHTS.EXTERNAL_USERS_FOUND;
                reasons.push(
                    `${externalUsers.length} possible external user(s) detected in the site`
                );
                recommendedActions.push(
                    'Review external users in SharePoint site permissions to confirm they still need access'
                );

                // Check for many external users
                if (externalUsers.length > EXTERNAL_USER_THRESHOLDS.MANY) {
                    score += RISK_WEIGHTS.MANY_EXTERNAL_USERS;
                    reasons.push(
                        `More than ${EXTERNAL_USER_THRESHOLDS.MANY} external users found — elevated risk`
                    );
                }

                // Check for external users with high privilege
                const externalAdmins = externalUsers.filter(
                    (u) => u.roleHint && u.roleHint.toLowerCase().includes('admin')
                );
                if (externalAdmins.length > 0) {
                    score += RISK_WEIGHTS.EXTERNAL_USERS_IN_OWNERS;
                    reasons.push(
                        `${externalAdmins.length} external user(s) have administrative roles`
                    );
                    recommendedActions.push(
                        'Urgently review external users with admin/owner privileges'
                    );
                }
            }
        }

        // Evaluate groups
        if (scanPartial.groups) {
            const groupsWithExternal = scanPartial.groups.filter(
                (g) => g.possibleExternalUserCount && g.possibleExternalUserCount > 0
            );

            if (groupsWithExternal.length > 0) {
                // Check if any Owner-type groups have external users
                const ownerGroups = groupsWithExternal.filter(
                    (g) => g.name.toLowerCase().includes('owner')
                );
                if (ownerGroups.length > 0) {
                    score += RISK_WEIGHTS.EXTERNAL_USERS_IN_OWNERS;
                    reasons.push('External users found in Owner group(s)');
                    recommendedActions.push(
                        'Review Owner group membership — external users in Owner groups have elevated permissions'
                    );
                }

                const memberGroups = groupsWithExternal.filter(
                    (g) => g.name.toLowerCase().includes('member')
                );
                if (memberGroups.length > 0) {
                    score += RISK_WEIGHTS.EXTERNAL_USERS_IN_MEMBERS;
                    reasons.push('External users found in Member group(s)');
                    recommendedActions.push(
                        'Confirm external users in Member groups should have edit-level access'
                    );
                }
            }

            // Check for empty groups
            const emptyGroups = scanPartial.groups.filter((g) => g.isEmpty === true);
            if (emptyGroups.length > 0) {
                score += RISK_WEIGHTS.EMPTY_GROUPS_WITH_PERMISSIONS;
                reasons.push(`${emptyGroups.length} empty group(s) found — may indicate stale permissions`);
                recommendedActions.push('Review empty SharePoint groups and remove if no longer needed');
            }
        }

        // Evaluate permission indicators
        if (scanPartial.permissionIndicators) {
            const siteIndicators = scanPartial.permissionIndicators.filter(
                (p) => p.scope === 'Site'
            );
            const libraryIndicators = scanPartial.permissionIndicators.filter(
                (p) => p.scope === 'Library'
            );

            // Site-level unique permissions
            const siteHasUnique = siteIndicators.some((p) => p.hasUniquePermissions === true);
            if (siteHasUnique) {
                score += RISK_WEIGHTS.SITE_UNIQUE_PERMISSIONS;
                reasons.push('Site has unique (broken) permission inheritance');
                recommendedActions.push(
                    'Review site permission inheritance — unique permissions can lead to access drift'
                );
            }

            // Library-level unique permissions
            const librariesWithUnique = libraryIndicators.filter(
                (p) => p.hasUniquePermissions === true
            );
            if (librariesWithUnique.length > 1) {
                score += RISK_WEIGHTS.MULTIPLE_LIBRARIES_UNIQUE_PERMISSIONS;
                reasons.push(
                    `${librariesWithUnique.length} document libraries have unique permissions`
                );
                recommendedActions.push(
                    'Review library permissions — multiple libraries with unique permissions increase governance complexity'
                );
            } else if (librariesWithUnique.length === 1) {
                score += RISK_WEIGHTS.SINGLE_LIBRARY_UNIQUE_PERMISSIONS;
                reasons.push('1 document library has unique permissions');
                recommendedActions.push(
                    'Review the library with unique permissions to confirm access is appropriate'
                );
            }
        }

        // Cap the score
        score = Math.min(score, MAX_RISK_SCORE);
        score = Math.max(score, MIN_RISK_SCORE);

        // Determine label
        const label = this.getLabel(score, scanPartial);

        // Add general recommended actions
        if (recommendedActions.length === 0) {
            recommendedActions.push(
                'No significant risk indicators found — continue routine governance reviews'
            );
        }

        // Always add these general recommendations
        recommendedActions.push(
            'Validate sharing settings with your Microsoft 365 admin'
        );
        recommendedActions.push(
            'Confirm whether external access is still needed before migration or Copilot rollout'
        );

        const result: IRiskScore = {
            score,
            label,
            reasons: reasons.length > 0 ? reasons : ['No significant risk indicators detected'],
            recommendedActions,
        };

        LoggerService.info(SOURCE, `Risk score calculated: ${score} (${label})`);
        return result;
    }

    /**
     * Maps a numeric score to a risk label.
     */
    private getLabel(score: number, scanPartial: Partial<IScanResult>): RiskLabel {
        // If there were significant errors, show "Unknown"
        if (
            scanPartial.errors &&
            scanPartial.errors.length > 2 &&
            (!scanPartial.users || scanPartial.users.length === 0)
        ) {
            return 'Unknown';
        }

        if (score <= RISK_THRESHOLDS.LOW_MAX) return 'Low';
        if (score <= RISK_THRESHOLDS.MEDIUM_MAX) return 'Medium';
        if (score <= RISK_THRESHOLDS.HIGH_MAX) return 'High';
        return 'Review Recommended';
    }
}
