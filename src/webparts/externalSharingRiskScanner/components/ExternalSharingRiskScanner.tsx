import * as React from 'react';
import { SPFI } from '@pnp/sp';
import { Icon } from '@fluentui/react/lib/Icon';

import styles from './ExternalSharingRiskScanner.module.scss';
import { LoadingState } from './LoadingState';

import { IScanResult } from '../models/IScanResult';
import { RiskLabel } from '../models/IRiskScore';
import { IGroupSummary } from '../models/IGroupSummary';
import { IPermissionIndicator } from '../models/IPermissionIndicator';
import { IUserRiskInfo } from '../models/IUserRiskInfo';

import { SharePointSiteService } from '../services/SharePointSiteService';
import { SharePointUserService } from '../services/SharePointUserService';
import { SharePointGroupService } from '../services/SharePointGroupService';
import { PermissionService } from '../services/PermissionService';
import { RiskScoringService } from '../services/RiskScoringService';
import { LoggerService } from '../services/LoggerService';

// ═══════════════════════════════════
// Props
// ═══════════════════════════════════

export interface IExternalSharingRiskScannerProps {
    sp: SPFI;
    internalDomains: string;
    siteUrl: string;
}

// ═══════════════════════════════════
// State
// ═══════════════════════════════════

interface IComponentState {
    phase: 'idle' | 'scanning' | 'done' | 'error';
    scanResult: IScanResult | null;
    errorMessage: string;
}

// ═══════════════════════════════════
// Component
// ═══════════════════════════════════

export class ExternalSharingRiskScanner extends React.Component<
    IExternalSharingRiskScannerProps,
    IComponentState
> {
    constructor(props: IExternalSharingRiskScannerProps) {
        super(props);
        this.state = { phase: 'idle', scanResult: null, errorMessage: '' };
    }

    // ── Scanning Logic ──

    private async runScan(): Promise<void> {
        this.setState({ phase: 'scanning', errorMessage: '' });

        try {
            const { sp, internalDomains, siteUrl } = this.props;
            const domainList = internalDomains
                ? internalDomains.split(',').map((d) => d.trim().toLowerCase()).filter(Boolean)
                : [];

            const siteService = new SharePointSiteService(sp);
            const userService = new SharePointUserService(sp);
            const groupService = new SharePointGroupService(sp);
            const permissionService = new PermissionService(sp);
            const riskService = new RiskScoringService();

            LoggerService.info('Scanner', `Starting scan for ${siteUrl}`);

            const siteSummary = await siteService.getSiteSummary();
            const users = await userService.getSiteUsers(domainList);
            const groups = await groupService.getSiteGroups(domainList);
            const permissionIndicators = await permissionService.getPermissionIndicators();

            const partialResult = {
                siteSummary,
                users,
                groups,
                permissionIndicators,
                warnings: [] as string[],
                errors: [] as any[],
            };

            const riskScore = riskService.calculateRiskScore(partialResult);

            const scanResult: IScanResult = {
                ...partialResult,
                riskScore,
            };

            LoggerService.info('Scanner', `Scan complete. Risk: ${riskScore.label} (${riskScore.score})`);
            this.setState({ phase: 'done', scanResult });
        } catch (err: any) {
            LoggerService.error('Scanner', 'Scan failed', err);
            this.setState({ phase: 'error', errorMessage: err.message || 'Unknown error' });
        }
    }

    // ── Render ──

    public render(): React.ReactElement {
        const { phase, scanResult, errorMessage } = this.state;

        if (phase === 'scanning') {
            return <LoadingState />;
        }

        if (phase === 'error') {
            return (
                <div className={styles.dashboardRoot}>
                    <div className={styles.topBar}>
                        <h1 className={styles.topBarTitle}>External Sharing Risk Scanner</h1>
                        <button className={styles.scanButton} onClick={() => this.runScan()}>
                            <Icon iconName="Refresh" /> Retry
                        </button>
                    </div>
                    <div className={styles.warningBanner}>
                        <Icon iconName="ErrorBadge" className={styles.warningBannerIcon} />
                        <span>Scan failed: {errorMessage}</span>
                    </div>
                </div>
            );
        }

        if (phase === 'idle' || !scanResult) {
            return this.renderPreScan();
        }

        return this.renderDashboard(scanResult);
    }

    // ── Pre-scan state ──

    private renderPreScan(): React.ReactElement {
        return (
            <div className={styles.dashboardRoot}>
                <div className={styles.topBar}>
                    <h1 className={styles.topBarTitle}>External Sharing Risk Scanner</h1>
                    <button
                        className={styles.scanButton}
                        onClick={() => this.runScan()}
                    >
                        <Icon iconName="Shield" /> Run Scan
                    </button>
                </div>
                <div className={styles.preScanAlert}>
                    <Icon iconName="Info" className={styles.preScanIcon} />
                    <span>
                        Click <strong>Run Scan</strong> to analyze this site&#39;s external sharing configuration,
                        user permissions, and group memberships for potential risk signals.
                    </span>
                </div>
            </div>
        );
    }

    // ── Dashboard (post-scan) ──

    private renderDashboard(result: IScanResult): React.ReactElement {
        const { riskScore, users, groups, permissionIndicators, siteSummary } = result;
        const externalUsers = users.filter((u) => u.isPossibleExternal);
        const uniquePermCount = permissionIndicators.filter((p) => p.hasUniquePermissions).length;

        return (
            <div className={styles.dashboardRoot}>
                {/* Top bar with title + rescan */}
                <div className={styles.topBar}>
                    <h1 className={styles.topBarTitle}>External Sharing Risk Scanner</h1>
                    <button
                        className={styles.scanButton}
                        onClick={() => this.runScan()}
                    >
                        <Icon iconName="Refresh" /> Re-scan
                    </button>
                </div>

                {/* ── Metric Cards ── */}
                <div className={styles.metricGrid}>
                    {this.renderMetricCard(
                        'Shield',
                        'Risk Score',
                        `${riskScore.score}/100`,
                        riskScore.label,
                        this.getRiskStatusClass(riskScore.label)
                    )}
                    {this.renderMetricCard(
                        'People',
                        'External Users',
                        String(externalUsers.length),
                        externalUsers.length === 0 ? 'None detected' : `${externalUsers.length} found`,
                        externalUsers.length === 0 ? 'Success' : 'Warning'
                    )}
                    {this.renderMetricCard(
                        'Group',
                        'Groups Reviewed',
                        String(groups.length),
                        `${groups.filter((g) => g.possibleExternalUserCount && g.possibleExternalUserCount > 0).length} with external`,
                        groups.filter((g) => g.possibleExternalUserCount && g.possibleExternalUserCount > 0).length > 0 ? 'Warning' : 'Success'
                    )}
                    {this.renderMetricCard(
                        'Lock',
                        'Unique Permissions',
                        String(uniquePermCount),
                        uniquePermCount > 0 ? `${uniquePermCount} detected` : 'Inheriting normally',
                        uniquePermCount > 0 ? 'Warning' : 'Success'
                    )}
                    {this.renderMetricCard(
                        'DocumentSet',
                        'Libraries Reviewed',
                        String(permissionIndicators.filter((p) => p.scope === 'Library').length),
                        'Scanned',
                        'Info'
                    )}
                </div>

                {/* ── Last Scanned ── */}
                <div className={styles.lastScannedRow}>
                    <Icon iconName="Clock" className={styles.lastScannedIcon} />
                    <span>Last scanned: {this.formatDate(siteSummary.scannedAt)}</span>
                </div>

                {/* ── Warnings ── */}
                {result.warnings.length > 0 && (
                    <div className={styles.warningBanner}>
                        <Icon iconName="Warning" className={styles.warningBannerIcon} />
                        <span>{result.warnings.join(' | ')}</span>
                    </div>
                )}

                {/* ── Main Assessment Grid: Risk + Key Findings ── */}
                <div className={styles.mainAssessmentGrid}>
                    {this.renderRiskAssessmentCard(result)}
                    {this.renderKeyFindingsCard(result)}
                </div>

                {/* ── Secondary Grid: Groups + Permissions + External Indicators ── */}
                <div className={styles.secondaryGrid}>
                    {this.renderGroupsCard(groups)}
                    {this.renderPermissionsCard(permissionIndicators)}
                    {this.renderExternalIndicatorsCard(externalUsers)}
                </div>

                {/* ── Recommended Actions ── */}
                {riskScore.recommendedActions.length > 0 && this.renderActionsCard(riskScore.recommendedActions)}

                {/* ── Footer ── */}
                <div className={styles.dashboardFooter}>
                    External Sharing Risk Scanner v1.0 &mdash;{' '}
                    <a
                        href="https://github.com/BillySharePoint/spfx-external-sharing-risk-scanner"
                        target="_blank"
                        rel="noopener noreferrer"
                        className={styles.footerLink}
                    >
                        GitHub
                    </a>
                </div>
            </div>
        );
    }

    // ═══════════════════════════════════
    // Sub-renders
    // ═══════════════════════════════════

    private renderMetricCard(
        icon: string,
        label: string,
        value: string,
        statusText: string,
        statusType: string
    ): React.ReactElement {
        const statusClassMap: Record<string, string> = {
            Success: styles.metricStatusSuccess,
            Warning: styles.metricStatusWarning,
            Info: styles.metricStatusInfo,
            Neutral: styles.metricStatusNeutral,
        };
        return (
            <div className={styles.metricCard}>
                <div className={styles.metricIconWrap}>
                    <Icon iconName={icon} />
                </div>
                <div className={styles.metricBody}>
                    <p className={styles.metricLabel}>{label}</p>
                    <p className={styles.metricValue}>{value}</p>
                    <p className={`${styles.metricStatus} ${statusClassMap[statusType] || styles.metricStatusNeutral}`}>
                        {statusText}
                    </p>
                </div>
            </div>
        );
    }

    // ── Risk Assessment Card ──

    private renderRiskAssessmentCard(result: IScanResult): React.ReactElement {
        const { riskScore } = result;
        const ringColor = this.getRiskColor(riskScore.label);

        return (
            <div className={`${styles.card}`}>
                <div className={styles.cardHeader}>
                    <Icon iconName="Shield" className={styles.cardHeaderIcon} />
                    <h2 className={styles.cardHeaderTitle}>Risk Assessment</h2>
                </div>
                <div className={styles.cardBody}>
                    <div className={styles.riskAssessmentBody}>
                        {/* Gauge panel */}
                        <div className={styles.gaugePanel}>
                            <div
                                className={styles.riskRing}
                                style={{
                                    background: `conic-gradient(${ringColor} ${riskScore.score * 3.6}deg, #e5e7eb ${riskScore.score * 3.6}deg)`,
                                }}
                            >
                                <div className={styles.riskRingInner} />
                                <div className={styles.riskRingContent}>
                                    <span className={styles.riskNumber} style={{ color: ringColor }}>
                                        {riskScore.score}
                                    </span>
                                    <span className={styles.riskSubLabel}>/ 100</span>
                                </div>
                            </div>
                            <span className={styles.riskLevelText} style={{ color: ringColor }}>
                                {riskScore.label} Risk
                            </span>
                        </div>

                        {/* Summary panel */}
                        <div className={styles.riskSummaryPanel}>
                            <h3>Assessment Summary</h3>
                            <p>{this.getRiskSummaryText(riskScore.label, riskScore.score)}</p>

                            {riskScore.reasons.length > 0 && (
                                <>
                                    <h4>Contributing Factors</h4>
                                    <ul>
                                        {riskScore.reasons.map((r, idx) => (
                                            <li key={idx}>{r}</li>
                                        ))}
                                    </ul>
                                </>
                            )}
                        </div>
                    </div>
                </div>
            </div>
        );
    }

    // ── Key Findings Card ──

    private renderKeyFindingsCard(result: IScanResult): React.ReactElement {
        const { users, groups, permissionIndicators } = result;
        const externalUsers = users.filter((u) => u.isPossibleExternal);
        const uniquePerms = permissionIndicators.filter((p) => p.hasUniquePermissions);
        const groupsWithExt = groups.filter(
            (g) => g.possibleExternalUserCount && g.possibleExternalUserCount > 0
        );

        const findings: { title: string; desc: string; type: string }[] = [];

        if (externalUsers.length === 0) {
            findings.push({
                title: 'No External Users Detected',
                desc: 'All users appear to be internal domain members.',
                type: 'success',
            });
        } else {
            findings.push({
                title: `${externalUsers.length} External User(s) Found`,
                desc: `Possible external users: ${externalUsers.slice(0, 3).map((u) => u.displayName).join(', ')}${externalUsers.length > 3 ? '...' : ''}`,
                type: 'warning',
            });
        }

        if (groupsWithExt.length > 0) {
            findings.push({
                title: `${groupsWithExt.length} Group(s) Contain External Users`,
                desc: groupsWithExt.map((g) => g.name).join(', '),
                type: 'warning',
            });
        } else {
            findings.push({
                title: 'Groups Are Internal Only',
                desc: 'No external users found in any SharePoint group.',
                type: 'success',
            });
        }

        if (uniquePerms.length > 0) {
            findings.push({
                title: `${uniquePerms.length} Unique Permission(s)`,
                desc: 'Some resources have broken inheritance.',
                type: 'info',
            });
        } else {
            findings.push({
                title: 'Consistent Permissions',
                desc: 'All items inherit permissions from the site.',
                type: 'success',
            });
        }

        if (result.errors.length > 0) {
            findings.push({
                title: `${result.errors.length} Scan Error(s)`,
                desc: result.errors[0].message,
                type: 'danger',
            });
        }

        return (
            <div className={`${styles.card}`}>
                <div className={styles.cardHeader}>
                    <Icon iconName="Lightbulb" className={styles.cardHeaderIcon} />
                    <h2 className={styles.cardHeaderTitle}>Key Findings</h2>
                </div>
                <div className={styles.cardBody}>
                    <div className={styles.findingsList}>
                        {findings.map((f, idx) => this.renderFinding(f, idx))}
                    </div>
                </div>
            </div>
        );
    }

    private renderFinding(
        finding: { title: string; desc: string; type: string },
        idx: number
    ): React.ReactElement {
        const typeMap: Record<string, { row: string; icon: string; iconSymbol: string }> = {
            success: { row: styles.findingSuccess, icon: styles.findingIconSuccess, iconSymbol: '✓' },
            warning: { row: styles.findingWarning, icon: styles.findingIconWarning, iconSymbol: '!' },
            info: { row: styles.findingInfo, icon: styles.findingIconInfo, iconSymbol: 'i' },
            danger: { row: styles.findingDanger, icon: styles.findingIconDanger, iconSymbol: '✕' },
        };
        const t = typeMap[finding.type] || typeMap.info;
        return (
            <div className={`${styles.findingItem} ${t.row}`} key={idx}>
                <div className={`${styles.findingIcon} ${t.icon}`}>{t.iconSymbol}</div>
                <div className={styles.findingBody}>
                    <p className={styles.findingTitle}>{finding.title}</p>
                    <p className={styles.findingDescription}>{finding.desc}</p>
                </div>
            </div>
        );
    }

    // ── Groups Card ──

    private renderGroupsCard(groups: IGroupSummary[]): React.ReactElement {
        const groupsWithExternal = groups.filter(
            (g) => g.possibleExternalUserCount && g.possibleExternalUserCount > 0
        );
        return (
            <div className={`${styles.card}`}>
                <div className={styles.cardHeader}>
                    <Icon iconName="Group" className={styles.cardHeaderIcon} />
                    <h2 className={styles.cardHeaderTitle}>SharePoint Groups</h2>
                </div>
                <div className={styles.cardBody}>
                    <div className={styles.cardStatsRow}>
                        <div>
                            <p className={styles.statValue}>{groups.length}</p>
                            <p className={styles.statLabel}>Total Groups</p>
                        </div>
                        <div>
                            <p className={`${styles.statValue} ${groupsWithExternal.length > 0 ? styles.warningText : styles.successText}`}>
                                {groupsWithExternal.length}
                            </p>
                            <p className={styles.statLabel}>With External</p>
                        </div>
                    </div>

                    {groupsWithExternal.length > 0 && (
                        <div className={styles.inlineWarningBanner}>
                            <Icon iconName="Warning" />
                            <span>{groupsWithExternal.length} group(s) contain possible external users</span>
                        </div>
                    )}

                    <div className={styles.tableWrap}>
                        <table className={styles.dataTable}>
                            <thead>
                                <tr>
                                    <th>Group Name</th>
                                    <th>Members</th>
                                    <th>External</th>
                                </tr>
                            </thead>
                            <tbody>
                                {groups.slice(0, 8).map((g, idx) => (
                                    <tr key={idx}>
                                        <td>{g.name}</td>
                                        <td>{g.userCount ?? '—'}</td>
                                        <td className={
                                            (g.possibleExternalUserCount ?? 0) > 0
                                                ? styles.statusReview
                                                : styles.statusOk
                                        }>
                                            {g.possibleExternalUserCount ?? 0}
                                        </td>
                                    </tr>
                                ))}
                            </tbody>
                        </table>
                    </div>
                </div>
            </div>
        );
    }

    // ── Permissions Card ──

    private renderPermissionsCard(permissions: IPermissionIndicator[]): React.ReactElement {
        const uniqueCount = permissions.filter((p) => p.hasUniquePermissions).length;
        return (
            <div className={`${styles.card}`}>
                <div className={styles.cardHeader}>
                    <Icon iconName="Lock" className={styles.cardHeaderIcon} />
                    <h2 className={styles.cardHeaderTitle}>Permission Inheritance</h2>
                </div>
                <div className={styles.cardBody}>
                    <div className={styles.cardStatsRow}>
                        <div>
                            <p className={styles.statValue}>{permissions.length}</p>
                            <p className={styles.statLabel}>Items Reviewed</p>
                        </div>
                        <div>
                            <p className={`${styles.statValue} ${uniqueCount > 0 ? styles.warningText : styles.successText}`}>
                                {uniqueCount}
                            </p>
                            <p className={styles.statLabel}>Unique Permissions</p>
                        </div>
                    </div>

                    <div className={styles.tableWrap}>
                        <table className={styles.dataTable}>
                            <thead>
                                <tr>
                                    <th>Resource</th>
                                    <th>Scope</th>
                                    <th>Status</th>
                                </tr>
                            </thead>
                            <tbody>
                                {permissions.slice(0, 8).map((p, idx) => (
                                    <tr key={idx}>
                                        <td>{p.title}</td>
                                        <td>{p.scope}</td>
                                        <td className={
                                            p.status === 'Ok'
                                                ? styles.statusOk
                                                : p.status === 'Review'
                                                    ? styles.statusReview
                                                    : styles.statusEmpty
                                        }>
                                            {p.status}
                                        </td>
                                    </tr>
                                ))}
                            </tbody>
                        </table>
                    </div>
                </div>
            </div>
        );
    }

    // ── External User Indicators Card ──

    private renderExternalIndicatorsCard(externalUsers: IUserRiskInfo[]): React.ReactElement {
        return (
            <div className={`${styles.card}`}>
                <div className={styles.cardHeader}>
                    <Icon iconName="ContactInfo" className={styles.cardHeaderIcon} />
                    <h2 className={styles.cardHeaderTitle}>External User Indicators</h2>
                </div>
                <div className={styles.cardBody}>
                    {externalUsers.length === 0 ? (
                        <div className={styles.successPanelLarge}>
                            <div className={styles.successCircleIcon}>✓</div>
                            <p className={styles.successPanelTitle}>No External Users Detected</p>
                            <p className={styles.successPanelText}>
                                All site users appear to belong to internal domains.
                                No external sharing indicators were found.
                            </p>
                        </div>
                    ) : (
                        <div className={styles.dangerPanelLarge}>
                            <div className={styles.dangerCircleIcon}>!</div>
                            <p className={styles.dangerPanelTitle}>
                                {externalUsers.length} External User(s) Detected
                            </p>
                            <p className={styles.dangerPanelText}>
                                {externalUsers.slice(0, 4).map((u) => u.displayName).join(', ')}
                                {externalUsers.length > 4 ? ` and ${externalUsers.length - 4} more` : ''}
                            </p>
                        </div>
                    )}
                </div>
            </div>
        );
    }

    // ── Recommended Actions ──

    private renderActionsCard(actions: string[]): React.ReactElement {
        const actionIcons = [
            'Shield', 'People', 'Lock', 'Settings', 'SkypeCheck', 'DocumentSearch',
        ];
        return (
            <div className={`${styles.card} ${styles.recommendedActionsCard}`}>
                <div className={styles.cardHeader}>
                    <Icon iconName="Rocket" className={styles.cardHeaderIcon} />
                    <h2 className={styles.cardHeaderTitle}>Recommended Actions</h2>
                </div>
                <div className={styles.cardBody}>
                    <div className={styles.actionGrid}>
                        {actions.map((action, idx) => (
                            <div className={styles.actionTile} key={idx}>
                                <div className={styles.actionIconWrap}>
                                    <Icon iconName={actionIcons[idx % actionIcons.length]} />
                                </div>
                                <span className={styles.actionText}>{action}</span>
                                <span className={styles.actionArrow}>›</span>
                            </div>
                        ))}
                    </div>
                </div>
            </div>
        );
    }

    // ═══════════════════════════════════
    // Helpers
    // ═══════════════════════════════════

    private getRiskColor(label: RiskLabel): string {
        switch (label) {
            case 'Low': return '#107c10';
            case 'Medium': return '#c2410c';
            case 'High': return '#d83b01';
            case 'Review Recommended': return '#2563eb';
            default: return '#6b7280';
        }
    }

    private getRiskStatusClass(label: RiskLabel): string {
        switch (label) {
            case 'Low': return 'Success';
            case 'Medium': return 'Warning';
            case 'High': return 'Warning';
            case 'Review Recommended': return 'Info';
            default: return 'Neutral';
        }
    }

    private getRiskSummaryText(label: RiskLabel, score: number): string {
        if (label === 'Low') {
            return `This site has a low risk score of ${score}/100. No significant external sharing concerns were identified. The site permissions appear well-configured and all users belong to expected internal domains.`;
        }
        if (label === 'Medium') {
            return `This site has a moderate risk score of ${score}/100. Some external sharing signals were detected that warrant review. Consider checking the contributing factors below and following the recommended actions.`;
        }
        if (label === 'High') {
            return `This site has an elevated risk score of ${score}/100. Multiple external sharing concerns were detected. Immediate review of external users and permission configurations is recommended.`;
        }
        if (label === 'Review Recommended') {
            return `This site scored ${score}/100 and requires further review. Some aspects of the permissions or user configuration could not be fully verified.`;
        }
        return `Risk assessment could not be fully completed. Score: ${score}/100.`;
    }

    private formatDate(dateStr: string): string {
        try {
            const d = new Date(dateStr);
            return d.toLocaleString(undefined, {
                year: 'numeric',
                month: 'short',
                day: 'numeric',
                hour: '2-digit',
                minute: '2-digit',
            });
        } catch {
            return dateStr;
        }
    }
}
