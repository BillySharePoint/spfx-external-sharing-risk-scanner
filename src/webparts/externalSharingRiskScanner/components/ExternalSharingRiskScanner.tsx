import * as React from 'react';
import { SPFI } from '@pnp/sp';

import styles from './ExternalSharingRiskScanner.module.scss';

import { IScanResult } from '../models/IScanResult';
import { IScannerError } from '../models/IScannerError';
import { RiskLabel } from '../models/IRiskScore';
import { IGroupSummary } from '../models/IGroupSummary';
import { IPermissionIndicator } from '../models/IPermissionIndicator';

import { SharePointSiteService } from '../services/SharePointSiteService';
import { SharePointUserService } from '../services/SharePointUserService';
import { SharePointGroupService } from '../services/SharePointGroupService';
import { PermissionService } from '../services/PermissionService';
import { RiskScoringService } from '../services/RiskScoringService';
import { LoggerService } from '../services/LoggerService';

import { LoadingState } from './LoadingState';

/* ────────────────────────────── Props ────────────────────────────── */

export interface IExternalSharingRiskScannerProps {
    sp: SPFI;
    internalDomains: string;
    siteUrl: string;
}

/* ────────────────────────────── Helpers ───────────────────────────── */

const riskColorMap: Record<string, string> = {
    Low: '#107c10',
    Medium: '#ffb900',
    High: '#d13438',
    'Review Recommended': '#d13438',
    Unknown: '#8a8886',
};

const riskBadgeClass = (label: RiskLabel): string => {
    switch (label) {
        case 'Low': return styles.badgeSuccess;
        case 'Medium': return styles.badgeWarning;
        case 'High': return styles.badgeDanger;
        case 'Review Recommended': return styles.badgeDanger;
        default: return styles.badgeNeutral;
    }
};

const parseInternalDomains = (raw: string): string[] =>
    raw
        .split(/[,;\n]+/)
        .map((d) => d.trim().toLowerCase())
        .filter((d) => d.length > 0);

/* ────────────────── Recommended Action Model ─────────────────────── */

interface IRecommendedAction {
    priority: 'High' | 'Medium' | 'Low';
    title: string;
    reason: string;
    suggestedOwner: string;
}

const deriveActions = (
    externalUserCount: number,
    emptyGroupCount: number,
    uniquePermCount: number,
    errorCount: number
): IRecommendedAction[] => {
    const actions: IRecommendedAction[] = [];

    if (externalUserCount > 0) {
        actions.push({
            priority: 'High',
            title: 'Review external users in SharePoint groups',
            reason: `${externalUserCount} possible external user(s) detected — confirm they still need access.`,
            suggestedOwner: 'Site owner or M365 admin',
        });
    }

    if (uniquePermCount > 0) {
        actions.push({
            priority: 'Medium',
            title: 'Review libraries with unique permissions',
            reason: 'Unique permissions can increase governance complexity and access drift.',
            suggestedOwner: 'Site owner',
        });
    }

    if (emptyGroupCount > 0) {
        actions.push({
            priority: 'Medium',
            title: 'Review empty SharePoint groups',
            reason: `${emptyGroupCount} empty group(s) may indicate stale or incomplete permission setup.`,
            suggestedOwner: 'Site owner or M365 admin',
        });
    }

    if (errorCount > 0) {
        actions.push({
            priority: 'Medium',
            title: 'Resolve scan errors',
            reason: `${errorCount} error(s) occurred — some data may be incomplete. Ensure sufficient permissions.`,
            suggestedOwner: 'Site Collection Administrator',
        });
    }

    actions.push({
        priority: 'Low',
        title: 'Validate external sharing settings with your M365 admin',
        reason: 'This web part only shows selected indicators from the current site context.',
        suggestedOwner: 'M365 admin',
    });

    actions.push({
        priority: 'Low',
        title: 'Confirm whether external access is still needed',
        reason: 'Review before migration, Copilot rollout, or tenant-wide governance audits.',
        suggestedOwner: 'Site owner',
    });

    return actions;
};

/* ────────────────────────────── Component ─────────────────────────── */

const ExternalSharingRiskScanner: React.FC<IExternalSharingRiskScannerProps> = (props) => {
    const { sp, internalDomains, siteUrl } = props;

    const [scanResult, setScanResult] = React.useState<IScanResult | null>(null);
    const [isScanning, setIsScanning] = React.useState<boolean>(false);
    const [hasScanned, setHasScanned] = React.useState<boolean>(false);
    const [showErrors, setShowErrors] = React.useState<boolean>(false);

    /* ── services ── */
    const siteService = React.useMemo(() => new SharePointSiteService(sp), [sp]);
    const userService = React.useMemo(() => new SharePointUserService(sp), [sp]);
    const groupService = React.useMemo(() => new SharePointGroupService(sp), [sp]);
    const permissionService = React.useMemo(() => new PermissionService(sp), [sp]);
    const riskScoringService = React.useMemo(() => new RiskScoringService(), []);

    const domains = React.useMemo(() => parseInternalDomains(internalDomains), [internalDomains]);
    const hasInternalDomains = domains.length > 0;

    /* ── scan ── */
    const runScan = React.useCallback(async (): Promise<void> => {
        setIsScanning(true);
        const errors: IScannerError[] = [];

        try {
            LoggerService.info('Scanner', `Starting scan for ${siteUrl}`);

            const [siteSummary, users, groups, permissionIndicators] = await Promise.all([
                siteService.getSiteSummary().catch((e) => {
                    errors.push({ source: 'Site Summary', message: String(e), technicalDetails: String(e) });
                    return { title: 'Unknown', url: siteUrl, scannedAt: new Date().toLocaleString() };
                }),
                userService.getSiteUsers(domains).catch((e) => {
                    errors.push({ source: 'User Scan', message: String(e), technicalDetails: String(e) });
                    return [];
                }),
                groupService.getSiteGroups(domains).catch((e) => {
                    errors.push({ source: 'Group Scan', message: String(e), technicalDetails: String(e) });
                    return [];
                }),
                permissionService.getPermissionIndicators().catch((e) => {
                    errors.push({ source: 'Permissions', message: String(e), technicalDetails: String(e) });
                    return [];
                }),
            ]);

            const warnings: string[] = [];
            if (!hasInternalDomains) {
                warnings.push('No internal domains configured. External user detection is limited to login name patterns only.');
            }

            const partial = { siteSummary, users, groups, permissionIndicators, errors, warnings };
            const riskScore = riskScoringService.calculateRiskScore(partial);

            const result: IScanResult = {
                siteSummary,
                users,
                groups,
                permissionIndicators,
                riskScore,
                warnings,
                errors,
            };

            setScanResult(result);
            setHasScanned(true);
            LoggerService.info('Scanner', `Scan complete — score: ${riskScore.score}`);
        } catch (error) {
            LoggerService.error('Scanner', 'Scan failed', error);
            errors.push({ source: 'Scanner', message: 'Unexpected error during scan.', technicalDetails: String(error) });
        } finally {
            setIsScanning(false);
        }
    }, [sp, siteUrl, domains, hasInternalDomains, siteService, userService, groupService, permissionService, riskScoringService]);

    /* ── derived data ── */
    const r = scanResult;
    const externalUserCount = r?.users?.filter((u) => u.isPossibleExternal).length ?? 0;
    const groupCount = r?.groups?.length ?? 0;
    const emptyGroupCount = r?.groups?.filter((g) => g.isEmpty).length ?? 0;
    const libraryIndicators = r?.permissionIndicators?.filter((p) => p.scope === 'Library') ?? [];
    const libraryCount = libraryIndicators.length;
    const uniquePermCount = libraryIndicators.filter((p) => p.hasUniquePermissions).length;

    const riskScore = r?.riskScore;
    const riskColor = riskScore ? (riskColorMap[riskScore.label] || '#8a8886') : '#8a8886';
    const riskAngle = riskScore ? riskScore.score * 3.6 : 0;

    /* ── findings ── */
    type Severity = 'success' | 'warning' | 'info' | 'danger';
    interface IFinding { severity: Severity; icon: string; title: string; description: string }

    const findings = React.useMemo((): IFinding[] => {
        if (!r) return [];
        const items: IFinding[] = [];

        if (externalUserCount === 0) {
            items.push({ severity: 'success', icon: '✓', title: 'No possible external users detected', description: 'No external users were found in the SharePoint groups scanned. Tenant-wide sharing settings and sharing links should still be reviewed separately.' });
        } else {
            items.push({ severity: 'danger', icon: '!', title: `${externalUserCount} possible external user(s) detected`, description: 'External users detected in SharePoint groups — review access promptly.' });
        }

        if (uniquePermCount > 0) {
            items.push({ severity: 'warning', icon: '⚠', title: `${uniquePermCount} librar${uniquePermCount === 1 ? 'y has' : 'ies have'} unique permissions`, description: 'Unique permissions are not automatically bad, but they should be reviewed because they can increase governance complexity.' });
        }

        if (emptyGroupCount > 0) {
            items.push({ severity: 'warning', icon: '⚠', title: `${emptyGroupCount} empty SharePoint group(s) found`, description: `${emptyGroupCount} SharePoint group(s) have no members. This may be expected in a test site, but in production it can indicate stale or incomplete permission setup.` });
        }

        if (groupCount > 0) {
            items.push({ severity: 'info', icon: 'i', title: `${groupCount} SharePoint groups scanned`, description: `${groupCount} SharePoint groups were scanned for possible external users.` });
        }

        if (r.errors.length > 0) {
            items.push({ severity: 'warning', icon: '⚠', title: 'Scan completed with errors', description: `${r.errors.length} error(s) occurred — some data may be incomplete.` });
        }

        return items;
    }, [r, externalUserCount, uniquePermCount, emptyGroupCount, groupCount]);

    /* ── recommended actions ── */
    const recommendedActions = React.useMemo((): IRecommendedAction[] => {
        if (!r) return [];
        return deriveActions(externalUserCount, emptyGroupCount, uniquePermCount, r.errors.length);
    }, [r, externalUserCount, emptyGroupCount, uniquePermCount]);

    /* ── severity CSS maps ── */
    const severityFindingCls: Record<Severity, string> = {
        success: styles.findingSuccess,
        warning: styles.findingWarning,
        info: styles.findingInfo,
        danger: styles.findingDanger,
    };
    const severityIconCls: Record<Severity, string> = {
        success: styles.findingIconSuccess,
        warning: styles.findingIconWarning,
        info: styles.findingIconInfo,
        danger: styles.findingIconDanger,
    };

    const priorityCls: Record<string, string> = {
        High: styles.priorityHigh,
        Medium: styles.priorityMedium,
        Low: styles.priorityLow,
    };

    /* ── metric card helper ── */
    const getMetricStatus = (label: string, count: number): string => {
        if (label === 'Possible external users') return count === 0 ? 'No external users detected' : 'Review recommended';
        if (label === 'Groups reviewed') return `${count} SharePoint group(s) scanned`;
        if (label === 'Libraries reviewed') return `${count} document librar${count === 1 ? 'y' : 'ies'} checked`;
        if (label === 'Unique permissions') return count === 0 ? 'All libraries inherit permissions' : `${count} librar${count === 1 ? 'y' : 'ies'} may need review`;
        return '';
    };

    /* ────────── JSX ────────── */

    return (
        <div className={styles.externalSharingRiskScanner}>
            <div className={styles.shell}>
                {/* ── Header / Command Bar ── */}
                <section className={styles.header}>
                    <div className={styles.headerLeft}>
                        <h1 className={styles.title}>External Sharing Risk Scanner</h1>
                        <p className={styles.description}>
                            Review external sharing and permission risk signals for the current SharePoint site.
                        </p>
                        {r && (
                            <p className={styles.timestamp}>Last scanned: {r.siteSummary.scannedAt}</p>
                        )}
                    </div>
                    <div className={styles.headerActions}>
                        <button
                            className={`${styles.button} ${styles.buttonPrimary}`}
                            type="button"
                            onClick={runScan}
                            disabled={isScanning}
                            aria-label={isScanning ? 'Scanning in progress' : 'Scan site for external sharing risks'}
                        >
                            {isScanning ? 'Scanning…' : 'Scan site'}
                        </button>
                    </div>
                </section>

                {/* ── Warning: no internal domains ── */}
                {!hasInternalDomains && (
                    <div className={`${styles.alert} ${styles.alertWarning}`} role="status">
                        <span className={styles.alertIcon}>⚠️</span>
                        <div>
                            <strong>No internal domains configured.</strong>{' '}
                            External user detection is limited to login name patterns only. Configure internal
                            domains in the web part properties for better detection.
                        </div>
                    </div>
                )}

                {/* ── Loading ── */}
                {isScanning && <LoadingState />}

                {/* ── Scan errors toggle ── */}
                {r && r.errors.length > 0 && !isScanning && (
                    <div className={`${styles.alert} ${styles.alertWarning}`} role="alert">
                        <span className={styles.alertIcon}>⚠️</span>
                        <div>
                            <strong>{r.errors.length} error(s) during scan.</strong>{' '}
                            <button
                                type="button"
                                className={styles.collapseToggle}
                                onClick={() => setShowErrors(!showErrors)}
                            >
                                <span className={`${styles.collapseArrow} ${showErrors ? styles.collapseArrowOpen : ''}`}>▶</span>
                                {showErrors ? 'Hide details' : 'Show details'}
                            </button>
                            {showErrors && (
                                <ul style={{ margin: '8px 0 0 16px', padding: 0 }}>
                                    {r.errors.map((e, i) => (
                                        <li key={i} style={{ marginBottom: 4 }}>{e.source}: {e.message}</li>
                                    ))}
                                </ul>
                            )}
                        </div>
                    </div>
                )}

                {/* ── Results ── */}
                {r && !isScanning && (
                    <>
                        {/* ── Dashboard Metric Cards ── */}
                        <section className={styles.metricsRow} aria-label="Summary metrics">
                            <MetricCard icon="👥" title="Possible external users" value={externalUserCount} status={getMetricStatus('Possible external users', externalUserCount)} statusType={externalUserCount === 0 ? 'success' : 'danger'} />
                            <MetricCard icon="🧩" title="Groups reviewed" value={groupCount} status={getMetricStatus('Groups reviewed', groupCount)} statusType="neutral" />
                            <MetricCard icon="📚" title="Libraries reviewed" value={libraryCount} status={getMetricStatus('Libraries reviewed', libraryCount)} statusType="neutral" />
                            <MetricCard icon="🔓" title="Unique permissions" value={uniquePermCount} status={getMetricStatus('Unique permissions', uniquePermCount)} statusType={uniquePermCount === 0 ? 'success' : 'warning'} />
                        </section>

                        {/* ── Risk Score + Key Findings ── */}
                        <section className={styles.riskScoreSection}>
                            {/* Risk Assessment Card */}
                            <article className={styles.card}>
                                <div className={styles.cardHeader}>
                                    <div>
                                        <h2 className={styles.cardTitle}>Risk Assessment</h2>
                                        <p className={styles.cardSubtitle}>{riskScore!.score} / 100 — {riskScore!.label}</p>
                                    </div>
                                    <span className={`${styles.badge} ${riskBadgeClass(riskScore!.label)}`}>
                                        {riskScore!.label}
                                    </span>
                                </div>
                                <div className={styles.cardBody}>
                                    <div className={styles.riskScoreVisual}>
                                        <div
                                            className={styles.riskRing}
                                            aria-label={`Risk score ${riskScore!.score} out of 100, ${riskScore!.label}`}
                                            style={{ background: `conic-gradient(${riskColor} 0deg ${riskAngle}deg, #f3f2f1 ${riskAngle}deg 360deg)` }}
                                        >
                                            <div className={styles.riskRingInner} />
                                            <div className={styles.riskRingContent}>
                                                <div className={styles.riskNumber} style={{ color: riskColor }}>{riskScore!.score}</div>
                                                <div className={styles.riskSubLabel}>/100</div>
                                            </div>
                                        </div>
                                        <div className={styles.riskDetails}>
                                            <p className={styles.riskSummary}>
                                                {riskScore!.reasons[0] || 'No significant risk indicators detected.'}
                                            </p>
                                            {riskScore!.reasons.length > 1 && (
                                                <ul className={styles.riskFactors}>
                                                    {riskScore!.reasons.slice(1).map((reason, i) => (
                                                        <li key={i}>{reason}</li>
                                                    ))}
                                                </ul>
                                            )}
                                        </div>
                                    </div>
                                    <div className={styles.riskBar}>
                                        <div className={styles.riskBarFill} style={{ width: `${riskScore!.score}%`, background: riskColor }} />
                                    </div>
                                </div>
                            </article>

                            {/* Key Findings Card */}
                            {findings.length > 0 && (
                                <article className={styles.card}>
                                    <div className={styles.cardHeader}>
                                        <div>
                                            <h2 className={styles.cardTitle}>Key Findings</h2>
                                            <p className={styles.cardSubtitle}>Prioritized signals from the latest scan</p>
                                        </div>
                                    </div>
                                    <div className={styles.cardBody}>
                                        <div className={styles.findingsList}>
                                            {findings.map((f, idx) => (
                                                <div key={idx} className={`${styles.finding} ${severityFindingCls[f.severity]}`}>
                                                    <div className={`${styles.findingIcon} ${severityIconCls[f.severity]}`}>{f.icon}</div>
                                                    <div className={styles.findingBody}>
                                                        <p className={styles.findingTitle}>{f.title}</p>
                                                        <p className={styles.findingDescription}>{f.description}</p>
                                                    </div>
                                                </div>
                                            ))}
                                        </div>
                                    </div>
                                </article>
                            )}
                        </section>

                        {/* ── Site Summary + Scope/Limitations ── */}
                        <section className={styles.twoColumnGrid}>
                            <SiteSummaryCard siteSummary={r.siteSummary} />
                            <ScopeLimitationsCard />
                        </section>

                        {/* ── Groups Table + Permission Inheritance ── */}
                        <section className={styles.twoColumnGrid}>
                            <GroupsTable groups={r.groups} />
                            <PermissionsCard indicators={libraryIndicators} />
                        </section>

                        {/* ── Recommended Actions ── */}
                        {recommendedActions.length > 0 && (
                            <section className={styles.sectionSpacing}>
                                <article className={styles.card}>
                                    <div className={styles.cardHeader}>
                                        <div>
                                            <h2 className={styles.cardTitle}>Recommended Actions</h2>
                                            <p className={styles.cardSubtitle}>Suggested governance follow-up steps</p>
                                        </div>
                                    </div>
                                    <div className={styles.cardBody}>
                                        <ol className={styles.actionsList}>
                                            {recommendedActions.map((action, idx) => (
                                                <li key={idx} className={styles.actionItem}>
                                                    <span className={`${styles.actionPriority} ${priorityCls[action.priority] || ''}`}>
                                                        {action.priority}
                                                    </span>
                                                    <div className={styles.actionBody}>
                                                        <p className={styles.actionTitle}>{action.title}</p>
                                                        <p className={styles.actionReason}>{action.reason}</p>
                                                        <p className={styles.actionOwner}>Suggested owner: {action.suggestedOwner}</p>
                                                    </div>
                                                </li>
                                            ))}
                                        </ol>
                                    </div>
                                </article>
                            </section>
                        )}

                        {/* ── Footer ── */}
                        <div className={styles.footer}>
                            <p style={{ margin: '0 0 6px 0' }}>
                                This public MVP provides site-level indicators only. For a full governance review,
                                validate results with the SharePoint admin center, Microsoft Purview, audit logs, and
                                tenant-level sharing settings.
                            </p>
                            <p style={{ margin: 0 }}>
                                Need deeper reporting, CSV export, or tenant-specific governance checks?{' '}
                                <a
                                    className={styles.footerLink}
                                    href="https://www.billyperalta.com/contact"
                                    target="_blank"
                                    rel="noopener noreferrer"
                                >
                                    Contact Billy Peralta
                                </a>{' '}
                                for Pro/private access or customization.
                            </p>
                        </div>
                    </>
                )}

                {/* ── Pre-scan ── */}
                {!hasScanned && !isScanning && (
                    <div className={`${styles.alert} ${styles.alertInfo}`} role="status">
                        <span className={styles.alertIcon}>ℹ️</span>
                        <div>
                            Click <strong>Scan site</strong> to begin analyzing external sharing risks for this
                            SharePoint site. The scan reviews SharePoint groups, document library permission
                            inheritance, and possible external user indicators.
                        </div>
                    </div>
                )}
            </div>
        </div>
    );
};

/* ─────────── Sub-components ─────────── */

/* ── Metric Card ── */
interface IMetricCardProps {
    icon: string;
    title: string;
    value: number;
    status: string;
    statusType: 'success' | 'warning' | 'danger' | 'neutral';
}

const statusColorMap: Record<string, string> = {
    success: '#107c10',
    warning: '#835b00',
    danger: '#d13438',
    neutral: '#605e5c',
};

const MetricCard: React.FC<IMetricCardProps> = ({ icon, title, value, status, statusType }) => (
    <div className={styles.metricCard}>
        <div className={styles.metricCardHeader}>
            <p className={styles.metricTitle}>{title}</p>
            <span className={styles.metricIcon}>{icon}</span>
        </div>
        <p className={styles.metricValue}>{value}</p>
        <p className={styles.metricStatus} style={{ color: statusColorMap[statusType] }}>{status}</p>
    </div>
);

/* ── Site Summary Card ── */
interface ISiteSummaryCardProps {
    siteSummary: { title: string; url: string; webId?: string; currentUserDisplayName?: string; currentUserPermissionLevel?: string; scannedAt: string };
}

const SiteSummaryCard: React.FC<ISiteSummaryCardProps> = ({ siteSummary }) => {
    const rows: Array<{ label: string; value: string }> = [
        { label: 'Site title', value: siteSummary.title },
        { label: 'Site URL', value: siteSummary.url },
        { label: 'Current user', value: siteSummary.currentUserDisplayName || '—' },
        { label: 'Permission level', value: siteSummary.currentUserPermissionLevel || '—' },
        { label: 'Scan time', value: siteSummary.scannedAt },
    ];
    if (siteSummary.webId) {
        rows.push({ label: 'Web ID', value: siteSummary.webId });
    }

    return (
        <article className={styles.card}>
            <div className={styles.cardHeader}>
                <div>
                    <h2 className={styles.cardTitle}>Site Summary</h2>
                    <p className={styles.cardSubtitle}>Current site context and identity</p>
                </div>
            </div>
            <div className={styles.cardBody}>
                <div className={styles.summaryGrid}>
                    {rows.map((row, i) => (
                        <div key={i} className={styles.summaryRow}>
                            <span className={styles.summaryLabel}>{row.label}</span>
                            <span className={styles.summaryValue}>{row.value}</span>
                        </div>
                    ))}
                </div>
            </div>
        </article>
    );
};

/* ── Scope & Limitations Card ── */
const ScopeLimitationsCard: React.FC = () => (
    <article className={styles.card}>
        <div className={styles.cardHeader}>
            <div>
                <h2 className={styles.cardTitle}>Scan Scope &amp; Limitations</h2>
                <p className={styles.cardSubtitle}>What this scan does and does not check</p>
            </div>
        </div>
        <div className={styles.cardBody}>
            <div className={styles.scopeSection}>
                <h3 className={styles.scopeTitle}>What this scan checks</h3>
                <ul className={styles.scopeList}>
                    <li>SharePoint groups for the current site</li>
                    <li>Basic external user patterns in group membership</li>
                    <li>Document library permission inheritance</li>
                    <li>Current site context and user permissions</li>
                </ul>
            </div>
            <div className={styles.scopeSection}>
                <h3 className={styles.scopeTitle}>What this scan does not fully check</h3>
                <ul className={styles.scopeList}>
                    <li>Tenant-wide sharing policies</li>
                    <li>Microsoft Purview alerts or DLP policies</li>
                    <li>All item-level permissions</li>
                    <li>Anonymous sharing links</li>
                    <li>Historical sharing events or audit logs</li>
                    <li>Sensitivity labels</li>
                </ul>
            </div>
            <p className={styles.scopeNote}>
                This public MVP provides site-level indicators only. For a full governance review,
                validate results with the SharePoint admin center, Microsoft Purview, audit logs,
                and tenant-level sharing settings.
            </p>
        </div>
    </article>
);

/* ── Groups Table ── */
interface IGroupsTableProps {
    groups: IGroupSummary[];
}

const GroupsTable: React.FC<IGroupsTableProps> = ({ groups }) => {
    const getStatus = (g: IGroupSummary): { text: string; cls: string; note: string } => {
        if (g.possibleExternalUserCount && g.possibleExternalUserCount > 0) {
            return { text: 'External detected', cls: styles.statusReview, note: 'Review recommended' };
        }
        if (g.isEmpty) {
            return { text: 'Empty', cls: styles.statusEmpty, note: 'No members found' };
        }
        return { text: 'Internal only', cls: styles.statusOk, note: '' };
    };

    const hasEmptyGroups = groups.filter((g) => g.isEmpty).length > 0;

    return (
        <article className={styles.card}>
            <div className={styles.cardHeader}>
                <div>
                    <h2 className={styles.cardTitle}>SharePoint Groups</h2>
                    <p className={styles.cardSubtitle}>Group membership review summary</p>
                </div>
            </div>
            <div className={styles.cardBody}>
                {hasEmptyGroups && (
                    <div className={styles.permissionSummary}>
                        Several SharePoint groups are empty. This may be expected in a new or test site,
                        but in production it can indicate stale or incomplete permission configuration.
                    </div>
                )}
                <div className={styles.tableWrap}>
                    <table className={styles.dataTable}>
                        <thead>
                            <tr>
                                <th>Group name</th>
                                <th>Members</th>
                                <th>Possible external</th>
                                <th>Status</th>
                                <th>Notes</th>
                            </tr>
                        </thead>
                        <tbody>
                            {groups.map((g, i) => {
                                const st = getStatus(g);
                                return (
                                    <tr key={i}>
                                        <td>{g.name}</td>
                                        <td>{g.userCount ?? '–'}</td>
                                        <td>{g.possibleExternalUserCount ?? '–'}</td>
                                        <td><span className={st.cls}>{st.text}</span></td>
                                        <td style={{ color: '#605e5c', fontSize: 12 }}>{st.note || (g.notes ? g.notes.join('; ') : '—')}</td>
                                    </tr>
                                );
                            })}
                            {groups.length === 0 && (
                                <tr><td colSpan={5} style={{ textAlign: 'center', color: '#a19f9d', padding: 24 }}>No groups found</td></tr>
                            )}
                        </tbody>
                    </table>
                </div>
            </div>
        </article>
    );
};

/* ── Permissions Card ── */
interface IPermissionsCardProps {
    indicators: IPermissionIndicator[];
}

const PermissionsCard: React.FC<IPermissionsCardProps> = ({ indicators }) => {
    const uniqueCount = indicators.filter((p) => p.hasUniquePermissions).length;

    return (
        <article className={styles.card}>
            <div className={styles.cardHeader}>
                <div>
                    <h2 className={styles.cardTitle}>Permission Inheritance</h2>
                    <p className={styles.cardSubtitle}>Document library inheritance status</p>
                </div>
            </div>
            <div className={styles.cardBody}>
                <div className={styles.permissionSummary}>
                    <span><span className={styles.permStatNumber}>{indicators.length}</span> libraries reviewed</span>
                    <span><span className={styles.permStatNumber}>{uniqueCount}</span> with unique permissions</span>
                </div>
                <div className={styles.tableWrap}>
                    <table className={styles.dataTable}>
                        <thead>
                            <tr>
                                <th>Library</th>
                                <th>Inheritance</th>
                                <th>Status</th>
                            </tr>
                        </thead>
                        <tbody>
                            {indicators.map((ind, i) => (
                                <tr key={i}>
                                    <td>{ind.title}</td>
                                    <td>{ind.hasUniquePermissions ? 'Unique' : 'Inherited'}</td>
                                    <td>
                                        <span className={ind.hasUniquePermissions ? styles.statusReview : styles.statusOk}>
                                            {ind.hasUniquePermissions ? '⚠ Review needed' : '✓ OK'}
                                        </span>
                                    </td>
                                </tr>
                            ))}
                            {indicators.length === 0 && (
                                <tr><td colSpan={3} style={{ textAlign: 'center', color: '#a19f9d', padding: 24 }}>No libraries found</td></tr>
                            )}
                        </tbody>
                    </table>
                </div>
                {uniqueCount > 0 && (
                    <p className={styles.permissionNote}>
                        Unique permissions are not automatically bad, but they should be reviewed because they can
                        increase governance complexity.
                    </p>
                )}
            </div>
        </article>
    );
};

export { ExternalSharingRiskScanner };
export default ExternalSharingRiskScanner;
