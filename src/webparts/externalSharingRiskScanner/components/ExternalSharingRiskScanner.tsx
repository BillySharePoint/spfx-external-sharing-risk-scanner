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
    Medium: '#ca5010',
    High: '#a4262c',
    'Review Recommended': '#a4262c',
    Unknown: '#6b7280',
};

const riskBadgeClass = (label: RiskLabel): string => {
    switch (label) {
        case 'Low': return styles.badgeSuccess;
        case 'Medium': return styles.badgeWarning;
        case 'High': return styles.badgeDanger;
        case 'Review Recommended': return styles.badgeDanger;
        default: return styles.badgeInfo;
    }
};

const parseInternalDomains = (raw: string): string[] =>
    raw
        .split(/[,;\n]+/)
        .map((d) => d.trim().toLowerCase())
        .filter((d) => d.length > 0);

/* ────────────────────────────── Component ─────────────────────────── */

const ExternalSharingRiskScanner: React.FC<IExternalSharingRiskScannerProps> = (props) => {
    const { sp, internalDomains, siteUrl } = props;

    const [scanResult, setScanResult] = React.useState<IScanResult | null>(null);
    const [isScanning, setIsScanning] = React.useState<boolean>(false);
    const [hasScanned, setHasScanned] = React.useState<boolean>(false);
    const [showDetails, setShowDetails] = React.useState<boolean>(false);

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
    const externalUserCount = scanResult?.users?.filter((u) => u.isPossibleExternal).length ?? 0;
    const groupCount = scanResult?.groups?.length ?? 0;
    const libraryIndicators = scanResult?.permissionIndicators?.filter((p) => p.scope === 'Library') ?? [];
    const libraryCount = libraryIndicators.length;
    const uniquePermCount = libraryIndicators.filter((p) => p.hasUniquePermissions).length;

    const riskScore = scanResult?.riskScore;
    const riskColor = riskScore ? (riskColorMap[riskScore.label] || '#6b7280') : '#6b7280';
    const riskAngle = riskScore ? riskScore.score * 3.6 : 0;

    /* ── findings ── */
    const findings = React.useMemo(() => {
        if (!scanResult) return [];
        const items: Array<{ severity: 'success' | 'warning' | 'info' | 'danger'; icon: string; title: string; description: string; badge: string }> = [];

        if (externalUserCount === 0) {
            items.push({ severity: 'success', icon: '✓', title: 'No external users detected', description: 'No external users were found in the SharePoint groups scanned.', badge: 'Good' });
        } else {
            items.push({ severity: 'danger', icon: '!', title: `${externalUserCount} possible external user(s)`, description: 'External users detected — review access promptly.', badge: 'Action' });
        }

        if (uniquePermCount > 0) {
            items.push({ severity: 'warning', icon: '!', title: 'Libraries need review', description: `${uniquePermCount} librar${uniquePermCount === 1 ? 'y has' : 'ies have'} unique permissions and may increase governance complexity.`, badge: 'Review' });
        }

        if (groupCount > 0) {
            items.push({ severity: 'info', icon: 'i', title: 'Groups scanned', description: `${groupCount} SharePoint groups were scanned for possible external users.`, badge: 'Info' });
        }

        if (scanResult.errors.length > 0) {
            items.push({ severity: 'warning', icon: '!', title: 'Scan completed with errors', description: `${scanResult.errors.length} error(s) — some data may be incomplete.`, badge: 'Review' });
        }

        return items;
    }, [scanResult, externalUserCount, uniquePermCount, groupCount]);

    /* ── severity CSS map ── */
    const severityFindingClass: Record<string, string> = {
        success: styles.findingSuccess,
        warning: styles.findingWarning,
        info: styles.findingInfo,
        danger: styles.findingDanger,
    };
    const severityBadgeClass: Record<string, string> = {
        success: styles.badgeSuccess,
        warning: styles.badgeWarning,
        info: styles.badgeInfo,
        danger: styles.badgeDanger,
    };

    /* ────────── JSX ────────── */

    return (
        <div className={styles.externalSharingRiskScanner}>
            <div className={styles.shell}>
                {/* ── Header ── */}
                <section className={styles.header}>
                    <div>
                        <div className={styles.kicker}>🛡️ SharePoint governance tool</div>
                        <h1 className={styles.title}>External Sharing Risk Scanner</h1>
                        <p className={styles.description}>
                            Review external sharing, SharePoint group membership, and permission inheritance
                            indicators for the current SharePoint site. This public MVP is read-only and designed to
                            support governance conversations, migration readiness reviews, and Copilot readiness checks.
                        </p>
                    </div>
                    <div className={styles.headerActions}>
                        {hasScanned && (
                            <button
                                className={`${styles.button} ${styles.buttonSecondary}`}
                                type="button"
                                onClick={() => setShowDetails(!showDetails)}
                            >
                                {showDetails ? 'Hide details' : 'View details'}
                            </button>
                        )}
                        <button
                            className={`${styles.button} ${styles.buttonPrimary}`}
                            type="button"
                            onClick={runScan}
                            disabled={isScanning}
                        >
                            {isScanning ? 'Scanning…' : 'Scan site'}
                        </button>
                    </div>
                </section>

                {/* ── Warning alert ── */}
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

                {/* ── Scan errors (detail toggle) ── */}
                {showDetails && scanResult && scanResult.errors.length > 0 && (
                    <div className={`${styles.alert} ${styles.alertWarning}`} role="alert">
                        <span className={styles.alertIcon}>⚠️</span>
                        <div>
                            <strong>{scanResult.errors.length} error(s) during scan:</strong>
                            <ul style={{ margin: '4px 0 0 16px', padding: 0 }}>
                                {scanResult.errors.map((e, i) => (
                                    <li key={i}>{e.source}: {e.message}</li>
                                ))}
                            </ul>
                        </div>
                    </div>
                )}

                {/* ── Results ── */}
                {scanResult && !isScanning && (
                    <>
                        {/* ── Overview Grid ── */}
                        <section className={styles.overviewGrid} aria-label="Risk overview">
                            {/* Risk Score Card */}
                            <article className={`${styles.card}`}>
                                <div className={styles.cardHeader} style={{ width: '100%' }}>
                                    <div>
                                        <h2 className={styles.cardTitle}>Risk Score</h2>
                                        <p className={styles.cardSubtitle}>Last scanned: {scanResult.siteSummary.scannedAt}</p>
                                    </div>
                                    <span className={`${styles.badge} ${riskBadgeClass(riskScore!.label)}`}>
                                        {riskScore!.label}
                                    </span>
                                </div>
                                <div className={styles.cardBody}>
                                    <div className={styles.riskScoreCard}>
                                        <div
                                            className={styles.riskRing}
                                            aria-label={`Risk score ${riskScore!.score} out of 100`}
                                            style={{
                                                background: `conic-gradient(${riskColor} 0deg ${riskAngle}deg, #edf2f7 ${riskAngle}deg 360deg)`,
                                            }}
                                        >
                                            <div className={styles.riskRingInner} />
                                            <div className={styles.riskRingContent}>
                                                <div className={styles.riskNumber} style={{ color: riskColor }}>
                                                    {riskScore!.score}
                                                </div>
                                                <div className={styles.riskLabel}>/100 Risk</div>
                                            </div>
                                        </div>

                                        <p className={styles.summaryText}>
                                            {riskScore!.reasons[0] || 'No significant risk indicators detected.'}
                                        </p>

                                        <ul className={styles.metaList}>
                                            <li className={styles.metaItem}>
                                                <span className={styles.metaLabel}>Current site</span>
                                                <span className={styles.metaValue}>{scanResult.siteSummary.title}</span>
                                            </li>
                                            <li className={styles.metaItem}>
                                                <span className={styles.metaLabel}>Permission level</span>
                                                <span className={styles.metaValue}>
                                                    {scanResult.siteSummary.currentUserPermissionLevel || 'Unknown'}
                                                </span>
                                            </li>
                                            <li className={styles.metaItem}>
                                                <span className={styles.metaLabel}>Risk status</span>
                                                <span className={styles.metaValue}>
                                                    {riskScore!.label === 'Low' ? 'Healthy' : 'Review needed'}
                                                </span>
                                            </li>
                                        </ul>
                                    </div>
                                </div>
                            </article>

                            {/* Scan Summary Card */}
                            <article className={styles.card}>
                                <div className={styles.cardHeader}>
                                    <div>
                                        <h2 className={styles.cardTitle}>Scan Summary</h2>
                                        <p className={styles.cardSubtitle}>Key indicators from the current SharePoint site</p>
                                    </div>
                                </div>
                                <div className={styles.cardBody}>
                                    <div className={styles.metricsGrid}>
                                        <MetricTile icon="👥" value={externalUserCount} label="Possible external users" note="External users detected in scanned groups." />
                                        <MetricTile icon="🧩" value={groupCount} label="Groups reviewed" note="SharePoint groups scanned for external access." />
                                        <MetricTile icon="📚" value={libraryCount} label="Libraries reviewed" note="Document libraries checked for inheritance status." />
                                        <MetricTile icon="🔓" value={uniquePermCount} label="Unique permission sets" note="Libraries with broken inheritance need review." />
                                    </div>
                                </div>
                            </article>
                        </section>

                        {/* ── Key Findings ── */}
                        {findings.length > 0 && (
                            <section className={styles.sectionGrid}>
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
                                                <div key={idx} className={`${styles.finding} ${severityFindingClass[f.severity] || ''}`}>
                                                    <div className={styles.findingIcon}>{f.icon}</div>
                                                    <div>
                                                        <p className={styles.findingTitle}>{f.title}</p>
                                                        <p className={styles.findingDescription}>{f.description}</p>
                                                    </div>
                                                    <span className={`${styles.badge} ${severityBadgeClass[f.severity] || ''}`}>
                                                        {f.badge}
                                                    </span>
                                                </div>
                                            ))}
                                        </div>
                                    </div>
                                </article>
                            </section>
                        )}

                        {/* ── Two-Column Grid: Groups + Permissions ── */}
                        <section className={styles.twoColumnGrid}>
                            <GroupsTable groups={scanResult.groups} />
                            <PermissionsTable indicators={libraryIndicators} />
                        </section>

                        {/* ── Recommended Actions ── */}
                        {riskScore && riskScore.recommendedActions.length > 0 && (
                            <section className={styles.sectionGrid} style={{ marginTop: 16 }}>
                                <article className={styles.card}>
                                    <div className={styles.cardHeader}>
                                        <div>
                                            <h2 className={styles.cardTitle}>Recommended Actions</h2>
                                            <p className={styles.cardSubtitle}>Suggested governance follow-up steps</p>
                                        </div>
                                    </div>
                                    <div className={styles.cardBody}>
                                        <ol className={styles.actionsList}>
                                            {riskScore.recommendedActions.map((action, idx) => (
                                                <li key={idx} className={styles.actionItem}>
                                                    <span className={styles.actionNumber}>{idx + 1}</span>
                                                    <div>
                                                        <p className={styles.actionTitle}>{action}</p>
                                                    </div>
                                                </li>
                                            ))}
                                        </ol>
                                    </div>
                                </article>
                            </section>
                        )}

                        {/* ── Footer Note ── */}
                        <div className={styles.footerNote}>
                            Scanner results provide indicators for governance review. For a full assessment, validate
                            results with the SharePoint admin center, Microsoft Purview, audit logs, and tenant-level
                            sharing settings.
                        </div>
                    </>
                )}

                {/* ── Pre-scan (no results yet) ── */}
                {!hasScanned && !isScanning && (
                    <div className={`${styles.alert} ${styles.alertInfo}`} role="status">
                        <span className={styles.alertIcon}>ℹ️</span>
                        <div>
                            Click <strong>Scan site</strong> to begin analyzing external sharing risks for this
                            SharePoint site.
                        </div>
                    </div>
                )}
            </div>
        </div>
    );
};

/* ─────────── Small inline sub-components ─────────── */

interface IMetricTileProps {
    icon: string;
    value: number;
    label: string;
    note: string;
}

const MetricTile: React.FC<IMetricTileProps> = ({ icon, value, label, note }) => (
    <div className={styles.metricCard}>
        <div className={styles.metricIcon}>{icon}</div>
        <p className={styles.metricValue}>{value}</p>
        <p className={styles.metricLabel}>{label}</p>
        <p className={styles.metricNote}>{note}</p>
    </div>
);

interface IGroupsTableProps {
    groups: IGroupSummary[];
}

const GroupsTable: React.FC<IGroupsTableProps> = ({ groups }) => {
    const getStatus = (g: IGroupSummary): { text: string; cls: string } => {
        if (g.possibleExternalUserCount && g.possibleExternalUserCount > 0) {
            return { text: 'Review', cls: styles.statusReview };
        }
        if (g.isEmpty) {
            return { text: 'Empty', cls: styles.statusOk };
        }
        return { text: 'OK', cls: styles.statusOk };
    };

    return (
        <article className={styles.card}>
            <div className={styles.cardHeader}>
                <div>
                    <h2 className={styles.cardTitle}>SharePoint Groups</h2>
                    <p className={styles.cardSubtitle}>Group membership review summary</p>
                </div>
            </div>
            <div className={styles.cardBody}>
                <div className={styles.tableWrap}>
                    <table className={styles.dataTable}>
                        <thead>
                            <tr>
                                <th>Group name</th>
                                <th>Members</th>
                                <th>Possible external</th>
                                <th>Status</th>
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
                                    </tr>
                                );
                            })}
                            {groups.length === 0 && (
                                <tr><td colSpan={4} style={{ textAlign: 'center', color: '#6b7280' }}>No groups found</td></tr>
                            )}
                        </tbody>
                    </table>
                </div>
            </div>
        </article>
    );
};

interface IPermissionsTableProps {
    indicators: IPermissionIndicator[];
}

const PermissionsTable: React.FC<IPermissionsTableProps> = ({ indicators }) => (
    <article className={styles.card}>
        <div className={styles.cardHeader}>
            <div>
                <h2 className={styles.cardTitle}>Permission Inheritance</h2>
                <p className={styles.cardSubtitle}>Document library inheritance status</p>
            </div>
        </div>
        <div className={styles.cardBody}>
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
                                <td>{ind.hasUniquePermissions ? 'Broken' : 'Inherited'}</td>
                                <td>
                                    <span className={ind.status === 'Ok' ? styles.statusOk : styles.statusReview}>
                                        {ind.status === 'Ok' ? 'OK' : 'Review needed'}
                                    </span>
                                </td>
                            </tr>
                        ))}
                        {indicators.length === 0 && (
                            <tr><td colSpan={3} style={{ textAlign: 'center', color: '#6b7280' }}>No libraries found</td></tr>
                        )}
                    </tbody>
                </table>
            </div>
        </div>
    </article>
);

export { ExternalSharingRiskScanner };
export default ExternalSharingRiskScanner;
