/**
 * Patterns and checks for detecting possible external users.
 * These are heuristics — not definitive security auditing.
 */

/**
 * Login name patterns that indicate a possible external/guest user.
 */
export const EXTERNAL_LOGIN_PATTERNS: Array<{ pattern: RegExp; reason: string }> = [
    {
        pattern: /#ext#/i,
        reason: 'Login name contains #EXT# (Azure AD guest pattern)',
    },
    {
        pattern: /^i:0#\.f\|membership\|.*#ext#/i,
        reason: 'Claims-based login with #EXT# identifier',
    },
    {
        pattern: /_urn:spo:guest/i,
        reason: 'Login contains SharePoint guest URN pattern',
    },
    {
        pattern: /^c:0[%-]\.c\|t\|/i,
        reason: 'External sharing claim detected in login name',
    },
    {
        pattern: /urn%3aspo%3aanon/i,
        reason: 'Anonymous sharing link claim detected',
    },
];

/**
 * Display name patterns that may indicate external/system accounts.
 * Used as secondary signals, not primary detection.
 */
export const EXTERNAL_DISPLAY_PATTERNS: Array<{ pattern: RegExp; reason: string }> = [
    {
        pattern: /\(External\)/i,
        reason: 'Display name contains "(External)" label',
    },
    {
        pattern: /\(Guest\)/i,
        reason: 'Display name contains "(Guest)" label',
    },
];

/**
 * System accounts to exclude from external user detection.
 * These are not real users and should not be flagged.
 */
export const SYSTEM_ACCOUNT_PATTERNS: RegExp[] = [
    /^SHAREPOINT\\system$/i,
    /^i:0#\.w\|nt authority\\local service$/i,
    /^c:0\('\.s\|true$/i,
    /^app@sharepoint$/i,
    /^SHAREPOINT\\.*$/i,
];

/**
 * Checks whether a login name matches known system account patterns.
 */
export function isSystemAccount(loginName: string): boolean {
    return SYSTEM_ACCOUNT_PATTERNS.some((pattern) => pattern.test(loginName));
}
