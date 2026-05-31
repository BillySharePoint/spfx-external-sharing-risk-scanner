/**
 * Risk scoring constants for the External Sharing Risk Scanner.
 * All thresholds and weights are transparent and deterministic.
 */

/** Score contribution weights */
export const RISK_WEIGHTS = {
    /** External users found in site */
    EXTERNAL_USERS_FOUND: 20,
    /** External users found in high-privilege groups (Owners) */
    EXTERNAL_USERS_IN_OWNERS: 25,
    /** External users found in Members group */
    EXTERNAL_USERS_IN_MEMBERS: 15,
    /** Site has unique (broken) permissions */
    SITE_UNIQUE_PERMISSIONS: 10,
    /** Multiple libraries have unique permissions */
    MULTIPLE_LIBRARIES_UNIQUE_PERMISSIONS: 15,
    /** Single library with unique permissions */
    SINGLE_LIBRARY_UNIQUE_PERMISSIONS: 5,
    /** Data retrieval failures (moves toward Unknown) */
    DATA_RETRIEVAL_FAILURE: 10,
    /** Large number of external users (>5) */
    MANY_EXTERNAL_USERS: 10,
    /** Empty groups that may indicate stale permissions */
    EMPTY_GROUPS_WITH_PERMISSIONS: 5,
};

/** Risk label thresholds based on score (0-100) */
export const RISK_THRESHOLDS = {
    LOW_MAX: 25,
    MEDIUM_MAX: 50,
    HIGH_MAX: 75,
    // Anything above HIGH_MAX = "Review Recommended"
};

/** Maximum score cap */
export const MAX_RISK_SCORE = 100;

/** Minimum score floor */
export const MIN_RISK_SCORE = 0;

/** Maximum number of libraries to scan for permissions */
export const MAX_LIBRARIES_TO_SCAN = 20;

/** External user count thresholds */
export const EXTERNAL_USER_THRESHOLDS = {
    /** Trigger "many external users" weight */
    MANY: 5,
};
