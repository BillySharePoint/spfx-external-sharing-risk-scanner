/**
 * Date/time utilities for the scanner.
 */

/**
 * Returns the current date/time formatted for display.
 * Uses the browser's locale for user-friendly formatting.
 */
export function getFormattedScanTime(): string {
    return new Date().toLocaleString(undefined, {
        year: 'numeric',
        month: 'short',
        day: 'numeric',
        hour: '2-digit',
        minute: '2-digit',
        second: '2-digit',
    });
}

/**
 * Returns an ISO timestamp string for data storage.
 */
export function getIsoTimestamp(): string {
    return new Date().toISOString();
}

/**
 * Formats a date string or Date object for display.
 */
export function formatDate(date: string | Date): string {
    const d = typeof date === 'string' ? new Date(date) : date;
    return d.toLocaleString(undefined, {
        year: 'numeric',
        month: 'short',
        day: 'numeric',
        hour: '2-digit',
        minute: '2-digit',
    });
}
