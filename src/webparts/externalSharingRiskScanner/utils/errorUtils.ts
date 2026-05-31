import { IScannerError } from '../models/IScannerError';

/**
 * Normalizes any caught error into a structured IScannerError.
 *
 * @param error - The caught error (could be Error, string, or unknown)
 * @param source - The logical source/service that produced the error
 * @returns A structured IScannerError
 */
export function normalizeError(error: unknown, source: string): IScannerError {
    if (error instanceof Error) {
        return {
            source,
            message: error.message || 'An unknown error occurred.',
            technicalDetails: error.stack,
        };
    }

    if (typeof error === 'string') {
        return {
            source,
            message: error,
        };
    }

    if (error && typeof error === 'object') {
        const errObj = error as Record<string, unknown>;
        return {
            source,
            message: (errObj.message as string) || (errObj.statusText as string) || 'An error occurred.',
            technicalDetails: JSON.stringify(error, null, 2),
        };
    }

    return {
        source,
        message: 'An unexpected error occurred.',
        technicalDetails: String(error),
    };
}

/**
 * Returns a user-friendly error message suitable for display.
 * Strips technical jargon and provides actionable guidance.
 */
export function getFriendlyErrorMessage(error: IScannerError): string {
    const msg = error.message.toLowerCase();

    if (msg.includes('403') || msg.includes('access denied') || msg.includes('unauthorized')) {
        return 'You do not have sufficient permissions to access this data. Try running the scanner as a site owner or admin.';
    }

    if (msg.includes('404') || msg.includes('not found')) {
        return 'The requested resource was not found. The site or list may have been removed.';
    }

    if (msg.includes('429') || msg.includes('throttl')) {
        return 'The request was throttled by SharePoint. Please wait a moment and try again.';
    }

    if (msg.includes('timeout') || msg.includes('timed out')) {
        return 'The request timed out. The site may be experiencing high load.';
    }

    if (msg.includes('network') || msg.includes('fetch')) {
        return 'A network error occurred. Please check your connection and try again.';
    }

    return error.message || 'An unexpected error occurred while scanning.';
}
