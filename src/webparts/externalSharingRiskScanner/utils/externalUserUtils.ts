import {
    EXTERNAL_LOGIN_PATTERNS,
    EXTERNAL_DISPLAY_PATTERNS,
    isSystemAccount,
} from '../constants/ExternalUserPatterns';

export interface IExternalDetectionResult {
    isPossibleExternal: boolean;
    reasons: string[];
}

/**
 * Determines if a user is possibly an external (guest) user based on heuristics.
 *
 * @param loginName - The user's login/claims name
 * @param email - The user's email address (if available)
 * @param displayName - The user's display name
 * @param internalDomains - Array of internal domain strings (lowercase)
 * @returns Detection result with flag and reasons
 */
export function detectExternalUser(
    loginName: string,
    email: string | undefined,
    displayName: string,
    internalDomains: string[]
): IExternalDetectionResult {
    const reasons: string[] = [];

    // Skip system accounts entirely
    if (isSystemAccount(loginName || '')) {
        return { isPossibleExternal: false, reasons: [] };
    }

    // Check login name patterns
    for (const check of EXTERNAL_LOGIN_PATTERNS) {
        if (check.pattern.test(loginName || '')) {
            reasons.push(check.reason);
        }
    }

    // Check display name patterns (secondary signal)
    for (const check of EXTERNAL_DISPLAY_PATTERNS) {
        if (check.pattern.test(displayName || '')) {
            reasons.push(check.reason);
        }
    }

    // Check email domain against internal domains
    if (email && internalDomains.length > 0) {
        const emailDomain = extractEmailDomain(email);
        if (emailDomain && !internalDomains.includes(emailDomain.toLowerCase())) {
            reasons.push(
                `Email domain "${emailDomain}" is not in the configured internal domains list`
            );
        }
    }

    // If no email available but login contains an email-like pattern, check that
    if (!email && loginName) {
        const extractedEmail = extractEmailFromLoginName(loginName);
        if (extractedEmail && internalDomains.length > 0) {
            const emailDomain = extractEmailDomain(extractedEmail);
            if (emailDomain && !internalDomains.includes(emailDomain.toLowerCase())) {
                reasons.push(
                    `Extracted email domain "${emailDomain}" from login name is not in internal domains`
                );
            }
        }
    }

    return {
        isPossibleExternal: reasons.length > 0,
        reasons,
    };
}

/**
 * Extracts the domain portion from an email address.
 */
export function extractEmailDomain(email: string): string | undefined {
    if (!email) return undefined;
    const atIndex = email.lastIndexOf('@');
    if (atIndex < 0) return undefined;
    return email.substring(atIndex + 1).toLowerCase().trim();
}

/**
 * Attempts to extract an email address from a claims-based login name.
 * Common patterns:
 *   i:0#.f|membership|user@domain.com
 *   i:0#.f|membership|user_domain.com#ext#@tenant.onmicrosoft.com
 */
export function extractEmailFromLoginName(loginName: string): string | undefined {
    if (!loginName) return undefined;

    // Try to find email after the last pipe character
    const parts = loginName.split('|');
    if (parts.length > 1) {
        const lastPart = parts[parts.length - 1];
        if (lastPart.includes('@')) {
            return lastPart.trim();
        }
    }

    // Try to find email directly in the string
    const emailMatch = loginName.match(/[\w.+-]+@[\w.-]+\.\w+/);
    if (emailMatch) {
        return emailMatch[0];
    }

    return undefined;
}

/**
 * Parses a comma-separated domain string into a clean array of lowercase domains.
 */
export function parseInternalDomains(domainsString: string | undefined): string[] {
    if (!domainsString || domainsString.trim() === '') {
        return [];
    }
    return domainsString
        .split(',')
        .map((d) => d.trim().toLowerCase())
        .filter((d) => d.length > 0);
}
