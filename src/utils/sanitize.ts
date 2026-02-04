import * as DOMPurify from 'dompurify';

/**
 * Sanitizes HTML content to prevent XSS attacks
 * @param dirty - Potentially unsafe HTML string
 * @returns Sanitized HTML safe for rendering
 */
export function sanitizeHtml(dirty: string | undefined | null): string {
    if (!dirty) return '';

    return DOMPurify.sanitize(dirty, {
        ALLOWED_TAGS: ['b', 'i', 'em', 'strong', 'a', 'br', 'p', 'div', 'span', 'ul', 'ol', 'li', 'h1', 'h2', 'h3', 'h4', 'h5', 'h6'],
        ALLOWED_ATTR: ['href', 'title', 'target', 'style'],
        ALLOW_DATA_ATTR: false,
        KEEP_CONTENT: true
    });
}

/**
 * Escapes HTML special characters for plain text display
 * @param text - Text to escape
 * @returns Escaped text
 */
export function escapeHtml(text: string | undefined | null): string {
    if (!text) return '';

    return text
        .replace(/&/g, "&amp;")
        .replace(/</g, "&lt;")
        .replace(/>/g, "&gt;")
        .replace(/"/g, "&quot;")
        .replace(/'/g, "&#039;");
}
