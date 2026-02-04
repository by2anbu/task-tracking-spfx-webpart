/**
 * Conditional logger that only logs in development mode
 * Prevents information disclosure in production
 */
export class Logger {
    private static isDevelopment =
        window.location.hostname === 'localhost' ||
        window.location.hostname.includes('workbench') ||
        window.location.hostname.includes('sharepoint.com');

    /**
     * Log informational messages (development only)
     */
    static log(...args: any[]): void {
        if (this.isDevelopment) {
            console.log(...args);
        }
    }

    /**
     * Log warning messages (development only)
     */
    static warn(...args: any[]): void {
        if (this.isDevelopment) {
            console.warn(...args);
        }
    }

    /**
     * Log error messages (always logged for debugging)
     */
    static error(...args: any[]): void {
        console.error(...args);
    }

    /**
     * Log debug messages (development only)
     */
    static debug(...args: any[]): void {
        if (this.isDevelopment) {
            console.debug(...args);
        }
    }
}
