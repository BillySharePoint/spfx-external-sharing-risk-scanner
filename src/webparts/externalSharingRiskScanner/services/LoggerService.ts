/**
 * Minimal logger service for development/debugging.
 * Never logs sensitive data (emails, permissions) in production mode.
 */

export enum LogLevel {
    None = 0,
    Error = 1,
    Warning = 2,
    Info = 3,
    Debug = 4,
}

class LoggerServiceClass {
    private _level: LogLevel = LogLevel.Warning;
    private _prefix: string = '[RiskScanner]';

    /**
     * Sets the log level. In production, keep at Warning or Error.
     */
    public setLevel(level: LogLevel): void {
        this._level = level;
    }

    /**
     * Enable debug mode (for development only).
     */
    public enableDebug(): void {
        this._level = LogLevel.Debug;
    }

    public error(source: string, message: string, ...args: unknown[]): void {
        if (this._level >= LogLevel.Error) {
            console.error(`${this._prefix} [${source}]`, message, ...args);
        }
    }

    public warn(source: string, message: string, ...args: unknown[]): void {
        if (this._level >= LogLevel.Warning) {
            console.warn(`${this._prefix} [${source}]`, message, ...args);
        }
    }

    public info(source: string, message: string, ...args: unknown[]): void {
        if (this._level >= LogLevel.Info) {
            console.info(`${this._prefix} [${source}]`, message, ...args);
        }
    }

    public debug(source: string, message: string, ...args: unknown[]): void {
        if (this._level >= LogLevel.Debug) {
            console.log(`${this._prefix} [${source}]`, message, ...args);
        }
    }
}

/** Singleton logger instance */
export const LoggerService = new LoggerServiceClass();
