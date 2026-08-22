/**
 * Main-process logger — parity target for the C# in-app log sink (AppendLog) plus file diagnostics.
 *
 * Stage-2 scope: leveled logging with a size-capped rotating file under userData/logs and an in-memory
 * ring buffer that the future in-app log viewer (FEATURE_PARITY_MATRIX 15.2) will read over IPC.
 */
import { app } from "electron";
import { appendFileSync, mkdirSync, statSync, renameSync, existsSync } from "node:fs";
import { join } from "node:path";

export type LogLevel = "debug" | "info" | "warn" | "error";

const LEVEL_ORDER: Record<LogLevel, number> = { debug: 0, info: 1, warn: 2, error: 3 };
export const MAX_LOG_BYTES = 5 * 1024 * 1024;
export const RING_BUFFER_CAPACITY = 1000;

export interface LogEntry {
  readonly ts: string;
  readonly level: LogLevel;
  readonly scope: string;
  readonly message: string;
}

class Logger {
  private ring: LogEntry[] = [];
  private filePath: string | null = null;

  init(): void {
    try {
      const dir = join(app.getPath("userData"), "logs");
      mkdirSync(dir, { recursive: true });
      this.filePath = join(dir, "main.log");
      this.rotateIfNeeded();
    } catch {
      // Logging must never crash the app; file sink stays disabled.
      this.filePath = null;
    }
  }

  /** Recent entries for the future in-app log viewer. Newest last. */
  recent(count = 200): LogEntry[] {
    return this.ring.slice(Math.max(0, this.ring.length - count));
  }

  log(level: LogLevel, scope: string, message: string): void {
    const entry: LogEntry = { ts: new Date().toISOString(), level, scope, message };
    this.ring.push(entry);
    if (this.ring.length > RING_BUFFER_CAPACITY) this.ring.shift();

    const line = `[${entry.ts}] [${level.toUpperCase()}] [${scope}] ${message}`;
    if (LEVEL_ORDER[level] >= LEVEL_ORDER.info) {
      // eslint-disable-next-line no-console
      console[level === "warn" ? "warn" : level === "error" ? "error" : "log"](line);
    }
    if (this.filePath) {
      try {
        appendFileSync(this.filePath, line + "\n", "utf8");
        this.rotateIfNeeded();
      } catch {
        this.filePath = null;
      }
    }
  }

  debug(scope: string, message: string): void {
    this.log("debug", scope, message);
  }

  info(scope: string, message: string): void {
    this.log("info", scope, message);
  }

  warn(scope: string, message: string): void {
    this.log("warn", scope, message);
  }

  error(scope: string, message: string): void {
    this.log("error", scope, message);
  }

  private rotateIfNeeded(): void {
    if (!this.filePath || !existsSync(this.filePath)) return;
    if (statSync(this.filePath).size < MAX_LOG_BYTES) return;
    const rotated = this.filePath.replace(/\.log$/, ".1.log");
    try {
      if (existsSync(rotated)) rmSilent(rotated);
      renameSync(this.filePath, rotated);
    } catch {
      /* rotation is best-effort */
    }
  }
}

function rmSilent(path: string): void {
  try {
    unlinkSyncShim(path);
  } catch {
    /* ignore */
  }
}

// Local import kept at bottom to avoid top-level node:fs spread above.
import { unlinkSync as unlinkSyncShim } from "node:fs";

export const logger = new Logger();
