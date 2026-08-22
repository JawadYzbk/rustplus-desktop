/**
 * Versioned JSON store base — the durability primitive for every persisted document.
 *
 * Guarantees (ELECTRON_ARCHITECTURE §6, fixing audit DATA_STORES risks):
 *  - Atomic writes: content is written to a temp file in the same directory, fsynced, then renamed over
 *    the target. A crash mid-write can never leave a half-written store.
 *  - schemaVersion on every document. Older versions go through an explicit migrate hook; newer versions
 *    (downgrade) are never guessed at.
 *  - Non-destructive failure: any unreadable/invalid/mismatched file is MOVED to `corrupt/` (never deleted),
 *    so data is recoverable by hand. Load then reports the outcome instead of silently returning defaults.
 */
import { dirname, basename, join } from "node:path";
import {
  closeSync,
  existsSync,
  mkdirSync,
  openSync,
  readFileSync,
  renameSync,
  writeSync,
} from "node:fs";

export interface VersionedDocument {
  schemaVersion: number;
}

export type LoadOutcome<T> =
  | { status: "loaded"; doc: T }
  | { status: "missing" }
  | { status: "quarantined"; quarantinePath: string; reason: string }
  | { status: "migrate-failed"; reason: string };

export interface JsonStoreOptions<T extends VersionedDocument> {
  /** Absolute path of the store file. */
  file: string;
  /** Current schema version written on every save. */
  schemaVersion: number;
  /** Structural validation for the current version (typically a zod safeParse wrapper). */
  validate: (doc: unknown) => doc is T;
  /**
   * Bring an older-version document up to the current version. Return null to declare it unrecoverable.
   * The original file is never modified or removed by migration failure.
   */
  migrate?: (raw: Record<string, unknown>, fromVersion: number) => T | null;
  /** Sink for diagnostics (main logger). */
  log?: (level: "warn" | "error", message: string) => void;
}

export class JsonStore<T extends VersionedDocument> {
  private readonly opts: JsonStoreOptions<T>;

  constructor(opts: JsonStoreOptions<T>) {
    this.opts = opts;
  }

  load(): LoadOutcome<T> {
    const { file, schemaVersion, validate, migrate, log } = this.opts;

    if (!existsSync(file)) return { status: "missing" };

    let parsed: unknown;
    try {
      parsed = JSON.parse(readFileSync(file, "utf8"));
    } catch (err) {
      const reason = `invalid JSON: ${err instanceof Error ? err.message : String(err)}`;
      const quarantinePath = this.quarantine();
      log?.("error", `${file}: ${reason}; quarantined to ${quarantinePath}`);
      return { status: "quarantined", quarantinePath, reason };
    }

    const version = extractVersion(parsed);

    if (version === schemaVersion) {
      if (validate(parsed)) return { status: "loaded", doc: parsed };
      const reason = "failed current-version validation";
      const quarantinePath = this.quarantine();
      log?.("error", `${file}: ${reason}; quarantined to ${quarantinePath}`);
      return { status: "quarantined", quarantinePath, reason };
    }

    if (version > schemaVersion) {
      const reason = `future schemaVersion ${version} (app knows ${schemaVersion}) — refusing to guess`;
      const quarantinePath = this.quarantine();
      log?.("error", `${file}: ${reason}; quarantined to ${quarantinePath}`);
      return { status: "quarantined", quarantinePath, reason };
    }

    // Older version: explicit migration or nothing.
    if (!migrate) {
      const reason = `legacy schemaVersion ${version} with no migrator registered`;
      const quarantinePath = this.quarantine();
      log?.("warn", `${file}: ${reason}; quarantined to ${quarantinePath}`);
      return { status: "quarantined", quarantinePath, reason };
    }

    if (!isRecord(parsed)) {
      const reason = `legacy document is not an object (schemaVersion ${version})`;
      const quarantinePath = this.quarantine();
      log?.("error", `${file}: ${reason}; quarantined to ${quarantinePath}`);
      return { status: "quarantined", quarantinePath, reason };
    }

    const migrated = migrate(parsed, version);
    if (migrated === null || !validate(migrated)) {
      const reason = `migration from schemaVersion ${version} failed`;
      log?.("error", `${file}: ${reason}; original file left untouched`);
      return { status: "migrate-failed", reason };
    }
    return { status: "loaded", doc: migrated };
  }

  /**
   * Atomic save: temp file in the same directory (same volume ⇒ rename is atomic), fsync, rename over.
   * Throws on failure — callers decide retry policy; nothing is silently dropped.
   */
  save(doc: T): void {
    const { file } = this.opts;
    const dir = dirname(file);
    mkdirSync(dir, { recursive: true });

    const payload = JSON.stringify(doc, null, 2);
    const tmp = `${file}.tmp`;
    const fd = openSync(tmp, "w");
    try {
      writeSync(fd, payload, "utf8");
      fsyncSync(fd); // flush to disk before the atomic rename
    } finally {
      closeSync(fd);
    }
    renameSync(tmp, file);
  }

  /** Move the current file into `<dir>/corrupt/<name>.<ts>` and return the new path. */
  private quarantine(): string {
    const { file } = this.opts;
    const dir = join(dirname(file), "corrupt");
    mkdirSync(dir, { recursive: true });
    const stamp = new Date().toISOString().replace(/[:.]/g, "-");
    const target = join(dir, `${basename(file)}.${stamp}`);
    try {
      renameSync(file, target);
      return target;
    } catch (err) {
      // If even the quarantine move fails, leave the original in place and say so.
      return `${target} (move failed: ${err instanceof Error ? err.message : String(err)}; original kept at ${file})`;
    }
  }
}

function fsyncFd(fd: number): void {
  // eslint-disable-next-line @typescript-eslint/no-var-requires
  const { fsyncSync } = require("node:fs") as { fsyncSync(fd: number): void };
  fsyncSync(fd);
}

function extractVersion(parsed: unknown): number {
  if (isRecord(parsed) && typeof parsed["schemaVersion"] === "number") {
    return parsed["schemaVersion"];
  }
  // No version marker at all ⇒ treat as version 0 (oldest possible) so migrators can claim it.
  return 0;
}

function isRecord(value: unknown): value is Record<string, unknown> {
  return typeof value === "object" && value !== null && !Array.isArray(value);
}
