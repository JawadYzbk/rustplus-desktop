/**
 * TrackingSettings store — tracking_settings.json parity.
 *
 * Behavior deltas vs legacy (deliberate, documented in ELECTRON_ARCHITECTURE §6):
 *  - atomic writes + schemaVersion instead of bare WriteAllText;
 *  - unknown keys are a validation failure (quarantine), so key renames can't silently fork state;
 *  - missing file → exact legacy defaults (same as legacy loader's catch→defaults path).
 */
import { join } from "node:path";
import { JsonStore } from "./json-store.js";
import { TRACKING_SETTINGS_DEFAULTS, trackingSettingsSchema, type TrackingSettings } from "@rpd/shared";

const SCHEMA_VERSION = 1;

type StoredDoc = TrackingSettings & { schemaVersion: number };

export class SettingsStore {
  private readonly json: JsonStore<StoredDoc>;
  private cache: TrackingSettings | null = null;

  constructor(userDataDir: string, log?: (level: "warn" | "error", message: string) => void) {
    this.json = new JsonStore<StoredDoc>({
      file: join(userDataDir, "tracking_settings.json"),
      schemaVersion: SCHEMA_VERSION,
      validate: (doc): doc is StoredDoc => {
        if (!trackingSettingsSchema.safeParse(stripVersion(doc)).success) return false;
        return isRecord(doc) && doc["schemaVersion"] === SCHEMA_VERSION;
      },
      log,
    });
  }

  get all(): TrackingSettings {
    if (this.cache) return this.cache;
    const outcome = this.json.load();
    if (outcome.status === "loaded") {
      const { schemaVersion: _v, ...settings } = outcome.doc;
      this.cache = settings;
    } else {
      this.cache = { ...TRACKING_SETTINGS_DEFAULTS };
    }
    return this.cache;
  }

  /** Patch selected keys and persist atomically. Unknown keys throw (contract breach). */
  patch(patch: Partial<TrackingSettings>): TrackingSettings {
    const merged: TrackingSettings = { ...this.all, ...patch };
    const parsed = trackingSettingsSchema.safeParse(merged);
    if (!parsed.success) {
      const issue = parsed.error.issues[0];
      throw new Error(`settings patch breached contract at ${issue?.path.join(".") ?? "?"}: ${issue?.message ?? "unknown"}`);
    }
    this.json.save({ ...merged, schemaVersion: SCHEMA_VERSION });
    this.cache = merged;
    return merged;
  }
}

function stripVersion(doc: unknown): unknown {
  if (typeof doc !== "object" || doc === null) return doc;
  const { schemaVersion: _v, ...rest } = doc as Record<string, unknown>;
  void _v;
  return rest;
}

function isRecord(value: unknown): value is Record<string, unknown> {
  return typeof value === "object" && value !== null && !Array.isArray(value);
}
