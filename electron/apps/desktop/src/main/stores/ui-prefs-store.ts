/**
 * Persisted shell preferences (sidebar pinned/width) — first consumer of JsonStore.
 * Parity anchor: the WPF app persisted sidebar state in settings; defaults 420 px width, pinned open
 * (audit UI_SHELL §2).
 */
import { join } from "node:path";
import { JsonStore, type LoadOutcome } from "./json-store.js";
import type { UiPrefs } from "@rpd/shared";
import { uiPrefsSchema } from "@rpd/shared";

export const UI_PREFS_DEFAULTS: UiPrefs = {
  sidebarPinned: true,
  sidebarWidth: 420,
};

const SCHEMA_VERSION = 1;

export class UiPrefsStore {
  private readonly json: JsonStore<UiPrefs & { schemaVersion: number }>;
  private cache: UiPrefs | null = null;

  constructor(userDataDir: string, log?: (level: "warn" | "error", message: string) => void) {
    this.json = new JsonStore<UiPrefs & { schemaVersion: number }>({
      file: join(userDataDir, "ui-prefs.json"),
      schemaVersion: SCHEMA_VERSION,
      validate: (doc): doc is UiPrefs & { schemaVersion: number } => uiPrefsSchema.safeParse(doc).success,
      log,
    });
  }

  get(): UiPrefs {
    if (this.cache) return this.cache;
    const outcome: LoadOutcome<UiPrefs & { schemaVersion: number }> = this.json.load();
    if (outcome.status === "loaded") {
      this.cache = { sidebarPinned: outcome.doc.sidebarPinned, sidebarWidth: outcome.doc.sidebarWidth };
    } else {
      // missing / quarantined / migrate-failed → defaults (quarantine preserved the bytes on disk).
      this.cache = { ...UI_PREFS_DEFAULTS };
    }
    return this.cache;
  }

  set(patch: Partial<UiPrefs>): UiPrefs {
    const next: UiPrefs = { ...this.get(), ...patch };
    this.json.save({ ...next, schemaVersion: SCHEMA_VERSION });
    this.cache = next;
    return next;
  }
}
