/**
 * Legacy migrator (M3) — imports C#-era stores from the legacy roots into the new Electron root.
 *
 * Source-of-truth mapping (audit DATA_STORES §1/§4/§6):
 *  - transformed: tracking_settings, profiles, hotkeys, hotkey_options, custom_alerts,
 *    tracked_players, tutorial-progress
 *  - copied verbatim (device-bound / opaque): rustplusjs-config.json, cache\*.json
 *    (minimap_settings, supabase_session kept for reference only — Laravel replaces it,
 *    notifications_history, map3d_consent, handshake_*)
 *  - deferred to their feature stages: Overlays\, 3DMaps\, player-wipes\, deaths\, raid-plan.json,
 *    WebView2 profiles, LOCALAPPDATA caches (avatars/icons/map_cache)
 *
 * Never destructive: reads legacy files only; writes only into the new root. Every source produces
 * exactly one report row.
 */
import { existsSync, mkdirSync, readdirSync, readFileSync, statSync, copyFileSync } from "node:fs";
import { join } from "node:path";
import {
  TRACKING_SETTINGS_DEFAULTS,
  TUTORIAL_PREFERENCES_DEFAULTS,
  trackingSettingsSchema,
  tutorialPreferencesSchema,
  tutorialProgressSchema,
  type LegacyRootInfo,
  type LegacySourceId,
  type LegacySourceInfo,
  type MigrationRow,
} from "@rpd/shared";
import { SettingsStore } from "../stores/settings-store.js";
import { ProfilesStore } from "../stores/profiles-store.js";
import {
  AlertTemplateStore,
  DeviceHotkeysStore,
  HotkeyOptionsStore,
  TrackedPlayersStore,
} from "../stores/legacy-stores.js";
import { TutorialProgressStore } from "../stores/tutorial-progress-store.js";

export interface LegacyRoots {
  appData: string;
  localData: string;
  deathsDir: string;
}

/** Default legacy roots (audit §1) — including the %APPDATA%\RustPlusDesktop deaths-folder quirk. */
export function defaultLegacyRoots(appData: string, localAppData: string): LegacyRoots {
  return {
    appData: join(appData, "RustPlusDesk"),
    localData: join(localAppData, "RustPlusDesk"),
    deathsDir: join(appData, "RustPlusDesktop", "deaths"),
  };
}

const COPY_THROUGH_FILES = ["rustplusjs-config.json"];

const COPY_THROUGH_CACHE_FILES = [
  "minimap_settings.json",
  "supabase_session.json",
  "notifications_history.json",
  "map3d_consent.json",
  "handshake_key.json",
  "handshake_jwt.json",
];

interface StoreBundle {
  settings: SettingsStore;
  profiles: ProfilesStore;
  hotkeys: DeviceHotkeysStore;
  hotkeyOptions: HotkeyOptionsStore;
  alerts: AlertTemplateStore;
  trackedPlayers: TrackedPlayersStore;
  tutorials: TutorialProgressStore;
}

export class LegacyMigrator {
  constructor(
    readonly roots: LegacyRoots,
    readonly userDataDir: string,
    private readonly stores: StoreBundle,
    private readonly log?: (level: "warn" | "error" | "info", message: string) => void,
  ) {}

  scan(): { roots: LegacyRootInfo[]; sources: LegacySourceInfo[] } {
    const roots: LegacyRootInfo[] = (
      [
        ["appData", this.roots.appData],
        ["localData", this.roots.localData],
        ["deathsDir", this.roots.deathsDir],
      ] as const
    ).map(([kind, path]) => ({ kind, path, exists: existsSync(path) }));

    const sources: LegacySourceInfo[] = SOURCE_DEFS.map((def) => {
      const p = join(this.roots.appData, def.rel);
      const exists = existsSync(p);
      let bytes: number | null = null;
      if (exists && def.kind === "file") bytes = statSync(p).size;
      else if (exists && def.kind === "dir") {
        try {
          bytes = readdirSync(p).length; // entry count for dirs
        } catch {
          bytes = null;
        }
      }
      return { id: def.id, label: def.label, location: def.rel, exists, bytes };
    });

    return { roots, sources };
  }

  run(): { startedAt: string; finishedAt: string; rows: MigrationRow[] } {
    const startedAt = new Date().toISOString();
    const rows: MigrationRow[] = [];
    for (const def of SOURCE_DEFS) {
      try {
        rows.push(def.execute(this));
      } catch (err) {
        const detail = err instanceof Error ? err.message : String(err);
        this.log?.("error", `migrator: ${def.id} failed: ${detail}`);
        rows.push({ source: def.label, status: "failed", detail });
      }
    }
    // Stage-owned data is reported as deferred so the UX shows the complete picture.
    for (const row of deferredRows()) rows.push(row);
    return { startedAt, finishedAt: new Date().toISOString(), rows };
  }

  // --- per-source executors (internal; called via SOURCE_DEFS) ---

  migrateSettings(): MigrationRow {
    const label = "Tracking settings";
    const file = join(this.roots.appData, "tracking_settings.json");
    if (!existsSync(file)) return { source: label, status: "missing" };

    let raw: unknown;
    try {
      raw = JSON.parse(readFileSync(file, "utf8"));
    } catch {
      return { source: label, status: "failed", detail: "not valid JSON" };
    }

    const warnings: string[] = [];
    // Permissive read: keep every key that validates, reset invalid ones to defaults, warn loudly
    // (legacy tolerated junk keys; we surface them instead of failing the whole import).
    const rawObj = (typeof raw === "object" && raw !== null && !Array.isArray(raw) ? raw : {}) as Record<
      string,
      unknown
    >;
    const merged: Record<string, unknown> = {};
    for (const [key, fallback] of Object.entries(TRACKING_SETTINGS_DEFAULTS)) {
      const candidate = key in rawObj ? rawObj[key] : fallback;
      const check = trackingSettingsSchema.pick({ [key]: true } as never).safeParse({ [key]: candidate });
      if (check.success) {
        merged[key] = candidate;
      } else if (key in rawObj) {
        merged[key] = fallback;
        warnings.push(`${key} invalid (${String(check.error.issues[0]?.message)}), reset to default`);
      } else {
        merged[key] = fallback;
      }
    }
    for (const key of Object.keys(rawObj)) {
      if (!(key in TRACKING_SETTINGS_DEFAULTS)) warnings.push(`unknown key dropped: ${key}`);
    }

    const parsed = trackingSettingsSchema.safeParse(merged);
    if (!parsed.success) {
      return {
        source: label,
        status: "failed",
        detail: `post-merge validation failed: ${parsed.error.issues[0]?.message}`,
      };
    }
    this.stores.settings.patch(parsed.data);
    return {
      source: label,
      status: warnings.length ? "warning" : "migrated",
      warnings: warnings.length ? warnings : undefined,
    };
  }

  migrateProfiles(): MigrationRow {
    const label = "Server profiles";
    const file = join(this.roots.appData, "profiles.json");
    if (!existsSync(file)) return { source: label, status: "missing" };

    let raw: unknown;
    try {
      raw = JSON.parse(readFileSync(file, "utf8"));
    } catch {
      return { source: label, status: "failed", detail: "not valid JSON" };
    }
    if (!Array.isArray(raw)) return { source: label, status: "failed", detail: "expected a JSON array of profiles" };

    let imported = 0;
    const failures: string[] = [];
    for (const [i, entry] of raw.entries()) {
      if (typeof entry !== "object" || entry === null) {
        failures.push(`profile #${i}: not an object`);
        continue;
      }
      const p = entry as Record<string, unknown>;
      if (typeof p["Host"] !== "string" || typeof p["Port"] !== "number") {
        failures.push(`profile #${i}: missing Host/Port`);
        continue;
      }
      this.stores.profiles.upsert({
        // Every legacy field passes through untouched; the core fields are normalized last
        // (later keys win) so missing/ill-typed values can't breach the contract.
        ...p,
        Name: typeof p["Name"] === "string" ? (p["Name"] as string) : "",
        Host: p["Host"] as string,
        Port: p["Port"] as number,
        SteamId64: typeof p["SteamId64"] === "string" ? (p["SteamId64"] as string) : "",
      });
      imported++;
    }
    if (failures.length) {
      return { source: label, status: "warning", detail: `${imported}/${raw.length} imported`, warnings: failures };
    }
    return { source: label, status: "migrated", detail: `${imported} profile(s), PlayerToken sealed at rest` };
  }

  migrateHotkeys(): MigrationRow {
    return migrateJsonMapFile(join(this.roots.appData, "hotkeys.json"), "Device hotkeys", (raw) =>
      this.stores.hotkeys.replaceAll(raw as Parameters<DeviceHotkeysStore["replaceAll"]>[0]),
    );
  }

  migrateAlertTemplates(): MigrationRow {
    return migrateJsonMapFile(join(this.roots.appData, "custom_alerts.json"), "Custom alert templates", (raw) =>
      this.stores.alerts.replaceAll(raw as Parameters<AlertTemplateStore["replaceAll"]>[0]),
    );
  }

  migrateHotkeyOptions(): MigrationRow {
    const label = "Hotkey options";
    const file = join(this.roots.appData, "hotkey_options.json");
    if (!existsSync(file)) return { source: label, status: "missing" };
    try {
      const raw = JSON.parse(readFileSync(file, "utf8")) as Record<string, unknown>;
      this.stores.hotkeyOptions.set({
        ParallelMode: raw["ParallelMode"] === true,
        ToggleDelayMs: typeof raw["ToggleDelayMs"] === "number" ? raw["ToggleDelayMs"] : undefined!,
      });
      return { source: label, status: "migrated" };
    } catch (err) {
      return { source: label, status: "failed", detail: err instanceof Error ? err.message : String(err) };
    }
  }

  migrateTrackedPlayers(): MigrationRow {
    const label = "Tracked players";
    const file = join(this.roots.appData, "tracked_players.json");
    if (!existsSync(file)) return { source: label, status: "missing" };
    try {
      const raw = JSON.parse(readFileSync(file, "utf8"));
      if (!Array.isArray(raw)) return { source: label, status: "failed", detail: "expected an array" };
      this.stores.trackedPlayers.replace(raw as Parameters<TrackedPlayersStore["replace"]>[0]);
      return { source: label, status: "migrated", detail: `${raw.length} player(s)` };
    } catch (err) {
      return { source: label, status: "failed", detail: err instanceof Error ? err.message : String(err) };
    }
  }

  migrateTutorialProgress(): MigrationRow {
    const label = "Tutorial progress";
    const file = join(this.roots.appData, "tutorial-progress.json");
    if (!existsSync(file)) return { source: label, status: "missing" };
    try {
      const raw = JSON.parse(readFileSync(file, "utf8")) as unknown;
      const doc = (typeof raw === "object" && raw !== null ? raw : {}) as {
        Tutorials?: Record<string, unknown>;
        Preferences?: Record<string, unknown>;
      };

      const warnings: string[] = [];
      const tutorials = doc["Tutorials"] ?? {};
      for (const [id, value] of Object.entries(tutorials)) {
        const check = tutorialProgressSchema.safeParse(value);
        if (!check.success) {
          warnings.push(`tutorial "${id}" invalid and skipped`);
          continue;
        }
        this.stores.tutorials.save(check.data);
      }
      if (doc["Preferences"]) {
        const prefs = tutorialPreferencesSchema.safeParse({
          ...TUTORIAL_PREFERENCES_DEFAULTS,
          ...doc["Preferences"],
        });
        if (prefs.success) this.stores.tutorials.savePreferences(prefs.data);
        else warnings.push("preferences invalid; defaults kept");
      }
      return {
        source: label,
        status: warnings.length ? "warning" : "migrated",
        detail: `${Object.keys(tutorials).length} tutorial(s)`,
        warnings: warnings.length ? warnings : undefined,
      };
    } catch (err) {
      return { source: label, status: "failed", detail: err instanceof Error ? err.message : String(err) };
    }
  }

  migrateFcmConfig(): MigrationRow {
    return copyThrough(this, "FCM pairing config", "rustplusjs-config.json");
  }

  migrateCacheFiles(): MigrationRow {
    const label = "Cache files";
    const cacheDir = join(this.roots.appData, "cache");
    if (!existsSync(cacheDir)) return { source: label, status: "missing" };

    const destCache = join(this.userDataDir, "cache");
    mkdirSync(destCache, { recursive: true });
    const copied: string[] = [];
    const skipped: string[] = [];
    for (const name of COPY_THROUGH_CACHE_FILES) {
      const src = join(cacheDir, name);
      if (!existsSync(src)) continue;
      copyFileSync(src, join(destCache, name));
      copied.push(name);
    }
    // Anything else in cache\ is stage-owned (map3d-parser-runtime etc.) — defer explicitly.
    for (const entry of readdirSync(cacheDir)) {
      if (!COPY_THROUGH_CACHE_FILES.includes(entry)) skipped.push(entry);
    }
    return {
      source: label,
      status: copied.length ? "copied" : "deferred",
      detail: copied.length ? `copied ${copied.join(", ")}` : "nothing to copy",
      warnings: skipped.length ? [`deferred: ${skipped.join(", ")}`] : undefined,
    };
  }
}

// --- source table ---------------------------------------------------------------

type SourceDef = {
  id: LegacySourceId;
  label: string;
  rel: string;
  kind: "file" | "dir";
  execute: (m: LegacyMigrator) => MigrationRow;
};

const SOURCE_DEFS: readonly SourceDef[] = [
  { id: "tracking_settings", label: "Tracking settings", rel: "tracking_settings.json", kind: "file", execute: (m) => m.migrateSettings() },
  { id: "profiles", label: "Server profiles", rel: "profiles.json", kind: "file", execute: (m) => m.migrateProfiles() },
  { id: "hotkeys", label: "Device hotkeys", rel: "hotkeys.json", kind: "file", execute: (m) => m.migrateHotkeys() },
  { id: "hotkey_options", label: "Hotkey options", rel: "hotkey_options.json", kind: "file", execute: (m) => m.migrateHotkeyOptions() },
  { id: "custom_alerts", label: "Custom alert templates", rel: "custom_alerts.json", kind: "file", execute: (m) => m.migrateAlertTemplates() },
  { id: "tracked_players", label: "Tracked players", rel: "tracked_players.json", kind: "file", execute: (m) => m.migrateTrackedPlayers() },
  { id: "tutorial_progress", label: "Tutorial progress", rel: "tutorial-progress.json", kind: "file", execute: (m) => m.migrateTutorialProgress() },
  { id: "fcm_config", label: "FCM pairing config", rel: "rustplusjs-config.json", kind: "file", execute: (m) => m.migrateFcmConfig() },
  { id: "cache_files", label: "Cache files", rel: "cache", kind: "dir", execute: (m) => m.migrateCacheFiles() },
];

// Deferred rows are appended so the report shows the full picture even when nothing else exists.
export function deferredRows(): MigrationRow[] {
  return [
    { source: "Map overlays (Overlays\\)", status: "deferred", detail: "stage 6 (maps)" },
    { source: "Generated 3D maps (3DMaps\\)", status: "deferred", detail: "stage 7 (maps 3D)" },
    { source: "Wipe tracker data (player-wipes\\)", status: "deferred", detail: "stage 9 (wipe tracker)" },
    { source: "Death logs (RustPlusDesktop\\deaths)", status: "deferred", detail: "stage 9 (death stats)" },
    { source: "Raid plan (LOCALAPPDATA raid-plan.json)", status: "deferred", detail: "stage 9 (raid calculator)" },
    { source: "WebView2 profiles (GeneticsLab SPA state)", status: "deferred", detail: "stage 8 (GeneticsLab)" },
  ];
}

// --- helpers ---------------------------------------------------------------------

function migrateJsonMapFile(file: string, label: string, apply: (raw: object) => void): MigrationRow {
  if (!existsSync(file)) return { source: label, status: "missing" };
  try {
    const raw = JSON.parse(readFileSync(file, "utf8")) as unknown;
    if (typeof raw !== "object" || raw === null || Array.isArray(raw)) {
      return { source: label, status: "failed", detail: "expected a JSON object" };
    }
    apply(raw);
    return { source: label, status: "migrated" };
  } catch (err) {
    return { source: label, status: "failed", detail: err instanceof Error ? err.message : String(err) };
  }
}

function copyThrough(m: LegacyMigrator, label: string, rel: string): MigrationRow {
  const src = join(m.roots.appData, rel);
  if (!existsSync(src)) return { source: label, status: "missing" };
  mkdirSync(m.userDataDir, { recursive: true });
  copyFileSync(src, join(m.userDataDir, rel));
  return { source: label, status: "copied" };
}
