import { z } from "zod";
import { defineChannel } from "./framework.js";
import { migrationRowSchema, type LegacyRootInfo, type LegacySourceInfo } from "../migration.js";

/** `app/getInfo` — renderer bootstrap snapshot (parity with the parts of App startup the UI needs first). */
export const appGetInfo = defineChannel(
  "app/getInfo",
  z.void(),
  z.object({
    version: z.string(),
    electron: z.string(),
    chrome: z.string(),
    node: z.string(),
    platform: z.string(),
    locale: z.string(),
    /** True when launched with RPD_SMOKE=1 (boot verification mode; window closes itself). */
    smokeMode: z.boolean(),
  }),
  "Static environment snapshot for renderer bootstrap.",
);

/** `app/logFromRenderer` — forwards renderer-side log lines into the main logger (single sink parity with AppendLog). */
export const logFromRenderer = defineChannel(
  "app/logFromRenderer",
  z.object({
    level: z.enum(["debug", "info", "warn", "error"]),
    scope: z.string().min(1).max(64),
    message: z.string().min(1).max(4000),
  }),
  z.void(),
  "Renderer log sink into the main-process rotating log.",
);

/** `uiPrefs/*` — persisted shell preferences (sidebar state). Backed by ui-prefs.json via JsonStore. */
export const uiPrefsSchema = z.object({
  sidebarPinned: z.boolean(),
  sidebarWidth: z.number().int().min(360).max(480),
});

export type UiPrefs = z.infer<typeof uiPrefsSchema>;

export const uiPrefsGet = defineChannel(
  "uiPrefs/get",
  z.void(),
  uiPrefsSchema,
  "Current persisted shell preferences.",
);

export const uiPrefsSet = defineChannel(
  "uiPrefs/set",
  z.object({
    sidebarPinned: z.boolean().optional(),
    sidebarWidth: z.number().int().min(360).max(480).optional(),
  }),
  uiPrefsSchema,
  "Patch shell preferences; returns the resulting full document.",
);

/** `migrate/*` — legacy-data migration (M3). Scan inventories C#-era sources; run imports them. */
export const migrateScan = defineChannel(
  "migrate/scan",
  z.void(),
  z.object({
    roots: z.array(
      z.object({
        kind: z.enum(["appData", "localData", "deathsDir"]),
        path: z.string(),
        exists: z.boolean(),
      }),
    ),
    sources: z.array(
      z.object({
        id: z.string(),
        label: z.string(),
        location: z.string(),
        exists: z.boolean(),
        bytes: z.number().nullable(),
      }),
    ),
  }),
  "Inventory legacy RustPlusDesk data sources without touching anything.",
);

export const migrateRun = defineChannel(
  "migrate/run",
  z.void(),
  z.object({
    startedAt: z.string(),
    finishedAt: z.string(),
    rows: z.array(migrationRowSchema),
  }),
  "Import all detected legacy sources into the new storage root; returns a per-source report.",
);

export type MigrateScanResult = { roots: LegacyRootInfo[]; sources: LegacySourceInfo[] };
export type MigrateRunResult = z.infer<typeof migrateRun["response"]>;

/** Registry consumed by preload (allow-list) and main (handler registration). Literal keys are mandatory:
 * they preserve per-channel def types (computed keys would collapse this to an index signature). */
export const ipcChannels = {
  "app/getInfo": appGetInfo,
  "app/logFromRenderer": logFromRenderer,
  "uiPrefs/get": uiPrefsGet,
  "uiPrefs/set": uiPrefsSet,
  "migrate/scan": migrateScan,
  "migrate/run": migrateRun,
};

export type IpcChannels = typeof ipcChannels;
export type IpcChannelName = keyof IpcChannels & string;

/** Compile-time guard: every registered key must equal its own def.name. */
const _nameParity: Readonly<{ [K in keyof IpcChannels]: IpcChannels[K]["name"] }> = {
  "app/getInfo": "app/getInfo",
  "app/logFromRenderer": "app/logFromRenderer",
  "uiPrefs/get": "uiPrefs/get",
  "uiPrefs/set": "uiPrefs/set",
  "migrate/scan": "migrate/scan",
  "migrate/run": "migrate/run",
};
void _nameParity;
