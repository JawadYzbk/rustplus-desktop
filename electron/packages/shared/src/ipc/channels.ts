import { z } from "zod";
import { defineChannel } from "./framework.js";
import { migrationRowSchema, type LegacyRootInfo, type LegacySourceInfo } from "../migration.js";
import { backupCreate, backupRestore, resetPerform } from "../backup.js";

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
export type BackupCreateResult = z.infer<typeof backupCreate["response"]>;
export type BackupRestoreResult = z.infer<typeof backupRestore["response"]>;
export type ResetPerformResult = z.infer<typeof resetPerform["response"]>;

/** Connection control — request payloads validated main-side against the live profile store. */
const connConnect = defineChannel(
  "conn/connect",
  z.object({
    host: z.string().min(1),
    port: z.number().int().min(1).max(65535),
    steamId64: z.string().regex(/^\d{17}$/),
    playerToken: z.string().min(1),
    useProxy: z.boolean().optional(),
  }),
  z.object({
    connected: z.boolean(),
    activeProxy: z.enum(["direct", "proxy"]).nullable(),
    host: z.string().nullable(),
    port: z.number().nullable(),
    consecutiveTimeouts: z.number(),
    teamChatPrimed: z.boolean(),
    clanChatPrimed: z.boolean(),
  }),
  "Open the Rust+ companion connection (dual-path direct/probe fallback).",
);

const connDisconnect = defineChannel(
  "conn/disconnect",
  z.object({}).strict(),
  z.object({
    connected: z.boolean(),
    activeProxy: z.enum(["direct", "proxy"]).nullable(),
    host: z.string().nullable(),
    port: z.number().nullable(),
    consecutiveTimeouts: z.number(),
    teamChatPrimed: z.boolean(),
    clanChatPrimed: z.boolean(),
  }),
  "Close the companion connection and cancel any pending silent reconnect.",
);

const connStatus = defineChannel(
  "conn/status",
  z.object({}).strict(),
  z.object({
    connected: z.boolean(),
    activeProxy: z.enum(["direct", "proxy"]).nullable(),
    host: z.string().nullable(),
    port: z.number().nullable(),
    consecutiveTimeouts: z.number(),
    teamChatPrimed: z.boolean(),
    clanChatPrimed: z.boolean(),
  }),
  "Current connection snapshot (no side effects).",
);

export type ConnSnapshotDto = z.infer<typeof connStatus["response"]>;

/** `profile/*` — server profiles + device trees (stage 5). Tokens never cross the bridge. */
export const profileList = defineChannel(
  "profile/list",
  z.void(),
  z.object({
    profiles: z.array(
      z.object({
        matchKey: z.string(),
        name: z.string(),
        host: z.string(),
        port: z.number().int(),
        steamId64: z.string(),
        deviceCount: z.number().int().min(0),
      }),
    ),
  }),
  "All stored server profiles with their stable match keys.",
);

/** Recursive device-node contract (SmartDevice subset the tree UI needs). */
const deviceNodeSchema: z.ZodType<DeviceNodeDto> = z.lazy(() =>
  z.object({
    entityId: z.number().int().min(0),
    kind: z.string().nullable(),
    name: z.string().nullable(),
    alias: z.string().nullable(),
    isGroup: z.boolean(),
    children: z.array(deviceNodeSchema),
    isMissing: z.boolean(),
    customIconId: z.number().int().nullable(),
    customIconShortName: z.string().nullable(),
    inGameAlarmTitle: z.string().nullable(),
    oilRigTriggerTarget: z.string().nullable(),
  }),
);

export interface DeviceNodeDto {
  entityId: number;
  kind: string | null;
  name: string | null;
  alias: string | null;
  isGroup: boolean;
  children: DeviceNodeDto[];
  isMissing: boolean;
  customIconId: number | null;
  customIconShortName: string | null;
  inGameAlarmTitle: string | null;
  oilRigTriggerTarget: string | null;
}

export const profileGetDevices = defineChannel(
  "profile/getDevices",
  z.object({ matchKey: z.string().min(1) }),
  z.object({ devices: z.array(deviceNodeSchema), found: z.boolean() }),
  "Device tree of one profile (empty list when the profile does not exist).",
);

export const profileSaveDevices = defineChannel(
  "profile/saveDevices",
  z.object({ matchKey: z.string().min(1), devices: z.array(deviceNodeSchema) }),
  z.object({ saved: z.boolean() }),
  "Replace a profile's device tree (whole-tree write, matching legacy Save()).",
);

export const profileActivate = defineChannel(
  "profile/activate",
  z.object({ matchKey: z.string().min(1) }),
  z.object({ activated: z.boolean() }),
  "Select the active server profile (engine + connection context follow this).",
);

/** `logic/*` — Logic Engine control (stage 5). Rules live on the profile record. */
export const logicStatus = defineChannel(
  "logic/status",
  z.void(),
  z.object({
    activeKey: z.string().nullable(),
    isRunning: z.boolean(),
    currentRuleName: z.string().nullable(),
    currentStepNumber: z.number().int(),
    currentStepType: z.string().nullable(),
    pendingRules: z.array(z.string()),
  }),
  "Logic Engine runtime surface (LogicEngineRuntimeService parity).",
);

export const logicStop = defineChannel(
  "logic/stop",
  z.void(),
  z.object({ stopped: z.boolean() }),
  "Request cancellation; the current rule aborts after the in-flight operation.",
);

export const logicRun = defineChannel(
  "logic/run",
  z.object({ ruleId: z.string().min(1) }),
  z.object({ accepted: z.boolean() }),
  "Manually enqueue a rule by id (legacy 'run now').",
);

export const logicGetRules = defineChannel(
  "logic/getRules",
  z.object({ matchKey: z.string().min(1) }),
  z.object({
    found: z.boolean(),
    isEngineActive: z.boolean(),
    rules: z.array(
      z.object({
        id: z.string(),
        name: z.string(),
        isEnabled: z.boolean(),
        isLoopEnabled: z.boolean(),
        loopCount: z.number().int().min(0),
        triggerType: z.enum(["SmartAlarm", "SmartSwitch", "ChatCommand", "RuleTriggered", "RuleCompleted"]),
        triggerEntityId: z.number().int().min(0),
        triggerCommand: z.string(),
        triggerRuleId: z.string(),
        triggerState: z.boolean(),
        conditionOperator: z.enum(["NONE", "AND", "OR"]),
        conditionDeviceEntityId: z.number().int().min(0),
        conditionDeviceState: z.boolean(),
        stepCount: z.number().int().min(0),
      }),
    ),
  }),
  "Rule summaries for one profile (full step editing lands with the rule editor).",
);

export const logicSaveRules = defineChannel(
  "logic/saveRules",
  z.object({
    matchKey: z.string().min(1),
    isEngineActive: z.boolean(),
    rules: z.array(
      z.object({
        id: z.string(),
        name: z.string().min(1).max(120),
        isEnabled: z.boolean(),
        isLoopEnabled: z.boolean(),
        loopCount: z.number().int().min(0),
        triggerType: z.enum(["SmartAlarm", "SmartSwitch", "ChatCommand", "RuleTriggered", "RuleCompleted"]),
        triggerEntityId: z.number().int().min(0),
        triggerCommand: z.string().max(120),
        triggerRuleId: z.string().max(120),
        triggerState: z.boolean(),
        conditionOperator: z.enum(["NONE", "AND", "OR"]),
        conditionDeviceEntityId: z.number().int().min(0),
        conditionDeviceState: z.boolean(),
      }),
    ),
  }),
  z.object({ saved: z.boolean() }),
  "Persist rule headers + engine-active flag (steps are preserved untouched for unchanged ids).",
);

/** One-way main→renderer event stream (NOT an invoke channel — no request/response schema).
 * Payload: { stream: "conn" | "poll" | "device", event: <ConnRuntime event> }. */
export const pushChannel = "conn/push";

/** Registry consumed by preload (allow-list) and main (handler registration). Literal keys are mandatory:
 * they preserve per-channel def types (computed keys would collapse this to an index signature). */
export const ipcChannels = {
  "app/getInfo": appGetInfo,
  "app/logFromRenderer": logFromRenderer,
  "uiPrefs/get": uiPrefsGet,
  "uiPrefs/set": uiPrefsSet,
  "migrate/scan": migrateScan,
  "migrate/run": migrateRun,
  "backup/create": backupCreate,
  "backup/restore": backupRestore,
  "reset/perform": resetPerform,
  "conn/connect": connConnect,
  "conn/disconnect": connDisconnect,
  "conn/status": connStatus,
  "profile/list": profileList,
  "profile/getDevices": profileGetDevices,
  "profile/saveDevices": profileSaveDevices,
  "profile/activate": profileActivate,
  "logic/status": logicStatus,
  "logic/stop": logicStop,
  "logic/run": logicRun,
  "logic/getRules": logicGetRules,
  "logic/saveRules": logicSaveRules,
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
  "backup/create": "backup/create",
  "backup/restore": "backup/restore",
  "reset/perform": "reset/perform",
  "conn/connect": "conn/connect",
  "conn/disconnect": "conn/disconnect",
  "conn/status": "conn/status",
  "profile/list": "profile/list",
  "profile/getDevices": "profile/getDevices",
  "profile/saveDevices": "profile/saveDevices",
  "profile/activate": "profile/activate",
  "logic/status": "logic/status",
  "logic/stop": "logic/stop",
  "logic/run": "logic/run",
  "logic/getRules": "logic/getRules",
  "logic/saveRules": "logic/saveRules",
};
void _nameParity;
