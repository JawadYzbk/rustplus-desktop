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

const cloudUserSchema = z.object({
  id: z.string(),
  steamId: z.string().nullable(),
  name: z.string().nullable(),
  displayName: z.string().nullable(),
  email: z.string().nullable(),
  providers: z.array(z.string()),
  hasPassword: z.boolean(),
});

const cloudCapabilitiesSchema = z.object({
  planCode: z.string(),
  isTrackerAvailable: z.boolean(),
  canTrackTeam: z.boolean(),
  canUseCloudSync: z.boolean(),
  canUseAdvancedViews: z.boolean(),
  canUseRouteReplay: z.boolean(),
  canExport: z.boolean(),
  maxTrackedPlayers: z.number().int().min(1),
  retainedWipes: z.number().int().min(1),
  cloudRetentionDays: z.number().int().min(0),
  fetchedAt: z.string(),
});

export const cloudLogin = defineChannel(
  "cloud/login",
  z.object({ email: z.string().email().max(320), password: z.string().min(1).max(1024) }),
  z.object({ signedIn: z.literal(true), user: cloudUserSchema }),
  "Sign in to the Laravel cloud account; the bearer token remains main-process only.",
);

export const cloudBootstrap = defineChannel(
  "cloud/bootstrap",
  z.object({}).strict(),
  z.object({
    signedIn: z.boolean(),
    user: cloudUserSchema.nullable(),
    capabilities: cloudCapabilitiesSchema.nullable(),
    error: z.string().nullable(),
  }),
  "Fetch Laravel client bootstrap and Player Wipe Tracker entitlements.",
);

export const cloudLogout = defineChannel(
  "cloud/logout",
  z.object({}).strict(),
  z.object({ signedIn: z.literal(false) }),
  "Clear the encrypted Laravel cloud session.",
);

const wipeVisitSchema = z.object({
  name: z.string(),
  startUtc: z.string(),
  endUtc: z.string(),
  entryX: z.number().nullable(),
  entryY: z.number().nullable(),
  exitX: z.number().nullable(),
  exitY: z.number().nullable(),
});

const wipeReplayPointSchema = z.object({
  timestampUtc: z.string(),
  x: z.number().nullable(),
  y: z.number().nullable(),
  state: z.enum(["moving", "stationary", "afk", "dead", "offline", "unknown"]),
  locationType: z.enum(["monument", "base", "open", "unknown"]),
  locationName: z.string().nullable(),
  grid: z.string().nullable(),
  event: z.enum(["death", "respawn"]).nullable(),
  sessionId: z.string(),
});

const wipeReplaySegmentSchema = z.object({
  startUtc: z.string(),
  endUtc: z.string(),
  state: z.enum(["moving", "stationary", "afk", "dead", "offline", "unknown"]),
});

const wipePlayerSchema = z.object({
  steamId: z.string().regex(/^\d{17}$/),
  name: z.string(),
  observationCount: z.number().int().min(0),
  summary: z.object({
    coverageSeconds: z.number(),
    unknownSeconds: z.number(),
    movingSeconds: z.number(),
    stationarySeconds: z.number(),
    afkSeconds: z.number(),
    deadSeconds: z.number(),
    offlineSeconds: z.number(),
    estimatedDistance: z.number(),
    deaths: z.number().int().min(0),
    monumentVisits: z.array(wipeVisitSchema),
  }),
  insights: z.object({
    firstSeenUtc: z.string().nullable(),
    lastSeenUtc: z.string().nullable(),
    sessionCount: z.number().int().min(0),
    topMonument: z.string().nullable(),
    topMonumentSeconds: z.number(),
    topMonumentVisits: z.number().int().min(0),
    longestBlindGapSeconds: z.number(),
    longestBlindGapStartUtc: z.string().nullable(),
    peakHourLocal: z.number().int().min(0).max(23).nullable(),
    peakHourActiveSeconds: z.number(),
    currentState: z.enum(["moving", "stationary", "afk", "dead", "offline", "unknown"]),
    currentLocationType: z.enum(["monument", "base", "open", "unknown"]),
    currentLocationName: z.string().nullable(),
    currentGrid: z.string().nullable(),
    currentAsOfUtc: z.string().nullable(),
    isLikelyOnline: z.boolean(),
  }),
  observations: z.array(wipeReplayPointSchema),
  segments: z.array(wipeReplaySegmentSchema),
});

export const wipeGetStatus = defineChannel(
  "wipe/getStatus",
  z.object({}).strict(),
  z.object({ serverKey: z.string().nullable(), wipeKey: z.string().nullable(), sessionId: z.string().nullable(), players: z.array(wipePlayerSchema) }),
  "Current local Player Wipe Tracker session and per-player summaries.",
);

export const wipeGetPlayer = defineChannel(
  "wipe/getPlayer",
  z.object({ steamId: z.string().regex(/^\d{17}$/) }),
  z.object({ player: wipePlayerSchema.nullable() }),
  "One local Player Wipe Tracker player summary and derived insights.",
);

export const wipeGetMap = defineChannel(
  "wipe/getMap",
  z.object({}).strict(),
  z.object({ map: z.object({ pngBase64: z.string(), imageWidth: z.number().int().positive(), imageHeight: z.number().int().positive(), worldSize: z.number(), worldRectX: z.number(), worldRectY: z.number(), worldRectWidth: z.number(), worldRectHeight: z.number() }).nullable() }),
  "Current local Wipe Tracker map image and world projection metadata.",
);

export const settingsGetWipe = defineChannel(
  "settings/getWipe",
  z.object({}).strict(),
  z.object({ enabled: z.boolean(), cloudBackupEnabled: z.boolean() }),
  "Player Wipe Tracker settings flags.",
);

export const settingsSetWipe = defineChannel(
  "settings/setWipe",
  z.object({ enabled: z.boolean().optional(), cloudBackupEnabled: z.boolean().optional() }),
  z.object({ enabled: z.boolean(), cloudBackupEnabled: z.boolean() }),
  "Persist Player Wipe Tracker settings flags.",
);

const wipeArchiveSchema = z.object({
  id: z.string().min(1),
  serverKey: z.string(),
  serverName: z.string(),
  wipeKey: z.string(),
  wipeStartedAtUtc: z.string().nullable(),
  firstObservedAtUtc: z.string().nullable(),
  lastObservedAtUtc: z.string().nullable(),
  playerCount: z.number().int().nullable(),
  storedBytes: z.number().int().nullable(),
  players: z.array(z.object({ steamId: z.string().regex(/^\d{17}$/), dayCount: z.number().int().min(0) })),
});

export const wipeGetCloudArchives = defineChannel(
  "wipe/getCloudArchives",
  z.object({}).strict(),
  z.object({ archives: z.array(wipeArchiveSchema) }),
  "List Laravel Player Wipe Tracker archives.",
);

export const wipeRestoreCloudArchive = defineChannel(
  "wipe/restoreCloudArchive",
  z.object({ archiveId: z.string().min(1).max(120) }),
  z.object({ archiveId: z.string(), players: z.number().int().min(0), days: z.number().int().min(0), observations: z.number().int().min(0), isCurrentWipe: z.boolean() }),
  "Restore a Laravel Player Wipe Tracker archive into local JSONL history.",
);

export const wipeDeleteCloudArchive = defineChannel(
  "wipe/deleteCloudArchive",
  z.object({ archiveId: z.string().min(1).max(120) }),
  z.object({ deleted: z.boolean() }),
  "Delete one Laravel Player Wipe Tracker archive.",
);

export const wipeDeleteAllCloud = defineChannel(
  "wipe/deleteAllCloud",
  z.object({}).strict(),
  z.object({ deleted: z.number().int().min(0) }),
  "Delete all Laravel Player Wipe Tracker archives.",
);

const deathSummarySchema = z.object({
  total: z.number().int().min(0),
  victims: z.number().int().min(0),
  avgSurvival: z.string(),
  longestSurvival: z.string(),
  peakHour: z.string(),
  deadliestPlace: z.string(),
  deadliestGrid: z.string(),
  byArea: z.array(z.object({ name: z.string(), type: z.enum(["monument", "base", "open"]), deaths: z.number().int().min(0), percent: z.number().int().min(0).max(100) })),
  byVictim: z.array(z.object({ victim: z.string(), deaths: z.number().int().min(0), avgSurvival: z.string() })),
  byLocation: z.array(z.object({ location: z.string(), type: z.enum(["monument", "base", "open"]), deaths: z.number().int().min(0) })),
  recent: z.array(z.object({ victim: z.string(), type: z.enum(["monument", "base", "open"]), location: z.string(), grid: z.string(), died: z.string() })),
  deathsPerDay: z.array(z.object({ day: z.string(), count: z.number().int().min(0) })),
});

export const deathsGetStats = defineChannel(
  "deaths/getStats",
  z.object({ search: z.string().max(200).optional(), player: z.string().max(200).optional(), type: z.enum(["all", "monument", "base", "open"]).optional(), range: z.enum(["all", "24h", "7d"]).optional() }).strict(),
  z.object({ serverKey: z.string().nullable(), players: z.array(z.string()), summary: deathSummarySchema }),
  "Read local JSONL death statistics for the active server with legacy filters.",
);

export const deathsClear = defineChannel(
  "deaths/clear",
  z.object({}).strict(),
  z.object({ cleared: z.boolean() }),
  "Clear the local death log for the active server.",
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

export const profilePair = defineChannel(
  "profile/pair",
  z.object({ link: z.string().min(1).max(4096), name: z.string().max(200).optional() }).strict(),
  z.object({
    activated: z.boolean(),
    profile: z.object({ matchKey: z.string(), name: z.string(), host: z.string(), port: z.number().int(), steamId64: z.string(), deviceCount: z.number().int().min(0) }),
  }),
  "Create or replace a server profile from a Rust+ pairing link and make it active.",
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
    pairedX: z.number().nullable(),
    pairedY: z.number().nullable(),
    pairedBySteamId: z.string().nullable(),
    pairedLocationCapturedAtMs: z.number().nullable(),
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
  pairedX: number | null;
  pairedY: number | null;
  pairedBySteamId: string | null;
  pairedLocationCapturedAtMs: number | null;
}

interface ExportedDeviceDto {
  entityId: number;
  kind: string | null;
  name: string | null;
  alias: string | null;
  isGroup: boolean;
  children: ExportedDeviceDto[] | null;
  customIconId: number | null;
  customIconShortName: string | null;
  inGameAlarmTitle: string | null;
  oilRigTrigger: string | null;
}

const exportedDeviceDtoSchema: z.ZodType<ExportedDeviceDto> = z.lazy(() =>
  z.object({
    entityId: z.number().int().min(0),
    kind: z.string().nullable(),
    name: z.string().nullable(),
    alias: z.string().nullable(),
    isGroup: z.boolean(),
    children: z.array(exportedDeviceDtoSchema).nullable(),
    customIconId: z.number().int().nullable(),
    customIconShortName: z.string().nullable(),
    inGameAlarmTitle: z.string().nullable(),
    oilRigTrigger: z.string().nullable(),
  }),
);

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

export const profileExportDevices = defineChannel(
  "profile/exportDevices",
  z.object({ matchKey: z.string().min(1) }),
  z.object({ saved: z.boolean(), canceled: z.boolean(), path: z.string().nullable(), bytes: z.number().int().min(0) }),
  "Save the selected profile's legacy-compatible device snapshot as JSON.",
);

export const profileImportPreview = defineChannel(
  "profile/importPreview",
  z.object({ matchKey: z.string().min(1) }),
  z.object({
    canceled: z.boolean(),
    path: z.string().nullable(),
    candidates: z.array(
      z.object({
        id: z.string().min(1),
        ownerSteamId: z.string(),
        ownerName: z.string(),
        entityId: z.number().int().min(0),
        kind: z.string().nullable(),
        name: z.string().nullable(),
        alias: z.string().nullable(),
        alreadyPresent: z.boolean(),
        fromPreviousWipe: z.boolean(),
        serverName: z.string(),
        existsState: z.enum(["?", "ok", "missing", "err", "local"]),
        originalDto: exportedDeviceDtoSchema,
      }),
    ),
  }),
  "Open a device snapshot and return selectable import candidates without changing the profile.",
);

export const profileApplyImport = defineChannel(
  "profile/applyImport",
  z.object({ matchKey: z.string().min(1), devices: z.array(exportedDeviceDtoSchema).max(1_000) }),
  z.object({ saved: z.boolean(), imported: z.number().int().min(0) }),
  "Add selected device snapshot entries to a profile, skipping existing entity IDs.",
);

export const profileDeleteDevice = defineChannel(
  "profile/deleteDevice",
  z.object({ matchKey: z.string().min(1), entityId: z.number().int().min(0) }),
  z.object({ removed: z.boolean(), reason: z.enum(["removed", "notFound", "notMissing"]) }),
  "Remove one missing leaf device from a profile; groups and live devices are protected.",
);

/** `deviceAutomation/*` — profile-scoped switch automation (stage 5). */
const deviceAutomationRuleSchema = z.object({
  id: z.string().min(1),
  name: z.string().min(1).max(120),
  isEnabled: z.boolean(),
  isExpanded: z.boolean(),
  conditionType: z.enum(["PlayerProximity", "GameTime"]),
  playerMatchMode: z.enum(["AnyOnline", "AllOnline", "Specific", "SpecificOffline", "AnyOffline", "AllOffline"]),
  specificPlayerSteamId: z.string().max(32),
  locationEntityId: z.number().int().min(0),
  distanceMeters: z.number().finite().min(1).max(1_000_000),
  startTime: z.string().max(32),
  endTime: z.string().max(32),
  targetEntityId: z.number().int().min(0),
  matchedState: z.boolean(),
  unmatchedState: z.boolean(),
});

export const deviceAutomationGetRules = defineChannel(
  "deviceAutomation/getRules",
  z.object({ matchKey: z.string().min(1) }),
  z.object({ found: z.boolean(), isActive: z.boolean(), rules: z.array(deviceAutomationRuleSchema) }),
  "Load one profile's DeviceAutomationRules and master switch.",
);

export const deviceAutomationSaveRules = defineChannel(
  "deviceAutomation/saveRules",
  z.object({
    matchKey: z.string().min(1),
    isActive: z.boolean(),
    rules: z.array(deviceAutomationRuleSchema).max(100),
  }),
  z.object({ saved: z.boolean() }),
  "Persist one profile's DeviceAutomationRules and master switch.",
);

/** `raid/*` — embedded raid dataset and calculator engine (stage 9). */
const raidResourceSchema = z.object({
  shortname: z.string(),
  itemId: z.number().int(),
  displayName: z.string(),
  amount: z.number().finite(),
});

const raidSourceSchema = z.object({
  sourceId: z.number().int(),
  prefabName: z.string(),
  itemId: z.number().int().nullable(),
  itemShortname: z.string(),
  itemSlug: z.string(),
  itemCategorySlug: z.string(),
  displayName: z.string(),
  kind: z.string(),
  rawDamage: z.number().finite(),
  craftCost: z.array(raidResourceSchema).nullable(),
  workbenchLevelRequired: z.number().int().nullable(),
});

const raidTargetSchema = z.object({
  targetId: z.number().int(),
  prefabName: z.string(),
  itemId: z.number().int().nullable(),
  itemShortname: z.string().nullable(),
  itemSlug: z.string().nullable(),
  itemCategorySlug: z.string().nullable(),
  buildingSlug: z.string().nullable(),
  buildingImage: z.string().nullable(),
  displayName: z.string(),
  buildingTier: z.string().nullable(),
  componentType: z.string(),
  startHealth: z.number().finite().positive(),
  category: z.string(),
});

const raidMethodSchema = z.object({
  source: raidSourceSchema,
  requiredItems: z.number().int().min(1),
  damagePerItem: z.number().finite(),
  totalDamage: z.number().finite(),
  overkill: z.number().finite().min(0),
  resources: z.array(raidResourceSchema),
  hasCraftCost: z.boolean(),
});

export const raidGetData = defineChannel(
  "raid/getData",
  z.void(),
  z.object({ sources: z.array(raidSourceSchema), targets: z.array(raidTargetSchema) }),
  "Load the validated embedded raid dataset for the calculator UI.",
);

export const raidCalculate = defineChannel(
  "raid/calculate",
  z.object({
    targetId: z.number().int().positive(),
    targetQuantity: z.number().int().min(1).max(100_000),
    sourceIds: z.array(z.number().int().positive()).max(100),
    mode: z.enum(["LowestSulfur", "LowestTotalResources", "FewestRaidItems", "Custom"]),
  }),
  z.object({
    methods: z.array(raidMethodSchema),
    recommended: raidMethodSchema.nullable(),
    combination: z.array(raidMethodSchema),
    resources: z.array(raidResourceSchema),
    items: z.array(z.object({ source: raidSourceSchema, amount: z.number().int().min(1) })),
  }),
  "Calculate raid methods and the selected best combination for one target.",
);

/** `recycler/*` — embedded recycler inputs and wild/safe-zone yield calculation (stage 9). */
const recyclerItemSchema = z.object({
  id: z.string(),
  shortName: z.string().min(1),
  displayName: z.string().min(1),
  category: z.string().min(1),
  stackSize: z.number().int().min(1),
});

const recyclerMetricSchema = z.object({
  expected: z.number().finite().min(0),
  guaranteed: z.number().finite().min(0),
  chance: z.number().finite().min(0),
  chancePercent: z.number().finite().min(0).max(100),
  min: z.number().finite().min(0),
  max: z.number().finite().min(0),
});

const recyclerOutputSchema = z.object({
  shortName: z.string().min(1),
  displayName: z.string().min(1),
  wild: recyclerMetricSchema,
  safe: recyclerMetricSchema,
});

export const recyclerGetData = defineChannel(
  "recycler/getData",
  z.void(),
  z.object({ items: z.array(recyclerItemSchema) }),
  "Load the validated embedded recycler input catalog.",
);

export const recyclerCalculate = defineChannel(
  "recycler/calculate",
  z.object({
    quantities: z.array(z.object({ shortName: z.string().min(1), quantity: z.number().int().min(0).max(2_000_000_000) })).max(1_000),
  }),
  z.object({
    outputs: z.array(recyclerOutputSchema),
    wildSeconds: z.number().finite().min(0),
    safeSeconds: z.number().finite().min(0),
  }),
  "Calculate wild and safe-zone recycler yields and processing times.",
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

/** Full rule incl. steps — the step editor reads/writes one rule at a time. Step fields are
 * optional in the bridge: parseLogicRule on the main side applies C# defaults, clamps and
 * unknown-enum tolerance (Models/LogicRule.cs parity), so the renderer can send partial steps. */
const logicStepSchema: z.ZodType<Record<string, unknown>> = z.lazy(() =>
  z
    .object({
      stepType: z.enum(["Wait", "Toggle", "CheckAvailability", "StartTimer"]).optional(),
      timerMinutes: z.number().optional(),
      timerTarget: z.enum(["Custom", "SmallOilRig", "LargeOilRig"]).optional(),
      timerName: z.string().max(120).optional(),
      showCrateOnMap: z.boolean().optional(),
      alarmTextHint: z.string().max(400).optional(),
      waitSeconds: z.number().min(0).optional(),
      targetEntityId: z.number().int().min(0).optional(),
      targetGroupName: z.string().max(120).optional(),
      toggleState: z.boolean().nullable().optional(),
      conditionOperator:
        z.enum(["IS_OFFLINE", "IS_ONLINE", "ALL_OFFLINE", "ANY_OFFLINE", "ALL_ONLINE", "ANY_ONLINE"]).optional(),
      conditionDeviceIdsCsv: z.string().max(2000).optional(),
      conditionalSteps: z.array(logicStepSchema).max(20).optional(),
    })
    .passthrough(),
);

export const logicGetRule = defineChannel(
  "logic/getRule",
  z.object({ matchKey: z.string().min(1), ruleId: z.string().min(1) }),
  z.object({
    found: z.boolean(),
    rule: z
      .object({
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
        steps: z.array(logicStepSchema),
      })
      .nullable(),
  }),
  "One full rule with steps for the step editor.",
);

export const logicSaveRule = defineChannel(
  "logic/saveRule",
  z.object({
    matchKey: z.string().min(1),
    rule: z.object({
      id: z.string().min(1),
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
      steps: z.array(logicStepSchema).max(50),
    }),
  }),
  z.object({ saved: z.boolean() }),
  "Replace one rule wholesale (header + steps); unknown ids are appended.",
);

export const logicGetTimers = defineChannel(
  "logic/getTimers",
  z.object({ matchKey: z.string().min(1) }),
  z.object({
    timers: z.array(
      z.object({
        id: z.string(),
        name: z.string(),
        command: z.string(),
        endTimeUtcMs: z.number(),
        enableCountdownAudio: z.boolean(),
        enableAlarmAudio: z.boolean(),
      }),
    ),
  }),
  "Custom timers for one profile (timers panel).",
);

export const logicAddTimer = defineChannel(
  "logic/addTimer",
  z.object({
    matchKey: z.string().min(1),
    name: z.string().max(120),
    hours: z.number().int().min(0).max(720),
    minutes: z.number().int().min(0).max(59),
    seconds: z.number().int().min(0).max(59),
  }),
  z.object({ ok: z.boolean(), id: z.string(), reason: z.enum(["limit", "letter", "duration"]).nullable() }),
  "BtnAddTimer_Click parity — five-limit / letter rule / duration-required validation.",
);

export const logicRemoveTimer = defineChannel(
  "logic/removeTimer",
  z.object({ matchKey: z.string().min(1), id: z.string() }),
  z.object({ removed: z.boolean() }),
  "Delete a custom timer by id.",
);

/** One-way main→renderer event stream (NOT an invoke channel — no request/response schema).
 * Payload: { stream: "conn" | "poll" | "device", event: <ConnRuntime event> }. */
export const pushChannel = "conn/push";

/** Registry consumed by preload (allow-list) and main (handler registration). Literal keys are mandatory:
 * they preserve per-channel def types (computed keys would collapse this to an index signature). */
export const ipcChannels = {
  "app/getInfo": appGetInfo,
  "app/logFromRenderer": logFromRenderer,
  "cloud/login": cloudLogin,
  "cloud/bootstrap": cloudBootstrap,
  "cloud/logout": cloudLogout,
  "wipe/getStatus": wipeGetStatus,
  "wipe/getPlayer": wipeGetPlayer,
  "wipe/getMap": wipeGetMap,
  "settings/getWipe": settingsGetWipe,
  "settings/setWipe": settingsSetWipe,
  "wipe/getCloudArchives": wipeGetCloudArchives,
  "wipe/restoreCloudArchive": wipeRestoreCloudArchive,
  "wipe/deleteCloudArchive": wipeDeleteCloudArchive,
  "wipe/deleteAllCloud": wipeDeleteAllCloud,
  "deaths/getStats": deathsGetStats,
  "deaths/clear": deathsClear,
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
  "profile/pair": profilePair,
  "profile/getDevices": profileGetDevices,
  "profile/saveDevices": profileSaveDevices,
  "profile/activate": profileActivate,
  "profile/exportDevices": profileExportDevices,
  "profile/importPreview": profileImportPreview,
  "profile/applyImport": profileApplyImport,
  "profile/deleteDevice": profileDeleteDevice,
  "deviceAutomation/getRules": deviceAutomationGetRules,
  "deviceAutomation/saveRules": deviceAutomationSaveRules,
  "raid/getData": raidGetData,
  "raid/calculate": raidCalculate,
  "recycler/getData": recyclerGetData,
  "recycler/calculate": recyclerCalculate,
  "logic/status": logicStatus,
  "logic/stop": logicStop,
  "logic/run": logicRun,
  "logic/getRules": logicGetRules,
  "logic/saveRules": logicSaveRules,
  "logic/getRule": logicGetRule,
  "logic/saveRule": logicSaveRule,
  "logic/getTimers": logicGetTimers,
  "logic/addTimer": logicAddTimer,
  "logic/removeTimer": logicRemoveTimer,
};

export type IpcChannels = typeof ipcChannels;
export type IpcChannelName = keyof IpcChannels & string;

/** Compile-time guard: every registered key must equal its own def.name. */
const _nameParity: Readonly<{ [K in keyof IpcChannels]: IpcChannels[K]["name"] }> = {
  "app/getInfo": "app/getInfo",
  "app/logFromRenderer": "app/logFromRenderer",
  "cloud/login": "cloud/login",
  "cloud/bootstrap": "cloud/bootstrap",
  "cloud/logout": "cloud/logout",
  "wipe/getStatus": "wipe/getStatus",
  "wipe/getPlayer": "wipe/getPlayer",
  "wipe/getMap": "wipe/getMap",
  "settings/getWipe": "settings/getWipe",
  "settings/setWipe": "settings/setWipe",
  "wipe/getCloudArchives": "wipe/getCloudArchives",
  "wipe/restoreCloudArchive": "wipe/restoreCloudArchive",
  "wipe/deleteCloudArchive": "wipe/deleteCloudArchive",
  "wipe/deleteAllCloud": "wipe/deleteAllCloud",
  "deaths/getStats": "deaths/getStats",
  "deaths/clear": "deaths/clear",
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
  "profile/pair": "profile/pair",
  "profile/getDevices": "profile/getDevices",
  "profile/saveDevices": "profile/saveDevices",
  "profile/activate": "profile/activate",
  "profile/exportDevices": "profile/exportDevices",
  "profile/importPreview": "profile/importPreview",
  "profile/applyImport": "profile/applyImport",
  "profile/deleteDevice": "profile/deleteDevice",
  "deviceAutomation/getRules": "deviceAutomation/getRules",
  "deviceAutomation/saveRules": "deviceAutomation/saveRules",
  "raid/getData": "raid/getData",
  "raid/calculate": "raid/calculate",
  "recycler/getData": "recycler/getData",
  "recycler/calculate": "recycler/calculate",
  "logic/status": "logic/status",
  "logic/stop": "logic/stop",
  "logic/run": "logic/run",
  "logic/getRules": "logic/getRules",
  "logic/saveRules": "logic/saveRules",
  "logic/getRule": "logic/getRule",
  "logic/saveRule": "logic/saveRule",
  "logic/getTimers": "logic/getTimers",
  "logic/addTimer": "logic/addTimer",
  "logic/removeTimer": "logic/removeTimer",
};
void _nameParity;
