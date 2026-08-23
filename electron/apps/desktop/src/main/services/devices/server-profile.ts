/**
 * ServerProfile persistence — port of Models/ServerProfile.cs (persisted surface) and
 * Services/Data/ProfileDataModule.cs.
 *
 * File contract: profiles.json is `JsonSerializer.Serialize(list, { WriteIndented:true })`
 * with DEFAULT options → PascalCase keys, EXCEPT DeathMarkers which carries an explicit
 * [JsonPropertyName("deathMarkers")]. Setters RUN during deserialization, so the load path
 * applies the same validation/clamping the UI setters do (ValidateCommand, prefix whitelist,
 * delay ranges) — reproduced here in normalize-on-parse.
 *
 * Unknown properties are captured and written back verbatim so files survive round-trips
 * across app versions in BOTH directions (a user can go back to the C# app).
 */
import { randomUUID } from "node:crypto";
import type { SmartDeviceNode } from "./device-data.js";

export interface ChatCommandMapping {
  label: string | null;
  command: string;
  entityId: number;
}

export interface CustomTimer {
  id: string;
  name: string;
  command: string;
  /** Epoch ms (C# DateTime EndTimeUtc serialized as ISO 8601). */
  endTimeUtcMs: number;
  enableCountdownAudio: boolean;
  enableAlarmAudio: boolean;
  createdNotified: boolean;
  notified60: boolean;
  notified30: boolean;
  notified10: boolean;
  notified3: boolean;
  countdownAudioPlayed: boolean;
  alarmPlayed: boolean;
  snoozedUntilUtcMs: number | null;
  autoDeleteAtUtcMs: number | null;
}

export interface ServerProfileData {
  host: string;
  port: number;
  steamId64: string;
  playerToken: string;
  battleMetricsId: string | null;
  localMapFilePath: string | null;
  localMapImagePath: string | null;
  customMapUrl: string | null;
  useFacepunchProxy: boolean;
  lastEventSource: string;
  devices: SmartDeviceNode[];
  cameraIds: string[];
  /** NOTE: legacy JSON key is camelCase "deathMarkers" ([JsonPropertyName]). */
  deathMarkers: unknown[];
  learnedDaySpeed: number; // default 12/50
  learnedNightSpeed: number; // default 12/10
  chatCommandsEnabled: boolean;
  cmdPop: string;
  cmdList: string;
  cmdTime: string;
  cmdPromote: string;
  cmdDeepSea: string;
  cmdCargo: string;
  cmdAfk: string;
  cmdOilRig: string;
  cmdHeli: string;
  cmdVendor: string;
  cmdUpkeepDetail: string;
  cmdCustomTimer: string;
  chatCommandPrefix: string; // only ! . , \ accepted
  chatCommandDelaySeconds: number; // 1..5
  chatResponseDelaySeconds: number; // 0..5
  switchCommandMappings: ChatCommandMapping[];
  upkeepCommandMappings: ChatCommandMapping[];
  alertCustomTimer: boolean;
  discordWebhookChatAlertsUrl: string;
  discordWebhookChatAlertsMention: string;
  discordWebhookChatAlertsEnabled: boolean;
  discordWebhookChatAlertsTts: boolean;
  discordWebhookChatAlertsExclusive: boolean;
  timerAlarmEnabled: boolean;
  timerAlarmAudioPath: string | null;
  timerCountdownAudioPath: string | null;
  timerAlarmSnoozeMinutes: number; // ≥0
  timerAlarmBeepDurationSeconds: number; // ≥1
  customTimers: CustomTimer[];
  rustMapsMapId: string | null;
  rustMapsFetchTimeMs: number | null;
  rustMapsWipeTimeMs: number | null;
  wipeTimeMs: number | null;
  logicRules: unknown[];
  isLogicEngineActive: boolean;
  deviceAutomationRules: unknown[];
  isDeviceAutomationActive: boolean;
  subscribedTeammateSteamIds: string[];
  /** Properties from other app versions, preserved verbatim across round-trips. */
  extra: Record<string, unknown>;
}

/** Stable identity used to remember the last selected server across restarts.
 *  PlayerToken deliberately excluded so a re-pair on the same server still matches. */
export function matchKey(p: Pick<ServerProfileData, "host" | "port" | "steamId64">): string {
  return `${p.host}:${p.port}|${p.steamId64}`;
}

// ------------------------------------------------------------------ validation

/** ValidateCommand parity: trim, strip leading '!', forbid leading digit → default. */
export function validateCommand(value: string | null | undefined, defaultValue: string): string {
  if (!value || value.trim().length === 0) return defaultValue;
  const trimmed = value.trim().replace(/^!+/, "");
  if (trimmed.length > 0 && /[0-9]/.test(trimmed[0]!)) return defaultValue;
  return trimmed;
}

const PREFIX_WHITELIST = new Set(["!", ".", ",", "\\"]);
const isoToMs = (v: unknown): number | null => {
  if (typeof v !== "string" || v.length === 0) return typeof v === "number" ? v : null;
  const t = Date.parse(v);
  return Number.isNaN(t) ? null : t;
};

function parseMappings(raw: unknown): ChatCommandMapping[] {
  if (!Array.isArray(raw)) return [];
  return raw.map((m) => {
    const r = (m ?? {}) as Record<string, unknown>;
    return {
      label: typeof r.Label === "string" ? r.Label : typeof r.label === "string" ? r.label : null,
      command:
        typeof r.Command === "string"
          ? r.Command
          : typeof r.command === "string"
            ? r.command
            : "",
      entityId:
        typeof r.EntityId === "number" ? r.EntityId : typeof r.entityId === "number" ? r.entityId : 0,
    };
  });
}

function serializeMappings(mappings: ChatCommandMapping[]): Record<string, unknown>[] {
  return mappings.map((m) => ({ Label: m.label, Command: m.command, EntityId: m.entityId }));
}

/** Exported for the LogicEngineService timer ticker (canonical ISO record format). */
export function parseTimer(raw: Record<string, unknown>): CustomTimer {
  return {
    id: typeof raw.Id === "string" ? raw.Id : typeof raw.id === "string" ? raw.id : randomUUID(),
    name: typeof raw.Name === "string" ? raw.Name : typeof raw.name === "string" ? raw.name : "Timer",
    command: typeof raw.Command === "string" ? raw.Command : typeof raw.command === "string" ? raw.command : "timer",
    endTimeUtcMs:
      isoToMs(raw.EndTimeUtc) ??
      isoToMs(raw.endTimeUtc) ??
      (typeof raw.EndTimeUtcMs === "number" ? raw.EndTimeUtcMs : null) ??
      Date.now(),
    // C# default is TRUE (ServerProfile.cs _enableCountdownAudio = true):
    enableCountdownAudio:
      (raw.EnableCountdownAudio ?? raw.enableCountdownAudio) !== false,
    enableAlarmAudio:
      (raw.EnableAlarmAudio ?? raw.enableAlarmAudio) === true,
    createdNotified: raw.CreatedNotified === true,
    notified60: raw.Notified60 === true,
    notified30: raw.Notified30 === true,
    notified10: raw.Notified10 === true,
    notified3: raw.Notified3 === true,
    countdownAudioPlayed: raw.CountdownAudioPlayed === true,
    alarmPlayed: raw.AlarmPlayed === true,
    snoozedUntilUtcMs: isoToMs(raw.SnoozedUntilUtc),
    autoDeleteAtUtcMs: isoToMs(raw.AutoDeleteAtUtc),
  };
}

/** Exported for the LogicEngineService timer ticker (canonical ISO record format). */
export function serializeTimer(t: CustomTimer): Record<string, unknown> {
  const iso = (ms: number | null): string | null => (ms === null ? null : new Date(ms).toISOString());
  return {
    Id: t.id,
    Name: t.name,
    Command: t.command,
    EndTimeUtc: iso(t.endTimeUtcMs),
    EnableCountdownAudio: t.enableCountdownAudio,
    EnableAlarmAudio: t.enableAlarmAudio,
    CreatedNotified: t.createdNotified,
    Notified60: t.notified60,
    Notified30: t.notified30,
    Notified10: t.notified10,
    Notified3: t.notified3,
    CountdownAudioPlayed: t.countdownAudioPlayed,
    AlarmPlayed: t.alarmPlayed,
    SnoozedUntilUtc: iso(t.snoozedUntilUtcMs),
    AutoDeleteAtUtc: iso(t.autoDeleteAtUtcMs),
  };
}

const KNOWN_KEYS = new Set<string>([
  "Host", "Port", "SteamId64", "PlayerToken", "BattleMetricsId",
  "LocalMapFilePath", "LocalMapImagePath", "CustomMapUrl",
  "UseFacepunchProxy", "LastEventSource", "Devices", "CameraIds", "deathMarkers",
  "LearnedDaySpeed", "LearnedNightSpeed", "ChatCommandsEnabled",
  "CmdPop", "CmdList", "CmdTime", "CmdPromote", "CmdDeepSea", "CmdCargo", "CmdAfk",
  "CmdOilRig", "CmdHeli", "CmdVendor", "CmdUpkeepDetail", "CmdCustomTimer",
  "ChatCommandPrefix", "ChatCommandDelaySeconds", "ChatResponseDelaySeconds",
  "SwitchCommandMappings", "UpkeepCommandMappings", "AlertCustomTimer",
  "DiscordWebhookChatAlertsUrl", "DiscordWebhookChatAlertsMention",
  "DiscordWebhookChatAlertsEnabled", "DiscordWebhookChatAlertsTts", "DiscordWebhookChatAlertsExclusive",
  "TimerAlarmEnabled", "TimerAlarmAudioPath", "TimerCountdownAudioPath",
  "TimerAlarmSnoozeMinutes", "TimerAlarmBeepDurationSeconds", "CustomTimers",
  "RustMapsMapId", "RustMapsFetchTime", "RustMapsWipeTime", "WipeTime",
  "LogicRules", "IsLogicEngineActive", "DeviceAutomationRules", "IsDeviceAutomationActive",
  "SubscribedTeammateSteamIds",
  // Runtime-only state the C# serializer also wrote but which never means anything on load:
  "IsConnected", "IsFullConnected",
]);

/**
 * Parse one profile object with load-path validation parity (setters run during
 * deserialization in C#, so invalid values collapse to defaults exactly like this).
 */
// eslint-disable-next-line complexity
export function parseServerProfile(raw: unknown): ServerProfileData {
  const r = (raw ?? {}) as Record<string, unknown>;
  const strOf = (k: string, dflt = ""): string => (typeof r[k] === "string" ? (r[k] as string) : dflt);
  const boolOf = (k: string, dflt = false): boolean => (typeof r[k] === "boolean" ? (r[k] as boolean) : dflt);
  const numOr = (k: string, dflt: number): number =>
    typeof r[k] === "number" && Number.isFinite(r[k] as number) ? (r[k] as number) : dflt;

  const prefixRaw = typeof r.ChatCommandPrefix === "string" ? (r.ChatCommandPrefix as string) : "";
  const devices = Array.isArray(r.Devices) ? (r.Devices as unknown[]) : [];

  return {
    host: strOf("Host"),
    port: numOr("Port", 28082),
    steamId64: strOf("SteamId64"),
    playerToken: strOf("PlayerToken"),
    battleMetricsId: typeof r.BattleMetricsId === "string" ? (r.BattleMetricsId as string) : null,
    localMapFilePath: typeof r.LocalMapFilePath === "string" ? (r.LocalMapFilePath as string) : null,
    localMapImagePath: typeof r.LocalMapImagePath === "string" ? (r.LocalMapImagePath as string) : null,
    customMapUrl: typeof r.CustomMapUrl === "string" ? (r.CustomMapUrl as string) : null,
    useFacepunchProxy: boolOf("UseFacepunchProxy"),
    lastEventSource: strOf("LastEventSource"),
    devices: parseDevices(r.Devices),
    cameraIds: Array.isArray(r.CameraIds) ? (r.CameraIds as unknown[]).filter((x): x is string => typeof x === "string") : [],
    deathMarkers: Array.isArray(r.deathMarkers) ? r.deathMarkers : [],
    learnedDaySpeed: numOr("LearnedDaySpeed", 12 / 50),
    learnedNightSpeed: numOr("LearnedNightSpeed", 12 / 10),
    chatCommandsEnabled: boolOf("ChatCommandsEnabled"),
    cmdPop: validateCommand(strOf("CmdPop"), "pop"),
    cmdList: validateCommand(strOf("CmdList"), "commands"),
    cmdTime: validateCommand(strOf("CmdTime"), "time"),
    cmdPromote: validateCommand(strOf("CmdPromote"), "promote"),
    cmdDeepSea: validateCommand(strOf("CmdDeepSea"), "deepsea"),
    cmdCargo: validateCommand(strOf("CmdCargo"), "cargo"),
    cmdAfk: validateCommand(strOf("CmdAfk"), "afk"),
    cmdOilRig: validateCommand(strOf("CmdOilRig"), "oilrig"),
    cmdHeli: validateCommand(strOf("CmdHeli"), "heli"),
    cmdVendor: validateCommand(strOf("CmdVendor"), "vendor"),
    cmdUpkeepDetail: validateCommand(strOf("CmdUpkeepDetail"), "upkeepdetail"),
    cmdCustomTimer: validateCommand(strOf("CmdCustomTimer"), "timer"),
    chatCommandPrefix: PREFIX_WHITELIST.has(prefixRaw) && prefixRaw.length > 0 ? prefixRaw : "!",
    chatCommandDelaySeconds: (() => {
      const v = numOr("ChatCommandDelaySeconds", 2);
      return v >= 1 && v <= 5 ? v : 2;
    })(),
    chatResponseDelaySeconds: (() => {
      const v = numOr("ChatResponseDelaySeconds", 0.5);
      return v >= 0 && v <= 5 ? v : 0.5;
    })(),
    switchCommandMappings: parseMappings(r.SwitchCommandMappings),
    upkeepCommandMappings: parseMappings(r.UpkeepCommandMappings),
    alertCustomTimer: boolOf("AlertCustomTimer", true),
    discordWebhookChatAlertsUrl: strOf("DiscordWebhookChatAlertsUrl"),
    discordWebhookChatAlertsMention: strOf("DiscordWebhookChatAlertsMention"),
    discordWebhookChatAlertsEnabled: boolOf("DiscordWebhookChatAlertsEnabled"),
    discordWebhookChatAlertsTts: boolOf("DiscordWebhookChatAlertsTts"),
    discordWebhookChatAlertsExclusive: boolOf("DiscordWebhookChatAlertsExclusive"),
    timerAlarmEnabled: boolOf("TimerAlarmEnabled", true),
    timerAlarmAudioPath: typeof r.TimerAlarmAudioPath === "string" ? (r.TimerAlarmAudioPath as string) : null,
    timerCountdownAudioPath: typeof r.TimerCountdownAudioPath === "string" ? (r.TimerCountdownAudioPath as string) : null,
    timerAlarmSnoozeMinutes: Math.max(0, numOr("TimerAlarmSnoozeMinutes", 5)),
    timerAlarmBeepDurationSeconds: Math.max(1, numOr("TimerAlarmBeepDurationSeconds", 5)),
    customTimers: Array.isArray(r.CustomTimers)
      ? (r.CustomTimers as unknown[]).map((t) => parseTimer((t ?? {}) as Record<string, unknown>))
      : [],
    rustMapsMapId: typeof r.RustMapsMapId === "string" ? (r.RustMapsMapId as string) : null,
    rustMapsFetchTimeMs: isoToMs(r.RustMapsFetchTime),
    rustMapsWipeTimeMs: isoToMs(r.RustMapsWipeTime),
    wipeTimeMs: isoToMs(r.WipeTime),
    logicRules: Array.isArray(r.LogicRules) ? r.LogicRules : [],
    isLogicEngineActive: boolOf("IsLogicEngineActive"),
    deviceAutomationRules: Array.isArray(r.DeviceAutomationRules) ? r.DeviceAutomationRules : [],
    isDeviceAutomationActive: boolOf("IsDeviceAutomationActive"),
    subscribedTeammateSteamIds: Array.isArray(r.SubscribedTeammateSteamIds)
      ? (r.SubscribedTeammateSteamIds as unknown[]).map((id) => String(id))
      : [],
    extra: Object.fromEntries(Object.entries(r).filter(([k]) => !KNOWN_KEYS.has(k))),
  };
}

// Devices inside profiles.json are full SmartDevice objects (not DTOs); map known fields with
// dual-casing tolerance, children recursive.
export function parseDevices(raw: unknown): SmartDeviceNode[] {
  if (!Array.isArray(raw)) return [];
  const walk = (d: unknown): SmartDeviceNode => {
    const r = (d ?? {}) as Record<string, unknown>;
    const childrenRaw = Array.isArray(r.Children) ? r.Children : Array.isArray(r.children) ? r.children : [];
    return {
      entityId: typeof r.EntityId === "number" ? r.EntityId : typeof r.entityId === "number" ? r.entityId : 0,
      kind: typeof r.Kind === "string" ? r.Kind : typeof r.kind === "string" ? r.kind : null,
      name: typeof r.Name === "string" ? r.Name : typeof r.name === "string" ? r.name : null,
      alias: typeof r.Alias === "string" ? r.Alias : typeof r.alias === "string" ? r.alias : null,
      isGroup: r.IsGroup === true || r.isGroup === true,
      children: childrenRaw.map(walk),
      isMissing: r.IsMissing === true || r.isMissing === true,
      customIconId: typeof r.CustomIconId === "number" ? r.CustomIconId : null,
      customIconShortName:
        typeof r.CustomIconShortName === "string" ? r.CustomIconShortName : typeof r.customIconShortName === "string" ? r.customIconShortName : null,
      inGameAlarmTitle:
        typeof r.InGameAlarmTitle === "string" ? r.InGameAlarmTitle : typeof r.inGameAlarmTitle === "string" ? r.inGameAlarmTitle : null,
      oilRigTriggerTarget:
        typeof r.OilRigTriggerTarget === "string" ? r.OilRigTriggerTarget : typeof r.oilRigTriggerTarget === "string" ? r.oilRigTriggerTarget : null,
    };
  };
  return raw.map((d: unknown): SmartDeviceNode => walk(d));
}

export function newEmptyServerProfile(): ServerProfileData {
  const p = parseServerProfile({});
  p.devices = [];
  return p;
}

/** PascalCase serialization matching System.Text.Json defaults, incl. the deathMarkers exception. */
export function serializeServerProfile(p: ServerProfileData): Record<string, unknown> {
  const iso = (ms: number | null): string | null => (ms === null ? null : new Date(ms).toISOString());
  return {
    Host: p.host,
    Port: p.port,
    SteamId64: p.steamId64,
    PlayerToken: p.playerToken,
    BattleMetricsId: p.battleMetricsId,
    LocalMapFilePath: p.localMapFilePath,
    LocalMapImagePath: p.localMapImagePath,
    CustomMapUrl: p.customMapUrl,
    IsConnected: false, // runtime state; legacy wrote it, always false at rest
    IsFullConnected: false,
    UseFacepunchProxy: p.useFacepunchProxy,
    LastEventSource: p.lastEventSource,
    Devices: serializeDevices(p.devices),
    CameraIds: [...p.cameraIds],
    deathMarkers: p.deathMarkers, // explicit [JsonPropertyName] on the C# side
    LearnedDaySpeed: p.learnedDaySpeed,
    LearnedNightSpeed: p.learnedNightSpeed,
    ChatCommandsEnabled: p.chatCommandsEnabled,
    CmdPop: p.cmdPop,
    CmdList: p.cmdList,
    CmdTime: p.cmdTime,
    CmdPromote: p.cmdPromote,
    CmdDeepSea: p.cmdDeepSea,
    CmdCargo: p.cmdCargo,
    CmdAfk: p.cmdAfk,
    CmdOilRig: p.cmdOilRig,
    CmdHeli: p.cmdHeli,
    CmdVendor: p.cmdVendor,
    CmdUpkeepDetail: p.cmdUpkeepDetail,
    ChatCommandPrefix: p.chatCommandPrefix,
    ChatCommandDelaySeconds: p.chatCommandDelaySeconds,
    ChatResponseDelaySeconds: p.chatResponseDelaySeconds,
    SwitchCommandMappings: serializeMappings(p.switchCommandMappings),
    UpkeepCommandMappings: serializeMappings(p.upkeepCommandMappings),
    CmdCustomTimer: p.cmdCustomTimer,
    AlertCustomTimer: p.alertCustomTimer,
    DiscordWebhookChatAlertsUrl: p.discordWebhookChatAlertsUrl,
    DiscordWebhookChatAlertsMention: p.discordWebhookChatAlertsMention,
    DiscordWebhookChatAlertsEnabled: p.discordWebhookChatAlertsEnabled,
    DiscordWebhookChatAlertsTts: p.discordWebhookChatAlertsTts,
    DiscordWebhookChatAlertsExclusive: p.discordWebhookChatAlertsExclusive,
    TimerAlarmEnabled: p.timerAlarmEnabled,
    TimerAlarmAudioPath: p.timerAlarmAudioPath,
    TimerCountdownAudioPath: p.timerCountdownAudioPath,
    TimerAlarmSnoozeMinutes: p.timerAlarmSnoozeMinutes,
    TimerAlarmBeepDurationSeconds: p.timerAlarmBeepDurationSeconds,
    CustomTimers: p.customTimers.map(serializeTimer),
    RustMapsMapId: p.rustMapsMapId,
    RustMapsFetchTime: iso(p.rustMapsFetchTimeMs),
    RustMapsWipeTime: iso(p.rustMapsWipeTimeMs),
    WipeTime: iso(p.wipeTimeMs),
    LogicRules: p.logicRules,
    IsLogicEngineActive: p.isLogicEngineActive,
    DeviceAutomationRules: p.deviceAutomationRules,
    IsDeviceAutomationActive: p.isDeviceAutomationActive,
    SubscribedTeammateSteamIds: [...p.subscribedTeammateSteamIds],
    ...p.extra, // unknown-from-our-view properties survive verbatim
  };
}

export function serializeDevices(devices: SmartDeviceNode[]): Record<string, unknown>[] {
  const walk = (d: SmartDeviceNode): Record<string, unknown> => ({
    EntityId: d.entityId,
    Kind: d.kind,
    Name: d.name,
    Alias: d.alias,
    IsGroup: d.isGroup,
    Children: d.children.map(walk),
    IsMissing: d.isMissing,
    CustomIconId: d.customIconId,
    CustomIconShortName: d.customIconShortName,
    InGameAlarmTitle: d.inGameAlarmTitle,
    OilRigTriggerTarget: d.oilRigTriggerTarget,
  });
  return devices.map(walk);
}

/** ProfileDataModule.LoadProfiles — corrupt file → empty list, never a crash. */
export function parseProfilesJson(json: string): ServerProfileData[] {
  try {
    const arr = JSON.parse(json) as unknown;
    if (!Array.isArray(arr)) return [];
    return arr.map((raw) => parseServerProfile(raw));
  } catch {
    return [];
  }
}

export function serializeProfilesJson(profiles: readonly ServerProfileData[]): string {
  return JSON.stringify(profiles.map(serializeServerProfile), null, 2);
}

// --------------------------------------------------------------- SyncChatCommands

/**
 * Lowest index whose command name is still free. Using Count + 1 collided after a deletion:
 * removing the middle of [upkeep, upkeep2, upkeep3] leaves Count = 2, so the next device
 * would claim "upkeep3" a second time and both would answer the same chat command.
 */
function nextFreeCommandIndex(mappings: readonly ChatCommandMapping[], nameFor: (i: number) => string): number {
  const taken = new Set(mappings.map((m) => m.command.toLowerCase()));
  for (let i = 1; ; i++) {
    if (!taken.has(nameFor(i).toLowerCase())) return i;
  }
}

const isSwitch = (d: SmartDeviceNode): boolean => d.kind === "SmartSwitch";
const isTcMonitor = (d: SmartDeviceNode): boolean =>
  (d.kind === "StorageMonitor" || d.kind === "Storage Monitor") &&
  // Storage details live on an optional bag; absent ⇒ counts as a TC candidate (parity L541).
  ((d as SmartDeviceNode & { storage?: { isToolCupboard?: boolean; itemsCount?: number } | null }).storage == null ||
    (d as SmartDeviceNode & { storage: { isToolCupboard?: boolean } }).storage.isToolCupboard === true ||
    ((d as SmartDeviceNode & { storage: { itemsCount?: number } }).storage.itemsCount ?? 0) === 0);

/** AllDevices parity: groups transparent, leaves AND group containers flattened depth-first. */
export function flattenAllDevices(devices: readonly SmartDeviceNode[]): SmartDeviceNode[] {
  const list: SmartDeviceNode[] = [];
  const walk = (source: readonly SmartDeviceNode[]): void => {
    for (const d of source) {
      if (!d.isGroup) list.push(d);
      walk(d.children);
    }
  };
  walk(devices);
  return list;
}

/** ServerProfile.SyncChatCommands parity (L513-570). */
export function syncChatCommands(profile: ServerProfileData): void {
  const all = flattenAllDevices(profile.devices);

  // Sync Switches
  const switches = all.filter(isSwitch);
  const validSwitchIds = new Set(switches.map((s) => s.entityId));
  for (let i = profile.switchCommandMappings.length - 1; i >= 0; i--) {
    const m = profile.switchCommandMappings[i]!;
    if (m.entityId !== 0 && !validSwitchIds.has(m.entityId)) profile.switchCommandMappings.splice(i, 1);
  }
  for (const sw of switches) {
    if (!profile.switchCommandMappings.some((m) => m.entityId === sw.entityId)) {
      const next = nextFreeCommandIndex(profile.switchCommandMappings, (i) => `switch${i}`);
      profile.switchCommandMappings.push({
        label: `Switch ${next}`,
        command: `switch${next}`,
        entityId: sw.entityId,
      });
    }
  }

  // Sync Upkeep (Storage Monitors on TCs)
  const tcs = all.filter(isTcMonitor);
  const validTcIds = new Set(tcs.map((s) => s.entityId));
  for (let i = profile.upkeepCommandMappings.length - 1; i >= 0; i--) {
    const m = profile.upkeepCommandMappings[i]!;
    if (m.entityId !== 0 && !validTcIds.has(m.entityId)) profile.upkeepCommandMappings.splice(i, 1);
  }
  for (const tc of tcs) {
    if (!profile.upkeepCommandMappings.some((m) => m.entityId === tc.entityId)) {
      const next = nextFreeCommandIndex(profile.upkeepCommandMappings, (i) => (i === 1 ? "upkeep" : `upkeep${i}`));
      profile.upkeepCommandMappings.push({
        label: next === 1 ? "Upkeep" : `Upkeep ${next}`,
        command: next === 1 ? "upkeep" : `upkeep${next}`,
        entityId: tc.entityId,
      });
    }
  }
}
