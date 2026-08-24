/** Typed access to the preload bridge. The only sanctioned way the renderer talks to the main process. */
import type { IpcChannelName } from "@rpd/shared";

type RpdBridge = {
  invoke(channel: IpcChannelName | string, payload: unknown): Promise<InvokeResult<unknown>>;
  getInfo(): Promise<InvokeResult<unknown>>;
  log(level: "debug" | "info" | "warn" | "error", scope: string, message: string): Promise<InvokeResult<unknown>>;
  cloudLogin(email: string, password: string): Promise<InvokeResult<unknown>>;
  cloudBootstrap(): Promise<InvokeResult<unknown>>;
  cloudLogout(): Promise<InvokeResult<unknown>>;
  wipeGetStatus(): Promise<InvokeResult<unknown>>;
  wipeGetPlayer(steamId: string): Promise<InvokeResult<unknown>>;
  wipeGetMap(): Promise<InvokeResult<unknown>>;
  wipeGetCloudArchives(): Promise<InvokeResult<unknown>>;
  wipeRestoreCloudArchive(archiveId: string): Promise<InvokeResult<unknown>>;
  wipeDeleteCloudArchive(archiveId: string): Promise<InvokeResult<unknown>>;
  wipeDeleteAllCloud(): Promise<InvokeResult<unknown>>;
  deathsGetStats(payload: unknown): Promise<InvokeResult<unknown>>;
  deathsClear(): Promise<InvokeResult<unknown>>;
  settingsGetWipe(): Promise<InvokeResult<unknown>>;
  settingsSetWipe(payload: unknown): Promise<InvokeResult<unknown>>;
  getUiPrefs(): Promise<InvokeResult<unknown>>;
  setUiPrefs(patch: { sidebarPinned?: boolean; sidebarWidth?: number }): Promise<InvokeResult<unknown>>;
  connectProfile(matchKey: string, useProxy?: boolean): Promise<InvokeResult<unknown>>;
  disconnect(): Promise<InvokeResult<unknown>>;
  connectionStatus(): Promise<InvokeResult<unknown>>;
  listProfiles(): Promise<InvokeResult<unknown>>;
  pairProfile(link: string, name?: string): Promise<InvokeResult<unknown>>;
  getDevices(matchKey: string): Promise<InvokeResult<unknown>>;
  saveDevices(matchKey: string, devices: unknown[]): Promise<InvokeResult<unknown>>;
  activateProfile(matchKey: string): Promise<InvokeResult<unknown>>;
  exportDevices(matchKey: string): Promise<InvokeResult<unknown>>;
  importDevicesPreview(matchKey: string): Promise<InvokeResult<unknown>>;
  applyImportedDevices(payload: unknown): Promise<InvokeResult<unknown>>;
  deleteDevice(payload: unknown): Promise<InvokeResult<unknown>>;
  deviceAutomationGetRules(matchKey: string): Promise<InvokeResult<unknown>>;
  deviceAutomationSaveRules(payload: unknown): Promise<InvokeResult<unknown>>;
  raidGetData(): Promise<InvokeResult<unknown>>;
  raidCalculate(payload: unknown): Promise<InvokeResult<unknown>>;
  recyclerGetData(): Promise<InvokeResult<unknown>>;
  recyclerCalculate(payload: unknown): Promise<InvokeResult<unknown>>;
  onPush(listener: (payload: { stream: string; event: unknown }) => void): () => void;
  logicStatus(): Promise<InvokeResult<unknown>>;
  logicStop(): Promise<InvokeResult<unknown>>;
  logicRun(ruleId: string): Promise<InvokeResult<unknown>>;
  logicGetRules(matchKey: string): Promise<InvokeResult<unknown>>;
  logicSaveRules(payload: unknown): Promise<InvokeResult<unknown>>;
  logicGetRule(matchKey: string, ruleId: string): Promise<InvokeResult<unknown>>;
  logicSaveRule(payload: unknown): Promise<InvokeResult<unknown>>;
  logicGetTimers(matchKey: string): Promise<InvokeResult<unknown>>;
  logicAddTimer(payload: unknown): Promise<InvokeResult<unknown>>;
  logicRemoveTimer(payload: unknown): Promise<InvokeResult<unknown>>;
};

export interface InvokeSuccess<T> {
  ok: true;
  data: T;
}
export interface InvokeFailure {
  ok: false;
  error: { code: string; message: string };
}
export type InvokeResult<T> = InvokeSuccess<T> | InvokeFailure;

declare global {
  interface Window {
    rpd: RpdBridge;
  }
}

export const bridge: RpdBridge = (globalThis as typeof globalThis & { rpd: RpdBridge }).rpd;

export async function invoke<T>(channel: IpcChannelName, payload?: unknown): Promise<T> {
  const result = await bridge.invoke(channel, payload);
  if (result.ok) return result.data as T;
  throw new Error(`ipc ${channel} failed [${result.error.code}]: ${result.error.message}`);
}

export interface CloudUser {
  id: string;
  steamId: string | null;
  name: string | null;
  displayName: string | null;
  email: string | null;
  providers: string[];
  hasPassword: boolean;
}

export interface PlayerWipeCapabilities {
  planCode: string;
  isTrackerAvailable: boolean;
  canTrackTeam: boolean;
  canUseCloudSync: boolean;
  canUseAdvancedViews: boolean;
  canUseRouteReplay: boolean;
  canExport: boolean;
  maxTrackedPlayers: number;
  retainedWipes: number;
  cloudRetentionDays: number;
  fetchedAt: string;
}

export interface CloudBootstrap {
  signedIn: boolean;
  user: CloudUser | null;
  capabilities: PlayerWipeCapabilities | null;
  error: string | null;
}

export async function cloudLogin(email: string, password: string): Promise<CloudUser> {
  const result = await bridge.cloudLogin(email, password);
  if (result.ok) return (result.data as { user: CloudUser }).user;
  throw new Error(result.error.message);
}

export async function cloudBootstrap(): Promise<CloudBootstrap> {
  const result = await bridge.cloudBootstrap();
  if (result.ok) return result.data as CloudBootstrap;
  throw new Error(result.error.message);
}

export async function cloudLogout(): Promise<void> {
  const result = await bridge.cloudLogout();
  if (!result.ok) throw new Error(result.error.message);
}

export interface WipePlayer {
  steamId: string;
  name: string;
  observationCount: number;
  summary: {
    coverageSeconds: number;
    unknownSeconds: number;
    movingSeconds: number;
    stationarySeconds: number;
    afkSeconds: number;
    deadSeconds: number;
    offlineSeconds: number;
    estimatedDistance: number;
    deaths: number;
    monumentVisits: Array<{ name: string; startUtc: string; endUtc: string; entryX: number | null; entryY: number | null; exitX: number | null; exitY: number | null }>;
  };
  insights: {
    firstSeenUtc: string | null;
    lastSeenUtc: string | null;
    sessionCount: number;
    topMonument: string | null;
    topMonumentSeconds: number;
    topMonumentVisits: number;
    longestBlindGapSeconds: number;
    longestBlindGapStartUtc: string | null;
    peakHourLocal: number | null;
    peakHourActiveSeconds: number;
    currentState: "moving" | "stationary" | "afk" | "dead" | "offline" | "unknown";
    currentLocationType: "monument" | "base" | "open" | "unknown";
    currentLocationName: string | null;
    currentGrid: string | null;
    currentAsOfUtc: string | null;
    isLikelyOnline: boolean;
  };
  observations: Array<{ timestampUtc: string; x: number | null; y: number | null; state: "moving" | "stationary" | "afk" | "dead" | "offline" | "unknown"; locationType: "monument" | "base" | "open" | "unknown"; locationName: string | null; grid: string | null; event: "death" | "respawn" | null; sessionId: string }>;
  segments: Array<{ startUtc: string; endUtc: string; state: "moving" | "stationary" | "afk" | "dead" | "offline" | "unknown" }>;
}

export interface WipeStatus {
  serverKey: string | null;
  wipeKey: string | null;
  sessionId: string | null;
  players: WipePlayer[];
}

export async function wipeGetStatus(): Promise<WipeStatus> {
  const result = await bridge.wipeGetStatus();
  if (result.ok) return result.data as WipeStatus;
  throw new Error(result.error.message);
}

export async function wipeGetPlayer(steamId: string): Promise<WipePlayer | null> {
  const result = await bridge.wipeGetPlayer(steamId);
  if (result.ok) return (result.data as { player: WipePlayer | null }).player;
  throw new Error(result.error.message);
}

export interface WipeMap {
  pngBase64: string;
  imageWidth: number;
  imageHeight: number;
  worldSize: number;
  worldRectX: number;
  worldRectY: number;
  worldRectWidth: number;
  worldRectHeight: number;
}

export async function wipeGetMap(): Promise<WipeMap | null> {
  const result = await bridge.wipeGetMap();
  if (result.ok) return (result.data as { map: WipeMap | null }).map;
  throw new Error(result.error.message);
}

export interface CloudArchive {
  id: string;
  serverKey: string;
  serverName: string;
  wipeKey: string;
  wipeStartedAtUtc: string | null;
  firstObservedAtUtc: string | null;
  lastObservedAtUtc: string | null;
  playerCount: number | null;
  storedBytes: number | null;
  players: Array<{ steamId: string; dayCount: number }>;
}

export async function wipeGetCloudArchives(): Promise<CloudArchive[]> {
  const result = await bridge.wipeGetCloudArchives();
  if (result.ok) return (result.data as { archives: CloudArchive[] }).archives;
  throw new Error(result.error.message);
}

export async function wipeRestoreCloudArchive(archiveId: string): Promise<{ archiveId: string; players: number; days: number; observations: number; isCurrentWipe: boolean }> {
  const result = await bridge.wipeRestoreCloudArchive(archiveId);
  if (result.ok) return result.data as { archiveId: string; players: number; days: number; observations: number; isCurrentWipe: boolean };
  throw new Error(result.error.message);
}

export async function wipeDeleteCloudArchive(archiveId: string): Promise<boolean> {
  const result = await bridge.wipeDeleteCloudArchive(archiveId);
  if (result.ok) return (result.data as { deleted: boolean }).deleted;
  throw new Error(result.error.message);
}

export async function wipeDeleteAllCloud(): Promise<number> {
  const result = await bridge.wipeDeleteAllCloud();
  if (result.ok) return (result.data as { deleted: number }).deleted;
  throw new Error(result.error.message);
}

export interface DeathStats {
  total: number;
  victims: number;
  avgSurvival: string;
  longestSurvival: string;
  peakHour: string;
  deadliestPlace: string;
  deadliestGrid: string;
  byArea: Array<{ name: string; type: "monument" | "base" | "open"; deaths: number; percent: number }>;
  byVictim: Array<{ victim: string; deaths: number; avgSurvival: string }>;
  byLocation: Array<{ location: string; type: "monument" | "base" | "open"; deaths: number }>;
  recent: Array<{ victim: string; type: "monument" | "base" | "open"; location: string; grid: string; died: string }>;
  deathsPerDay: Array<{ day: string; count: number }>;
}

export async function getDeathStats(filters: { search?: string; player?: string; type?: "all" | "monument" | "base" | "open"; range?: "all" | "24h" | "7d" } = {}): Promise<{ serverKey: string | null; players: string[]; summary: DeathStats }> {
  const result = await bridge.deathsGetStats(filters);
  if (result.ok) return result.data as { serverKey: string | null; players: string[]; summary: DeathStats };
  throw new Error(result.error.message);
}

export async function clearDeathLog(): Promise<boolean> {
  const result = await bridge.deathsClear();
  if (result.ok) return (result.data as { cleared: boolean }).cleared;
  throw new Error(result.error.message);
}

export interface WipeSettings {
  enabled: boolean;
  cloudBackupEnabled: boolean;
}

export async function getWipeSettings(): Promise<WipeSettings> {
  const result = await bridge.settingsGetWipe();
  if (result.ok) return result.data as WipeSettings;
  throw new Error(result.error.message);
}

export async function setWipeSettings(patch: Partial<WipeSettings>): Promise<WipeSettings> {
  const result = await bridge.settingsSetWipe(patch);
  if (result.ok) return result.data as WipeSettings;
  throw new Error(result.error.message);
}

/** Typed helpers for declared channels (components never touch window.rpd directly). */

export interface UiPrefs {
  sidebarPinned: boolean;
  sidebarWidth: number;
}

export function getUiPrefs(): Promise<UiPrefs | null> {
  return bridge.getUiPrefs().then((r) => (r.ok ? (r.data as UiPrefs) : null));
}

const uiPrefsTimers = new Map<string, ReturnType<typeof setTimeout>>();
const UI_PREFS_DEBOUNCE_MS = 400;

/** Debounced persistence of shell prefs; rapid changes coalesce into one store write. */
export function setUiPrefsDebounced(patch: { sidebarPinned?: boolean; sidebarWidth?: number }): void {
  const key = "uiPrefs";
  const existing = uiPrefsTimers.get(key);
  if (existing) clearTimeout(existing);
  uiPrefsTimers.set(
    key,
    setTimeout(() => {
      uiPrefsTimers.delete(key);
      void bridge.setUiPrefs(patch).then((r) => {
        if (!r.ok) void bridge.log("warn", "shell", `persisting ui prefs failed: ${r.error.message}`);
      });
    }, UI_PREFS_DEBOUNCE_MS),
  );
}

// ------------------------------------------------------------------ profile/* helpers

export interface ProfileSummary {
  matchKey: string;
  name: string;
  host: string;
  port: number;
  steamId64: string;
  deviceCount: number;
}

export interface ConnSnapshot {
  connected: boolean;
  activeProxy: "direct" | "proxy" | null;
  host: string | null;
  port: number | null;
  consecutiveTimeouts: number;
  teamChatPrimed: boolean;
  clanChatPrimed: boolean;
}

export async function connectProfile(matchKey: string, useProxy?: boolean): Promise<ConnSnapshot> {
  const r = await bridge.connectProfile(matchKey, useProxy);
  if (r.ok) return r.data as ConnSnapshot;
  throw new Error(r.error.message);
}

export async function disconnect(): Promise<ConnSnapshot> {
  const r = await bridge.disconnect();
  if (r.ok) return r.data as ConnSnapshot;
  throw new Error(r.error.message);
}

export async function getConnectionStatus(): Promise<ConnSnapshot> {
  const r = await bridge.connectionStatus();
  if (r.ok) return r.data as ConnSnapshot;
  throw new Error(r.error.message);
}

export interface DeviceNode {
  entityId: number;
  kind: string | null;
  name: string | null;
  alias: string | null;
  isGroup: boolean;
  children: DeviceNode[];
  isMissing: boolean;
  customIconId?: number | null;
  customIconShortName?: string | null;
  inGameAlarmTitle?: string | null;
  oilRigTriggerTarget?: string | null;
  pairedX?: number | null;
  pairedY?: number | null;
  pairedBySteamId?: string | null;
  pairedLocationCapturedAtMs?: number | null;
}

export async function listProfiles(): Promise<ProfileSummary[]> {
  const r = await bridge.listProfiles();
  return r.ok ? (r.data as { profiles: ProfileSummary[] }).profiles : [];
}

export async function pairProfile(link: string, name?: string): Promise<{ activated: boolean; profile: ProfileSummary }> {
  const r = await bridge.pairProfile(link, name);
  if (r.ok) return r.data as { activated: boolean; profile: ProfileSummary };
  throw new Error(r.error.message);
}

export async function getDevices(matchKey: string): Promise<DeviceNode[] | null> {
  const r = await bridge.getDevices(matchKey);
  return r.ok ? ((r.data as { devices: DeviceNode[]; found: boolean }).found ? (r.data as { devices: DeviceNode[] }).devices : null) : null;
}

export async function saveDevices(matchKey: string, devices: DeviceNode[]): Promise<boolean> {
  const r = await bridge.saveDevices(matchKey, devices);
  return r.ok ? (r.data as { saved: boolean }).saved : false;
}

export async function activateProfile(matchKey: string): Promise<boolean> {
  const r = await bridge.activateProfile(matchKey);
  return r.ok ? (r.data as { activated: boolean }).activated : false;
}

export interface DeviceImportCandidate {
  id: string;
  ownerSteamId: string;
  ownerName: string;
  entityId: number;
  kind: string | null;
  name: string | null;
  alias: string | null;
  alreadyPresent: boolean;
  fromPreviousWipe: boolean;
  serverName: string;
  existsState: "?" | "ok" | "missing" | "err" | "local";
  originalDto: unknown;
}

export async function exportDevices(matchKey: string): Promise<{ saved: boolean; canceled: boolean; path: string | null }> {
  const r = await bridge.exportDevices(matchKey);
  if (r.ok) return r.data as { saved: boolean; canceled: boolean; path: string | null };
  throw new Error(`profile/exportDevices failed [${r.error.code}]: ${r.error.message}`);
}

export async function importDevicesPreview(matchKey: string): Promise<{ canceled: boolean; path: string | null; candidates: DeviceImportCandidate[] }> {
  const r = await bridge.importDevicesPreview(matchKey);
  if (r.ok) return r.data as { canceled: boolean; path: string | null; candidates: DeviceImportCandidate[] };
  throw new Error(`profile/importPreview failed [${r.error.code}]: ${r.error.message}`);
}

export async function applyImportedDevices(matchKey: string, devices: unknown[]): Promise<{ saved: boolean; imported: number }> {
  const r = await bridge.applyImportedDevices({ matchKey, devices });
  if (r.ok) return r.data as { saved: boolean; imported: number };
  throw new Error(`profile/applyImport failed [${r.error.code}]: ${r.error.message}`);
}

export async function deleteDevice(matchKey: string, entityId: number): Promise<{ removed: boolean; reason: string }> {
  const r = await bridge.deleteDevice({ matchKey, entityId });
  if (r.ok) return r.data as { removed: boolean; reason: string };
  throw new Error(`profile/deleteDevice failed [${r.error.code}]: ${r.error.message}`);
}

// ------------------------------------------------------------------ deviceAutomation/* helpers

export interface DeviceAutomationRuleDto {
  id: string;
  name: string;
  isEnabled: boolean;
  isExpanded: boolean;
  conditionType: "PlayerProximity" | "GameTime";
  playerMatchMode: "AnyOnline" | "AllOnline" | "Specific" | "SpecificOffline" | "AnyOffline" | "AllOffline";
  specificPlayerSteamId: string;
  locationEntityId: number;
  distanceMeters: number;
  startTime: string;
  endTime: string;
  targetEntityId: number;
  matchedState: boolean;
  unmatchedState: boolean;
}

export interface DeviceAutomationPageData {
  found: boolean;
  isActive: boolean;
  rules: DeviceAutomationRuleDto[];
}

export async function getDeviceAutomationRules(matchKey: string): Promise<DeviceAutomationPageData> {
  const r = await bridge.deviceAutomationGetRules(matchKey);
  if (r.ok) return r.data as DeviceAutomationPageData;
  throw new Error(`deviceAutomation/getRules failed [${r.error.code}]: ${r.error.message}`);
}

export async function saveDeviceAutomationRules(
  matchKey: string,
  isActive: boolean,
  rules: DeviceAutomationRuleDto[],
): Promise<boolean> {
  const r = await bridge.deviceAutomationSaveRules({ matchKey, isActive, rules });
  return r.ok ? (r.data as { saved: boolean }).saved : false;
}

// ------------------------------------------------------------------ raid/* helpers

export interface RaidResourceDto {
  shortname: string;
  itemId: number;
  displayName: string;
  amount: number;
}

export interface RaidSourceDto {
  sourceId: number;
  prefabName: string;
  itemId: number | null;
  itemShortname: string;
  itemSlug: string;
  itemCategorySlug: string;
  displayName: string;
  kind: string;
  rawDamage: number;
  craftCost: RaidResourceDto[] | null;
  workbenchLevelRequired: number | null;
}

export interface RaidTargetDto {
  targetId: number;
  prefabName: string;
  itemId: number | null;
  itemShortname: string | null;
  itemSlug: string | null;
  itemCategorySlug: string | null;
  buildingSlug: string | null;
  buildingImage: string | null;
  displayName: string;
  buildingTier: string | null;
  componentType: string;
  startHealth: number;
  category: string;
}

export interface RaidMethodDto {
  source: RaidSourceDto;
  requiredItems: number;
  damagePerItem: number;
  totalDamage: number;
  overkill: number;
  resources: RaidResourceDto[];
  hasCraftCost: boolean;
}

export interface RaidDataPage {
  sources: RaidSourceDto[];
  targets: RaidTargetDto[];
}

export interface RaidCalculation {
  methods: RaidMethodDto[];
  recommended: RaidMethodDto | null;
  combination: RaidMethodDto[];
  resources: RaidResourceDto[];
  items: Array<{ source: RaidSourceDto; amount: number }>;
}

export async function getRaidData(): Promise<RaidDataPage> {
  const r = await bridge.raidGetData();
  if (r.ok) return r.data as RaidDataPage;
  throw new Error(`raid/getData failed [${r.error.code}]: ${r.error.message}`);
}

export async function calculateRaid(payload: {
  targetId: number;
  targetQuantity: number;
  sourceIds: number[];
  mode: "LowestSulfur" | "LowestTotalResources" | "FewestRaidItems" | "Custom";
}): Promise<RaidCalculation> {
  const r = await bridge.raidCalculate(payload);
  if (r.ok) return r.data as RaidCalculation;
  throw new Error(`raid/calculate failed [${r.error.code}]: ${r.error.message}`);
}

// ------------------------------------------------------------------ recycler/* helpers

export interface RecyclerItemDto {
  id: string;
  shortName: string;
  displayName: string;
  category: string;
  stackSize: number;
}

export interface RecyclerMetricDto {
  expected: number;
  guaranteed: number;
  chance: number;
  chancePercent: number;
  min: number;
  max: number;
}

export interface RecyclerOutputDto {
  shortName: string;
  displayName: string;
  wild: RecyclerMetricDto;
  safe: RecyclerMetricDto;
}

export interface RecyclerCalculation {
  outputs: RecyclerOutputDto[];
  wildSeconds: number;
  safeSeconds: number;
}

export async function getRecyclerData(): Promise<RecyclerItemDto[]> {
  const r = await bridge.recyclerGetData();
  if (r.ok) return (r.data as { items: RecyclerItemDto[] }).items;
  throw new Error(`recycler/getData failed [${r.error.code}]: ${r.error.message}`);
}

export async function calculateRecycler(quantities: Array<{ shortName: string; quantity: number }>): Promise<RecyclerCalculation> {
  const r = await bridge.recyclerCalculate({ quantities });
  if (r.ok) return r.data as RecyclerCalculation;
  throw new Error(`recycler/calculate failed [${r.error.code}]: ${r.error.message}`);
}

// ------------------------------------------------------------------ logic/* helpers

export interface LogicStatus {
  activeKey: string | null;
  isRunning: boolean;
  currentRuleName: string | null;
  currentStepNumber: number;
  currentStepType: string | null;
  pendingRules: string[];
}

export interface RuleHeaderDto {
  id: string;
  name: string;
  isEnabled: boolean;
  isLoopEnabled: boolean;
  loopCount: number;
  triggerType: "SmartAlarm" | "SmartSwitch" | "ChatCommand" | "RuleTriggered" | "RuleCompleted";
  triggerEntityId: number;
  triggerCommand: string;
  triggerRuleId: string;
  triggerState: boolean;
  conditionOperator: "NONE" | "AND" | "OR";
  conditionDeviceEntityId: number;
  conditionDeviceState: boolean;
  stepCount: number;
}

export interface RulesPageData {
  found: boolean;
  isEngineActive: boolean;
  rules: RuleHeaderDto[];
}

export async function getLogicStatus(): Promise<LogicStatus> {
  const r = await bridge.logicStatus();
  if (r.ok) return r.data as LogicStatus;
  throw new Error(`logic/status failed [${r.error.code}]: ${r.error.message}`);
}

export async function stopLogic(): Promise<void> {
  const r = await bridge.logicStop();
  if (!r.ok) throw new Error(`logic/stop failed: ${r.error.message}`);
}

export async function runRule(ruleId: string): Promise<boolean> {
  const r = await bridge.logicRun(ruleId);
  return r.ok ? (r.data as { accepted: boolean }).accepted : false;
}

export async function getRules(matchKey: string): Promise<RulesPageData> {
  const r = await bridge.logicGetRules(matchKey);
  if (r.ok) return r.data as RulesPageData;
  throw new Error(`logic/getRules failed [${r.error.code}]: ${r.error.message}`);
}

export interface RuleHeaderInput extends Omit<RuleHeaderDto, "stepCount"> {}

export async function saveRules(
  matchKey: string,
  isEngineActive: boolean,
  rules: RuleHeaderInput[],
): Promise<boolean> {
  const r = await bridge.logicSaveRules({ matchKey, isEngineActive, rules });
  return r.ok ? (r.data as { saved: boolean }).saved : false;
}

/** Full rule incl. steps (step editor surface). */
export interface StepDto {
  stepType: "Wait" | "Toggle" | "CheckAvailability" | "StartTimer";
  timerMinutes?: number;
  timerTarget?: "Custom" | "SmallOilRig" | "LargeOilRig";
  timerName?: string;
  showCrateOnMap?: boolean;
  alarmTextHint?: string;
  waitSeconds?: number;
  targetEntityId?: number;
  targetGroupName?: string;
  toggleState?: boolean | null;
  conditionOperator?:
    | "IS_OFFLINE"
    | "IS_ONLINE"
    | "ALL_OFFLINE"
    | "ANY_OFFLINE"
    | "ALL_ONLINE"
    | "ANY_ONLINE";
  conditionDeviceIdsCsv?: string;
  conditionalSteps?: StepDto[];
}

export interface FullRuleDto {
  id: string;
  name: string;
  isEnabled: boolean;
  isLoopEnabled: boolean;
  loopCount: number;
  triggerType: RuleHeaderDto["triggerType"];
  triggerEntityId: number;
  triggerCommand: string;
  triggerRuleId: string;
  triggerState: boolean;
  conditionOperator: RuleHeaderDto["conditionOperator"];
  conditionDeviceEntityId: number;
  conditionDeviceState: boolean;
  steps: StepDto[];
}

export async function getRule(matchKey: string, ruleId: string): Promise<FullRuleDto | null> {
  const r = await bridge.logicGetRule(matchKey, ruleId);
  if (r.ok) return (r.data as { found: boolean; rule: FullRuleDto | null }).rule;
  throw new Error(`logic/getRule failed [${r.error.code}]: ${r.error.message}`);
}

export async function saveRule(matchKey: string, rule: FullRuleDto): Promise<boolean> {
  const r = await bridge.logicSaveRule({ matchKey, rule });
  return r.ok ? (r.data as { saved: boolean }).saved : false;
}

// ---- Custom timers (timers panel) ---------------------------------------------

export interface TimerDto {
  id: string;
  name: string;
  command: string;
  endTimeUtcMs: number;
  enableCountdownAudio: boolean;
  enableAlarmAudio: boolean;
}

export async function getTimers(matchKey: string): Promise<TimerDto[]> {
  const r = await bridge.logicGetTimers(matchKey);
  if (r.ok) return (r.data as { timers: TimerDto[] }).timers;
  throw new Error(`logic/getTimers failed [${r.error.code}]: ${r.error.message}`);
}

export type AddTimerResult =
  | { ok: true; id: string }
  | { ok: false; reason: "limit" | "letter" | "duration" };

export async function addTimer(
  matchKey: string,
  name: string,
  hours: number,
  minutes: number,
  seconds: number,
): Promise<AddTimerResult> {
  const r = await bridge.logicAddTimer({ matchKey, name, hours, minutes, seconds });
  if (!r.ok) throw new Error(`logic/addTimer failed [${r.error.code}]: ${r.error.message}`);
  const d = r.data as { ok: boolean; id: string; reason: "limit" | "letter" | "duration" | null };
  return d.ok ? { ok: true, id: d.id } : { ok: false, reason: d.reason ?? "duration" };
}

export async function removeTimer(matchKey: string, id: string): Promise<boolean> {
  const r = await bridge.logicRemoveTimer({ matchKey, id });
  return r.ok ? (r.data as { removed: boolean }).removed : false;
}
