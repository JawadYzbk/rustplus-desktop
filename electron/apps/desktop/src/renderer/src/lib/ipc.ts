/** Typed access to the preload bridge. The only sanctioned way the renderer talks to the main process. */
import type { IpcChannelName } from "@rpd/shared";

type RpdBridge = {
  invoke(channel: IpcChannelName | string, payload: unknown): Promise<InvokeResult<unknown>>;
  getInfo(): Promise<InvokeResult<unknown>>;
  log(level: "debug" | "info" | "warn" | "error", scope: string, message: string): Promise<InvokeResult<unknown>>;
  getUiPrefs(): Promise<InvokeResult<unknown>>;
  setUiPrefs(patch: { sidebarPinned?: boolean; sidebarWidth?: number }): Promise<InvokeResult<unknown>>;
  listProfiles(): Promise<InvokeResult<unknown>>;
  getDevices(matchKey: string): Promise<InvokeResult<unknown>>;
  saveDevices(matchKey: string, devices: unknown[]): Promise<InvokeResult<unknown>>;
  activateProfile(matchKey: string): Promise<InvokeResult<unknown>>;
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

export const bridge: RpdBridge = window.rpd;

export async function invoke<T>(channel: IpcChannelName, payload?: unknown): Promise<T> {
  const result = await bridge.invoke(channel, payload);
  if (result.ok) return result.data as T;
  throw new Error(`ipc ${channel} failed [${result.error.code}]: ${result.error.message}`);
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
}

export async function listProfiles(): Promise<ProfileSummary[]> {
  const r = await bridge.listProfiles();
  return r.ok ? (r.data as { profiles: ProfileSummary[] }).profiles : [];
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
