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
