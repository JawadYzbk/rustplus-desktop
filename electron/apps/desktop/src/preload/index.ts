/**
 * Preload bridge — the ONLY renderer-facing surface.
 *
 * Security posture (ELECTRON_ARCHITECTURE §3):
 *  - contextBridge exposes exactly one namespace (`rpd`) with typed methods.
 *  - Every invoke is validated against the shared zod contract before sending and after receiving.
 *  - No dynamic channel names: the allow-list is the shared registry, keys of which are compile-time known.
 */
import { contextBridge, ipcRenderer } from "electron";
import type { IpcChannelName, IpcError } from "@rpd/shared";
import { ipcChannels, pushChannel } from "@rpd/shared";

export interface InvokeSuccess<T> {
  ok: true;
  data: T;
}
export interface InvokeFailure {
  ok: false;
  error: IpcError;
}
export type InvokeResult<T> = InvokeSuccess<T> | InvokeFailure;

async function invoke<C extends IpcChannelName>(
  channel: C,
  payload: unknown,
): Promise<InvokeResult<unknown>> {
  const def = ipcChannels[channel];
  if (!def) {
    return { ok: false, error: { code: "unknown_channel", message: `channel "${channel}" is not declared` } };
  }
  const parsedRequest = def.request.safeParse(payload);
  if (!parsedRequest.success) {
    return { ok: false, error: { code: "bad_request", message: "request failed schema validation" } };
  }
  try {
    const raw: unknown = await ipcRenderer.invoke(channel, parsedRequest.data);
    return raw as InvokeResult<unknown>;
  } catch (err) {
    return {
      ok: false,
      error: { code: "handler_error", message: err instanceof Error ? err.message : String(err) },
    };
  }
}

const api = {
  /** Typed invoke for any declared channel (renderer code uses the wrappers below or this directly). */
  invoke,

  /** Bootstrap snapshot. */
  getInfo: () => invoke("app/getInfo", undefined) as Promise<InvokeResult<unknown>>,

  /** Forward a renderer log line into the main logger. */
  log: (level: "debug" | "info" | "warn" | "error", scope: string, message: string) =>
    invoke("app/logFromRenderer", { level, scope, message }) as Promise<InvokeResult<unknown>>,

  cloudLogin: (email: string, password: string) => invoke("cloud/login", { email, password }) as Promise<InvokeResult<unknown>>,
  cloudBootstrap: () => invoke("cloud/bootstrap", {}) as Promise<InvokeResult<unknown>>,
  cloudLogout: () => invoke("cloud/logout", {}) as Promise<InvokeResult<unknown>>,
  wipeGetStatus: () => invoke("wipe/getStatus", {}) as Promise<InvokeResult<unknown>>,
  wipeGetPlayer: (steamId: string) => invoke("wipe/getPlayer", { steamId }) as Promise<InvokeResult<unknown>>,
  wipeGetMap: () => invoke("wipe/getMap", {}) as Promise<InvokeResult<unknown>>,
  wipeGetCloudArchives: () => invoke("wipe/getCloudArchives", {}) as Promise<InvokeResult<unknown>>,
  wipeRestoreCloudArchive: (archiveId: string) => invoke("wipe/restoreCloudArchive", { archiveId }) as Promise<InvokeResult<unknown>>,
  wipeDeleteCloudArchive: (archiveId: string) => invoke("wipe/deleteCloudArchive", { archiveId }) as Promise<InvokeResult<unknown>>,
  wipeDeleteAllCloud: () => invoke("wipe/deleteAllCloud", {}) as Promise<InvokeResult<unknown>>,
  deathsGetStats: (payload: unknown) => invoke("deaths/getStats", payload) as Promise<InvokeResult<unknown>>,
  deathsClear: () => invoke("deaths/clear", {}) as Promise<InvokeResult<unknown>>,
  settingsGetWipe: () => invoke("settings/getWipe", {}) as Promise<InvokeResult<unknown>>,
  settingsSetWipe: (payload: unknown) => invoke("settings/setWipe", payload) as Promise<InvokeResult<unknown>>,

  /** Persisted shell preferences (sidebar). */
  getUiPrefs: () => invoke("uiPrefs/get", undefined) as Promise<InvokeResult<unknown>>,
  setUiPrefs: (patch: { sidebarPinned?: boolean; sidebarWidth?: number }) =>
    invoke("uiPrefs/set", patch) as Promise<InvokeResult<unknown>>,

  /** Server profiles + device trees (stage 5). Tokens never cross the bridge. */
  listProfiles: () => invoke("profile/list", undefined) as Promise<InvokeResult<unknown>>,
  pairProfile: (link: string, name?: string) => invoke("profile/pair", { link, ...(name ? { name } : {}) }) as Promise<InvokeResult<unknown>>,
  getDevices: (matchKey: string) =>
    invoke("profile/getDevices", { matchKey }) as Promise<InvokeResult<unknown>>,
  saveDevices: (matchKey: string, devices: unknown[]) =>
    invoke("profile/saveDevices", { matchKey, devices }) as Promise<InvokeResult<unknown>>,
  activateProfile: (matchKey: string) =>
    invoke("profile/activate", { matchKey }) as Promise<InvokeResult<unknown>>,
  exportDevices: (matchKey: string) =>
    invoke("profile/exportDevices", { matchKey }) as Promise<InvokeResult<unknown>>,
  importDevicesPreview: (matchKey: string) =>
    invoke("profile/importPreview", { matchKey }) as Promise<InvokeResult<unknown>>,
  applyImportedDevices: (payload: unknown) =>
    invoke("profile/applyImport", payload) as Promise<InvokeResult<unknown>>,
  deleteDevice: (payload: unknown) =>
    invoke("profile/deleteDevice", payload) as Promise<InvokeResult<unknown>>,

  deviceAutomationGetRules: (matchKey: string) =>
    invoke("deviceAutomation/getRules", { matchKey }) as Promise<InvokeResult<unknown>>,
  deviceAutomationSaveRules: (payload: unknown) =>
    invoke("deviceAutomation/saveRules", payload) as Promise<InvokeResult<unknown>>,
  raidGetData: () => invoke("raid/getData", undefined) as Promise<InvokeResult<unknown>>,
  raidCalculate: (payload: unknown) => invoke("raid/calculate", payload) as Promise<InvokeResult<unknown>>,
  recyclerGetData: () => invoke("recycler/getData", undefined) as Promise<InvokeResult<unknown>>,
  recyclerCalculate: (payload: unknown) => invoke("recycler/calculate", payload) as Promise<InvokeResult<unknown>>,

  /** Logic Engine control (stage 5). */
  logicStatus: () => invoke("logic/status", undefined) as Promise<InvokeResult<unknown>>,
  logicStop: () => invoke("logic/stop", undefined) as Promise<InvokeResult<unknown>>,
  logicRun: (ruleId: string) => invoke("logic/run", { ruleId }) as Promise<InvokeResult<unknown>>,
  logicGetRules: (matchKey: string) =>
    invoke("logic/getRules", { matchKey }) as Promise<InvokeResult<unknown>>,
  logicSaveRules: (payload: unknown) =>
    invoke("logic/saveRules", payload) as Promise<InvokeResult<unknown>>,
  logicGetRule: (matchKey: string, ruleId: string) =>
    invoke("logic/getRule", { matchKey, ruleId }) as Promise<InvokeResult<unknown>>,
  logicSaveRule: (payload: unknown) =>
    invoke("logic/saveRule", payload) as Promise<InvokeResult<unknown>>,
  logicGetTimers: (matchKey: string) =>
    invoke("logic/getTimers", { matchKey }) as Promise<InvokeResult<unknown>>,
  logicAddTimer: (payload: unknown) =>
    invoke("logic/addTimer", payload) as Promise<InvokeResult<unknown>>,
  logicRemoveTimer: (payload: unknown) =>
    invoke("logic/removeTimer", payload) as Promise<InvokeResult<unknown>>,

  /** One-way runtime event stream (connection lifecycle, polls, device state).
   * Returns an unsubscribe function. Payload shape: { stream, event }. */
  onPush: (listener: (payload: { stream: string; event: unknown }) => void): (() => void) => {
    const wrapped = (_event: unknown, payload: unknown): void => {
      const p = (payload ?? {}) as { stream?: string; event?: unknown };
      listener({ stream: p.stream ?? "", event: p.event });
    };
    ipcRenderer.on(pushChannel, wrapped);
    return () => ipcRenderer.removeListener(pushChannel, wrapped);
  },
};

export type RpdApi = typeof api;

contextBridge.exposeInMainWorld("rpd", api);
