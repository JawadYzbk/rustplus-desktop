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
import { ipcChannels } from "@rpd/shared";

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

  /** Persisted shell preferences (sidebar). */
  getUiPrefs: () => invoke("uiPrefs/get", undefined) as Promise<InvokeResult<unknown>>,
  setUiPrefs: (patch: { sidebarPinned?: boolean; sidebarWidth?: number }) =>
    invoke("uiPrefs/set", patch) as Promise<InvokeResult<unknown>>,
};

export type RpdApi = typeof api;

contextBridge.exposeInMainWorld("rpd", api);
