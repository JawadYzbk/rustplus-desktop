/** Typed access to the preload bridge. The only sanctioned way the renderer talks to the main process. */
import type { IpcChannelName } from "@rpd/shared";

type RpdBridge = {
  invoke(channel: IpcChannelName | string, payload: unknown): Promise<InvokeResult<unknown>>;
  getInfo(): Promise<InvokeResult<unknown>>;
  log(level: "debug" | "info" | "warn" | "error", scope: string, message: string): Promise<InvokeResult<unknown>>;
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
