/**
 * Protocol request contracts over a connected rustplus.js instance (2.5.0 proto).
 * Raw AppRequest fields — the audit's "raw contract" tier of the legacy 3-path cascade, minus the
 * reflection gymnastics that existed only to survive HandyS11 beta drift.
 */
import type { RustPlusInstance } from "./rustplus-js-transport.js";

/** AppError enum: protobufjs decodes 0 (= null) for success. */
export function responseError(message: unknown): string | null {
  const err = (message as { response?: { error?: unknown } } | null)?.response?.error;
  if (err === undefined || err === null || err === 0 || err === "null" || err === "") return null;
  return typeof err === "string" ? err : String(err);
}

export class ProtocolError extends Error {
  constructor(
    readonly code: string,
    readonly request: Record<string, unknown>,
  ) {
    super(`Rust+ request failed (${code}): ${Object.keys(request).join("/")}`);
    this.name = "ProtocolError";
  }
}

const DEFAULT_TIMEOUT_MS = 10_000;

/** Send a request; resolve with response payload or throw on error/timeout/missing response. */
export async function request(
  instance: Pick<RustPlusInstance, "sendRequestAsync">,
  data: Record<string, unknown>,
  timeoutMs = DEFAULT_TIMEOUT_MS,
): Promise<Record<string, unknown>> {
  const message = (await instance.sendRequestAsync(data, timeoutMs)) as {
    response?: Record<string, unknown>;
  } | null;
  if (!message || !message.response) throw new Error(`Rust+ request returned no response: ${Object.keys(data).join("/")}`);
  const code = responseError(message);
  if (code !== null) throw new ProtocolError(code, data);
  return message.response;
}

/** Request builders — exact proto field names from rustplus.proto@2.5.0. */
export const rq = {
  getInfo: (): Record<string, unknown> => ({ getInfo: {} }),
  getMap: (): Record<string, unknown> => ({ getMap: {} }),
  getTime: (): Record<string, unknown> => ({ getTime: {} }),
  getTeamInfo: (): Record<string, unknown> => ({ getTeamInfo: {} }),
  getTeamChat: (): Record<string, unknown> => ({ getTeamChat: {} }),
  getMapMarkers: (): Record<string, unknown> => ({ getMapMarkers: {} }),
  sendTeamMessage: (message: string): Record<string, unknown> => ({ sendTeamMessage: { message } }),
  promoteToLeader: (steamId: string): Record<string, unknown> => ({ promoteToLeader: { steamId } }),
  subscribeEntity: (entityId: number): Record<string, unknown> => ({
    entityId,
    setSubscription: { value: true },
  }),
  unsubscribeEntity: (entityId: number): Record<string, unknown> => ({
    entityId,
    setSubscription: { value: false },
  }),
  setEntityValue: (entityId: number, value: boolean): Record<string, unknown> => ({
    entityId,
    setEntityValue: { entityId, value },
  }),
  getEntityInfo: (entityId: number): Record<string, unknown> => ({
    entityId,
    getEntityInfo: {},
  }),
  checkSubscription: (entityId: number): Record<string, unknown> => ({
    entityId,
    checkSubscription: {},
  }),
} as const;

/** Convenience wrappers used by device/chat layers. */
export interface ProtocolApi {
  raw: Pick<RustPlusInstance, "sendRequestAsync">;
  send(data: Record<string, unknown>, timeoutMs?: number): Promise<Record<string, unknown>>;
  toggleSwitch(entityId: number, on: boolean): Promise<Record<string, unknown>>;
  switchState(entityId: number): Promise<boolean | null>;
}

export function makeProtocol(instance: Pick<RustPlusInstance, "sendRequestAsync">): ProtocolApi {
  return {
    raw: instance,
    send: (data, timeoutMs) => request(instance, data, timeoutMs),
    toggleSwitch: (entityId, on) => request(instance, rq.setEntityValue(entityId, on)),
    switchState: async (entityId) => {
      const res = await request(instance, rq.getEntityInfo(entityId));
      // AppEntityInfo.checkSubscription path carries entityInfo.payload.value for switches.
      const payload = (
        res["entityInfo"] as
          | { payload?: { value?: boolean | number | string } }
          | undefined
      )?.payload;
      if (!payload || payload.value === undefined) return null;
      return Boolean(payload.value);
    },
  };
}
