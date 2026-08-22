/**
 * Typed IPC handler registry (main side).
 *
 * Every channel declared in @rpd/shared gets exactly one zod-validated handler here. Uses the native
 * invoke/handle promise protocol: handlers return plain values; this registry wraps results into an
 * {ok:true,data}/{ok:false,error} envelope so contract breaches and failures never leak unshaped errors
 * across the bridge.
 */
import type {
  ChannelDef,
  IpcChannels,
  IpcError,
} from "@rpd/shared";
import type { z } from "zod";
import { logger } from "./logger.js";
import { errText } from "./util.js";

export type HandlerMapOf<TChannels extends Record<string, ChannelDef<z.ZodTypeAny, z.ZodTypeAny>>> = {
  [K in keyof TChannels & string]: (
    payload: z.infer<TChannels[K]["request"]>,
  ) => z.infer<TChannels[K]["response"]> | Promise<z.infer<TChannels[K]["response"]>>;
};

/** Minimal structural surface of Electron's ipcMain — enables unit testing without the Electron runtime. */
export interface IpcRegistrar {
  handle(channel: string, listener: (raw: unknown) => Promise<unknown>): void;
}

export type InvokeEnvelope = { ok: true; data: unknown } | { ok: false; error: IpcError };

/**
 * Registers validated handlers for every declared channel. Throws at startup if a declared channel is
 * missing a handler (wiring bug) — parity principle: fail loudly, never silently degrade.
 */
export function registerIpcHandlers<
  TChannels extends Record<string, ChannelDef<z.ZodTypeAny, z.ZodTypeAny>>,
>(
  registrar: IpcRegistrar,
  channels: TChannels,
  handlers: HandlerMapOf<TChannels>,
): Record<string, (raw: unknown) => Promise<InvokeEnvelope>> {
  const registered: Record<string, (raw: unknown) => Promise<InvokeEnvelope>> = {};

  for (const [name, def] of Object.entries(channels)) {
    const handler = handlers[name as keyof TChannels & string] as
      | ((payload: unknown) => unknown)
      | undefined;
    if (!handler) {
      throw new Error(`IPC channel "${name}" is declared but has no registered handler`);
    }

    const wrapped = async (raw: unknown): Promise<InvokeEnvelope> => {
      const parsedRequest = def.request.safeParse(raw);
      if (!parsedRequest.success) {
        const message = `request failed schema validation: ${firstIssue(parsedRequest.error)}`;
        logger.warn("ipc", `bad_request on ${name}: ${message}`);
        return { ok: false, error: { code: "bad_request", message } };
      }
      try {
        const result = await handler(parsedRequest.data);
        const parsedResponse = def.response.safeParse(result);
        if (!parsedResponse.success) {
          logger.error("ipc", `response breached contract on ${name}: ${firstIssue(parsedResponse.error)}`);
          return { ok: false, error: { code: "handler_error", message: "response failed schema validation" } };
        }
        return { ok: true, data: parsedResponse.data };
      } catch (err) {
        logger.error("ipc", `handler_error on ${name}: ${errText(err)}`);
        return { ok: false, error: { code: "handler_error", message: errText(err).slice(0, 2000) } };
      }
    };

    registered[name] = wrapped;
    registrar.handle(name, wrapped);
  }

  return registered;
}

function firstIssue(error: z.ZodError): string {
  return error.issues[0]?.message ?? "unknown issue";
}

/** Wires the shared channel registry onto Electron's real ipcMain. */
export function createRegistrar(channels: IpcChannels): {
  register(handlers: HandlerMapOf<IpcChannels>): void;
} {
  let registered = false;
  return {
    register(handlers: HandlerMapOf<IpcChannels>): void {
      if (registered) throw new Error("IPC handlers already registered");
      registered = true;
      // Lazy electron import keeps this module unit-testable without the runtime.
      void import("electron").then(({ ipcMain }) => {
        registerIpcHandlers(
          {
            handle: (channel, listener) => {
              ipcMain.handle(channel, (_event, raw: unknown) => listener(raw));
            },
          },
          channels,
          handlers,
        );
      });
    },
  };
}
