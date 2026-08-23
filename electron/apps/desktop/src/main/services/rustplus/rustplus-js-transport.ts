/**
 * Transport implementation over @liamcottle/rustplus.js 2.5.0 — the exact version vendored in the
 * legacy runtime (RustPlusDesktop/runtime/rustplus-cli), answering audit open question §10.5 with
 * byte-compatible protobuf behavior.
 *
 * Library mechanics this wraps (verified against rustplus.js@2.5.0 source):
 *  - `connect()` is fire-and-forget: loads the proto file, then opens the WebSocket; outcomes arrive
 *    as events `connected` / `error`; `disconnected` fires on socket close.
 *  - Requests go through `sendRequestAsync(data, timeoutMs)` which auto-stamps seq/playerId/playerToken.
 *  - Smart-switch toggles and entity subscriptions are raw contract fields here (2.5.0 proto has no
 *    dedicated TurnSmartSwitch/AppPoke messages): setEntityValue / setSubscription.
 */
import { EventEmitter } from "node:events";
import type { RustTransport, ConnectOptions } from "./connection-core.js";
import type { Clock } from "./timing.js";
import { realClock } from "./timing.js";

/** Structural view of the untyped rustplus.js instance. */
export interface RustPlusInstance {
  connect(): void;
  disconnect(): void;
  isConnected(): boolean;
  sendRequestAsync(data: Record<string, unknown>, timeoutMilliseconds?: number): Promise<unknown>;
  on(event: string, listener: (...args: unknown[]) => void): unknown;
  removeListener(event: string, listener: (...args: unknown[]) => void): unknown;
  removeAllListeners?(event?: string): unknown;
}

export type RustPlusFactory = (
  server: string,
  port: number,
  playerId: string,
  playerToken: string,
  useFacepunchProxy: boolean,
) => RustPlusInstance;

/** Real factory used in production. */
export const realRustPlusFactory: RustPlusFactory = (server, port, playerId, playerToken, useFacepunchProxy) => {
  // eslint-disable-next-line @typescript-eslint/no-var-requires -- untyped CJS lib, default export is the class
  const RustPlus = require("@liamcottle/rustplus.js") as new (
    server: string,
    port: number,
    playerId: string,
    playerToken: string,
    useFacepunchProxy?: boolean,
  ) => RustPlusInstance;
  return new RustPlus(server, port, playerId, playerToken, useFacepunchProxy);
};

export type SocketLifecycle = "connected" | "disconnected";

/**
 * Lifecycle events re-emitted from the underlying instance after a successful connect():
 * "disconnected" (socket close), "socket-error", "message" (AppMessage not consumed by a callback).
 */
export class RustPlusJsTransport implements RustTransport {
  private instance: RustPlusInstance | null = null;
  readonly events = new EventEmitter();

  constructor(
    private readonly factory: RustPlusFactory,
    private readonly clock: Clock = realClock,
  ) {}

  /** The live instance for protocol requests — valid only while connected. */
  get current(): RustPlusInstance {
    if (!this.instance) throw new Error("transport not connected");
    return this.instance;
  }

  async connect(opts: ConnectOptions): Promise<void> {
    await this.disconnect(0);

    const inst = this.factory(opts.host, opts.port, opts.steamId64, opts.playerToken, opts.useProxy);

    await new Promise<void>((resolve, reject) => {
      let settled = false;
      const settleOk = (): void => {
        if (!settled) {
          settled = true;
          cleanup();
          resolve();
        }
      };
      const settleErr = (err: unknown): void => {
        if (!settled) {
          settled = true;
          cleanup();
          reject(err instanceof Error ? err : new Error(String(err)));
        }
      };
      const onConnected = (): void => settleOk();
      const onError = (e: unknown): void =>
        settleErr(e instanceof Error ? e : new Error(`websocket error: ${String(e)}`));
      const cleanup = (): void => {
        inst.removeListener("connected", onConnected);
        inst.removeListener("error", onError);
      };

      inst.on("connected", onConnected);
      inst.on("error", onError);

      // Start the socket BEFORE arming the timeout guard so an instantly-reported handshake outcome
      // always wins the microtask-order tie against the guard (deterministic under virtual clocks).
      inst.connect();
      void this.clock.sleep(opts.probeTimeoutMs).then(() => {
        if (!settled) settleErr(new Error(`the operation timed out after ${opts.probeTimeoutMs}ms (handshake)`));
      });
    });

    this.instance = inst;

    // Re-emit lifecycle so watchdog/layers can observe without touching the raw instance.
    inst.on("disconnected", () => this.events.emit("disconnected"));
    inst.on("error", (e: unknown) => this.events.emit("socket-error", e));
    inst.on("message", (m: unknown) => this.events.emit("message", m));
  }

  async disconnect(_timeoutMs?: number): Promise<void> {
    const inst = this.instance;
    this.instance = null;
    if (!inst) return;
    try {
      inst.removeAllListeners?.("message");
    } catch {
      /* structural optional */
    }
    inst.disconnect(); // terminate() under the hood — immediate, no graceful close needed
  }

  isConnected(): boolean {
    try {
      return this.instance?.isConnected() ?? false;
    } catch {
      return false;
    }
  }
}
