/**
 * Connection core — connect/disconnect lifecycle semantics ported from RustPlusClientReal.ConnectAsync
 * (L5860-5944) + DisconnectAsync (L5946-5991):
 *  - every connect begins with full teardown of the previous socket;
 *  - dual-path proxy: preferred path first (profile.UseFacepunchProxy), automatic fallback to the
 *    opposite, error "Rust+ nicht erreichbar (direkt & Proxy)" when both fail;
 *  - per-connection subscription/chat state cleared on connect AND disconnect;
 *  - attempt budget: probe raced against 7 s delay with 6 s CTS → modeled as a 6 s probe timeout
 *    inside the transport plus a 7 s hard cap here.
 *
 * The transport is injected; the real @liamcottle/rustplus.js transport lands next round.
 */
import type { Clock } from "./timing.js";
import { realClock } from "./timing.js";

export interface RustTransport {
  /**
   * Establish the socket AND complete an initial info probe. Throws on failure.
   * `useProxy` selects Facepunch proxy vs direct connection.
   */
  connect(opts: ConnectOptions): Promise<void>;
  disconnect(timeoutMs?: number): Promise<void>;
}

export interface ConnectOptions {
  host: string;
  port: number;
  steamId64: string;
  playerToken: string;
  useProxy: boolean;
  /** Legacy: GetInfoAsync raced against 7 s delay with 6 s CTS. */
  probeTimeoutMs: number;
}

export interface ConnectionProfileRef {
  host: string;
  port: number;
  steamId64: string;
  playerToken: string;
  UseFacepunchProxy?: boolean;
}

/** Per-connection subscription state — audit §8: cleared on connect AND disconnect. */
export interface ConnectionState {
  connected: boolean;
  activeProxy: "direct" | "proxy" | null;
  subOnce: Set<number>;
  subscribed: Set<number>;
  teamChatPrimed: boolean;
  clanChatPrimed: boolean;
}

const PROBE_TIMEOUT_MS = 6_000;
const HARD_CAP_MS = 7_000;
const DISCONNECT_CAP_MS = 2_000;

export class ConnectionCore {
  readonly state: ConnectionState = freshState();
  private sessionActive = false;

  constructor(
    private readonly transport: RustTransport,
    private readonly clock: Clock = realClock,
  ) {}

  get isConnected(): boolean {
    return this.state.connected;
  }

  async connect(profile: ConnectionProfileRef): Promise<ConnectionState> {
    // 1. Forced clean slate ("prevent overlapping socket leaks"). Legacy DisconnectAsync no-ops when
    // no API instance exists — mirrored via sessionActive.
    await this.disconnect();

    const preferred = profile.UseFacepunchProxy === true ? "proxy" : "direct";

    for (const choice of [preferred, preferred === "proxy" ? "direct" : "proxy"] as const) {
      try {
        await withTimeout(
          () =>
            this.transport.connect({
              host: profile.host,
              port: profile.port,
              steamId64: profile.steamId64,
              playerToken: profile.playerToken,
              useProxy: choice === "proxy",
              probeTimeoutMs: PROBE_TIMEOUT_MS,
            }),
          HARD_CAP_MS,
          this.clock,
        );
        this.state.connected = true;
        this.state.activeProxy = choice;
        this.sessionActive = true;
        return this.state;
      } catch (err) {
        if (choice !== preferred) {
          throw new ProxyExhaustedError(err);
        }
        // else: fall through to opposite path
      }
    }
    throw new ProxyExhaustedError(new Error("unreachable")); // defensive; loop always throws inside
  }

  async disconnect(): Promise<void> {
    if (this.sessionActive) {
      try {
        await withTimeout(() => this.transport.disconnect(DISCONNECT_CAP_MS), DISCONNECT_CAP_MS, this.clock);
      } catch {
        // Graceful close capped at 2 s — dispose regardless (legacy parity).
      }
      this.sessionActive = false;
    }
    Object.assign(this.state, freshState());
  }

  /** Subscription bookkeeping helpers used by higher layers. */
  markSubscribed(entityId: number): void {
    this.state.subOnce.add(entityId);
    this.state.subscribed.add(entityId);
  }

  needsSubscribeOnce(entityId: number): boolean {
    return !this.state.subOnce.has(entityId);
  }
}

export class ProxyExhaustedError extends Error {
  constructor(readonly cause_: unknown) {
    super("Rust+ nicht erreichbar (direkt & Proxy)");
    this.name = "ProxyExhaustedError";
  }
}

function freshState(): ConnectionState {
  return {
    connected: false,
    activeProxy: null,
    subOnce: new Set(),
    subscribed: new Set(),
    teamChatPrimed: false,
    clanChatPrimed: false,
  };
}

/**
 * Hard deadline around `start()`. Work is registered FIRST so an instantly-settling result wins the
 * tie against the guard (deterministic under virtual clocks); a never-settling operation loses to
 * the guard, which is exactly the legacy CTS/race behavior.
 */
function withTimeout<T>(start: () => Promise<T>, ms: number, clock: Clock): Promise<T> {
  return new Promise<T>((resolve, reject) => {
    let settled = false;
    const settleErr = (e: unknown): void => {
      if (!settled) {
        settled = true;
        reject(e);
      }
    };
    const settleOk = (v: T): void => {
      if (!settled) {
        settled = true;
        resolve(v);
      }
    };
    void start().then(settleOk, settleErr);
    void clock.sleep(ms).then(() => settleErr(new Error(`operation timed out after ${ms}ms`)));
  });
}
