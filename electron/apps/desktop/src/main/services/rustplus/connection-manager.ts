/**
 * ConnectionManager — composes the connectivity primitives behind one lifecycle, porting the
 * observable behavior of RustPlusClientReal + the reconnect semantics of OnConnectionLost:
 *  - every contract send passes the rate limiter (50/25s⁻¹/333ms);
 *  - responses feed the 5-consecutive-timeout detector; immediate-loss errors fire instantly;
 *  - connection lost → guarded exponential-backoff silent reconnect (2 s ×2 → 60 s);
 *  - 5 consecutive failed status polls (fed by the feature-layer polling loops) → silent refresh;
 *  - chats are pull-to-primed once per connection right after connect.
 */
import { EventEmitter } from "node:events";
import { ConnectionCore, ProxyExhaustedError, type ConnectionProfileRef } from "./connection-core.js";
import type { RustPlusInstance } from "./rustplus-js-transport.js";
import { BackoffPolicy, RateLimiter, TimeoutDetector, realClock, type Clock } from "./timing.js";
import { makeProtocol, rq, request, type ProtocolApi } from "./protocol.js";
import { ChatPrimer } from "./watchdog.js";

/** Structural surface of the transport the manager drives (rustplus.js-backed). */
export interface ManagedTransport {
  connect(opts: {
    host: string;
    port: number;
    steamId64: string;
    playerToken: string;
    useProxy: boolean;
    probeTimeoutMs: number;
  }): Promise<void>;
  disconnect(timeoutMs?: number): Promise<void>;
  /** Live protocol endpoint; meaningful only between successful connect and disconnect. */
  readonly current: Pick<RustPlusInstance, "sendRequestAsync">;
}

export type ManagerEvent =
  | { kind: "connecting"; proxyPreferred: boolean }
  | { kind: "connected"; proxy: "direct" | "proxy" }
  | { kind: "lost"; reason: string }
  | { kind: "reconnectingIn"; delayMs: number }
  | { kind: "silentRefreshStarted" }
  | { kind: "disconnected" };

export interface ConnSnapshot {
  connected: boolean;
  activeProxy: "direct" | "proxy" | null;
  host: string | null;
  port: number | null;
  consecutiveTimeouts: number;
  teamChatPrimed: boolean;
  clanChatPrimed: boolean;
}

const DEFAULT_REQUEST_TIMEOUT_MS = 10_000;

export class ConnectionManager extends EventEmitter {
  private readonly core: ConnectionCore;
  private readonly limiter: RateLimiter;
  private readonly detector: TimeoutDetector;
  private readonly backoff: BackoffPolicy;
  private readonly chat: ChatPrimer;

  private profile: ConnectionProfileRef | null = null;
  private protocol: ProtocolApi | null = null;
  private reconnecting = false;
  private statusFailures = 0;

  constructor(
    private readonly transport: ManagedTransport,
    private readonly clock: Clock = realClock,
  ) {
    super();
    this.core = new ConnectionCore(transport, clock);
    this.limiter = new RateLimiter(clock);
    this.detector = new TimeoutDetector();
    this.backoff = new BackoffPolicy();
    // Provider form: the live endpoint changes on every (re)connect.
    this.chat = new ChatPrimer(() => {
      if (!this.protocol) throw new Error("not connected");
      return this.protocol;
    });
  }

  get isConnected(): boolean {
    return this.core.isConnected;
  }

  snapshot(): ConnSnapshot {
    return {
      connected: this.core.isConnected,
      activeProxy: this.core.state.activeProxy,
      host: this.profile?.host ?? null,
      port: this.profile?.port ?? null,
      consecutiveTimeouts: this.detector.current,
      teamChatPrimed: this.core.state.teamChatPrimed,
      clanChatPrimed: this.core.state.clanChatPrimed,
    };
  }

  /** Full connect: dual-path probe + per-connection resets + chat priming. */
  async connect(profile: ConnectionProfileRef): Promise<ConnSnapshot> {
    this.cancelReconnect();
    this.profile = profile;
    this.emit("connecting", { proxyPreferred: profile.UseFacepunchProxy === true });
    try {
      const state = await this.core.connect(profile);
      this.protocol = makeProtocol(this.transport.current);
      await this.primeChats();
      this.emit("connected", { proxy: state.activeProxy! });
      return this.snapshot();
    } catch (err) {
      if (err instanceof ProxyExhaustedError) this.emit("lost", { reason: err.message });
      throw err;
    }
  }

  /** Rate-limited contract send with timeout-detector accounting. */
  async send(
    data: Record<string, unknown>,
    timeoutMs = DEFAULT_REQUEST_TIMEOUT_MS,
  ): Promise<Record<string, unknown>> {
    if (!this.core.isConnected || !this.protocol) throw new Error("not connected");
    await this.limiter.acquire();
    try {
      const res = await request(this.protocol.raw, data, timeoutMs);
      this.detector.success(); // valid response resets the counter (legacy L7470)
      return res;
    } catch (err) {
      if (this.detector.record(err)) this.handleLost(String(err instanceof Error ? err.message : err));
      throw err;
    }
  }

  /** Toggle a smart switch through the raw contract (rate-limited like every send). */
  async setEntityValue(entityId: number, value: boolean): Promise<void> {
    await this.send(rq.setEntityValue(entityId, value));
  }

  /** Status-poll bookkeeping; 5 consecutive failures trigger a silent refresh. */
  recordStatusResult(ok: boolean): void {
    if (!ok) {
      this.statusFailures += 1;
      if (this.statusFailures >= 5) {
        this.statusFailures = 0;
        this.handleLost("5 consecutive failed status polls");
      }
    } else {
      this.statusFailures = 0;
    }
  }

  async disconnect(): Promise<ConnSnapshot> {
    this.cancelReconnect();
    this.backoff.reset();
    this.statusFailures = 0;
    this.detector.success();
    await this.core.disconnect();
    this.protocol = null;
    this.emit("disconnected");
    return this.snapshot();
  }

  private async primeChats(): Promise<void> {
    try {
      await this.chat.primeTeamChat(this.core.state);
    } catch {
      /* best effort */
    }
    try {
      await this.chat.primeClanChat(this.core.state);
    } catch {
      /* best effort */
    }
  }

  /** Connection-lost path: guarded exponential-backoff silent reconnect. */
  private handleLost(reason: string): void {
    if (!this.profile || this.reconnecting) return;
    this.reconnecting = true;
    this.emit("lost", { reason });
    void this.reconnectLoop();
  }

  private async reconnectLoop(): Promise<void> {
    while (this.reconnecting && this.profile) {
      const delayMs = this.backoff.next();
      this.emit("reconnectingIn", { delayMs });
      await this.clock.sleep(delayMs);
      if (!this.reconnecting) return; // disconnect() cancelled while we slept
      this.emit("silentRefreshStarted");
      try {
        await this.core.connect(this.profile);
        this.protocol = makeProtocol(this.transport.current);
        await this.primeChats();
        this.backoff.reset();
        this.detector.success();
        this.reconnecting = false;
        this.emit("connected", { proxy: this.core.state.activeProxy! });
        return;
      } catch {
        continue; // next backoff step
      }
    }
  }

  private cancelReconnect(): void {
    this.reconnecting = false;
  }
}
