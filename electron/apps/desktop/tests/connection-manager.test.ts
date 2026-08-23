/**
 * ConnectionManager lifecycle tests — rate-limited sends, timeout-detector-driven reconnect,
 * backoff sequence, chat priming per connection, watchdog-triggered silent refresh.
 */
import { describe, expect, it } from "vitest";
import {
  ConnectionManager,
  type ManagedTransport,
} from "../src/main/services/rustplus/connection-manager.js";
import { ProxyExhaustedError } from "../src/main/services/rustplus/connection-core.js";
import type { Clock } from "../src/main/services/rustplus/timing.js";
import type { RustPlusInstance } from "../src/main/services/rustplus/rustplus-js-transport.js";

/** Lets every queued microtask (a full reconnect-loop pass) settle before assertions. */
const drain = async (): Promise<void> => {
  for (let i = 0; i < 3; i++) await new Promise<void>((r) => setImmediate(r));
};

function manualClock(): Clock & { now_: number } {
  const c = {
    now_: 0,
    now() {
      return this.now_;
    },
    async sleep(ms: number) {
      this.now_ += ms;
    },
  };
  return c as unknown as Clock & { now_: number };
}

/** Scripted endpoint: responses by request-type key; failures configurable. */
class ScriptedEndpoint implements Pick<RustPlusInstance, "sendRequestAsync"> {
  sent: string[] = [];
  failNext = 0;
  failWith: (data: Record<string, unknown>) => Error = () => new Error("request timed out");

  async sendRequestAsync(data: Record<string, unknown>): Promise<unknown> {
    const key = Object.keys(data).find((k) => k !== "entityId") ?? "unknown";
    this.sent.push(key);
    if (this.failNext > 0) {
      this.failNext -= 1;
      throw this.failWith(data);
    }
    return { response: { error: 0 } };
  }
}

/** Transport whose connect() can be scripted to fail/succeed per attempt. */
class ScriptedTransport implements ManagedTransport {
  readonly endpoint = new ScriptedEndpoint();
  attempts: boolean[] = []; // useProxy flag per attempt
  /** Outcomes consumed per connect(): true=resolve, false=reject. */
  outcomes: Array<"ok" | "fail"> = [];
  hangs: number[] = []; // indexes of connect() calls that never settle
  private n = -1;

  async connect(opts: { useProxy: boolean }): Promise<void> {
    this.n += 1;
    this.attempts.push(opts.useProxy);
    if (this.hangs.includes(this.n)) return new Promise(() => undefined);
    const outcome = this.outcomes.shift() ?? "ok";
    if (outcome === "fail") throw new Error("connect failed");
  }

  async disconnect(): Promise<void> {}

  get current() {
    return this.endpoint;
  }
}

const profile = { host: "1.2.3.4", port: 28082, steamId64: "765", playerToken: "tok", UseFacepunchProxy: false };

function events(mgr: ConnectionManager): Array<{ kind: string; [k: string]: unknown }> {
  const log: Array<{ kind: string; [k: string]: unknown }> = [];
  mgr.on("connecting", () => log.push({ kind: "connecting" }));
  mgr.on("connected", (e) => log.push({ kind: "connected", ...e }));
  mgr.on("lost", (e) => log.push({ kind: "lost", ...e }));
  mgr.on("reconnectingIn", (e) => log.push({ kind: "reconnectingIn", ...e }));
  mgr.on("silentRefreshStarted", () => log.push({ kind: "silentRefreshStarted" }));
  mgr.on("disconnected", () => log.push({ kind: "disconnected" }));
  return log;
}

describe("ConnectionManager.connect", () => {
  it("connects, primes team+clan chat once each, snapshot reflects state", async () => {
    const t = new ScriptedTransport();
    const mgr = new ConnectionManager(t, manualClock());
    const snap = await mgr.connect(profile);

    expect(snap.connected).toBe(true);
    expect(snap.activeProxy).toBe("direct");
    expect(snap.teamChatPrimed).toBe(true);
    expect(snap.clanChatPrimed).toBe(true);
    expect(t.endpoint.sent.filter((k) => k === "getTeamChat")).toHaveLength(2); // team + clan prime

    // A fresh connect() is a NEW connection → legacy re-primes per connection.
    await mgr.connect(profile);
    expect(t.endpoint.sent.filter((k) => k === "getTeamChat")).toHaveLength(4);
  });

  it("both proxy paths failing surfaces the legacy German error and a lost event", async () => {
    const t = new ScriptedTransport();
    t.outcomes = ["fail", "fail"];
    const mgr = new ConnectionManager(t, manualClock());
    const log = events(mgr);
    await expect(mgr.connect(profile)).rejects.toThrowError(ProxyExhaustedError);
    expect(log.some((e) => e.kind === "lost")).toBe(true);
    expect(mgr.isConnected).toBe(false);
  });
});

describe("rate limiting + detector accounting in send()", () => {
  it("successful sends reset the consecutive-timeout counter", async () => {
    const t = new ScriptedTransport();
    const mgr = new ConnectionManager(t, manualClock());
    await mgr.connect(profile);
    t.endpoint.failNext = 4;
    for (let i = 0; i < 4; i++) await mgr.send({ getInfo: {} }).catch(() => undefined);
    expect(mgr.snapshot().consecutiveTimeouts).toBe(4);
    await mgr.send({ getInfo: {} }); // success resets
    expect(mgr.snapshot().consecutiveTimeouts).toBe(0);
  });

  it("5 consecutive timeouts declare loss and schedule exponential-backoff reconnects", async () => {
    const clock = manualClock();
    const t = new ScriptedTransport();
    const mgr = new ConnectionManager(t, clock);
    const log = events(mgr);
    await mgr.connect(profile);

    // Every subsequent send times out. Reconnect iter-1 burns BOTH proxy paths (direct fail +
    // proxy fail → ProxyExhaustedError → continue); iter-2 succeeds on the direct path.
    t.outcomes = ["fail", "fail"];
    t.endpoint.failNext = 5;
    t.endpoint.failWith = () => new Error("the operation timed out");
    for (let i = 0; i < 5; i++) await mgr.send({ getTime: {} }).catch(() => undefined);

    const lostIdx = log.findIndex((e) => e.kind === "lost");
    expect(lostIdx).toBeGreaterThanOrEqual(0);
    expect(log[lostIdx + 1]).toMatchObject({ kind: "reconnectingIn", delayMs: 2000 });

    // Virtual time advances instantly through sleeps — reconnect loop runs to completion.
    await drain();
    expect(log.some((e) => e.kind === "silentRefreshStarted")).toBe(true);
    // First silent refresh failed (outcomes exhausted) → second step at 4000 ms.
    expect(
      log.some((e) => e.kind === "reconnectingIn" && e.delayMs === 4000),
    ).toBe(true);
    expect(mgr.isConnected).toBe(true); // eventually reconnected
    expect(mgr.snapshot().teamChatPrimed).toBe(true); // re-primed on the fresh connection
  });

  it("immediate-loss errors fire instantly regardless of counter state", async () => {
    const clock = manualClock();
    const t = new ScriptedTransport();
    const mgr = new ConnectionManager(t, clock);
    const log = events(mgr);
    await mgr.connect(profile);
    t.outcomes = ["ok"]; // reconnect succeeds immediately
    t.endpoint.failNext = 1;
    t.endpoint.failWith = () => new Error("connection closed");
    await mgr.send({ getInfo: {} }).catch(() => undefined);
    await drain();
    expect(log.some((e) => e.kind === "lost")).toBe(true);
    expect(mgr.isConnected).toBe(true);
  });

  it("disconnect() cancels a pending reconnect loop", async () => {
    const clock = manualClock();
    const t = new ScriptedTransport();
    const mgr = new ConnectionManager(t, clock);
    await mgr.connect(profile);
    t.outcomes = ["fail", "fail", "fail"];
    t.endpoint.failNext = 5;
    t.endpoint.failWith = () => new Error("timed out");
    for (let i = 0; i < 5; i++) await mgr.send({ getTime: {} }).catch(() => undefined);
    await mgr.disconnect(); // aborts mid-loop
    const before = t.attempts.length;
    clock.sleep(10_000); // let any stray timers run virtually
    await Promise.resolve();
    expect(t.attempts.length).toBeLessThanOrEqual(before + 1);
    expect(mgr.isConnected).toBe(false);
  });
});

describe("watchdog via recordStatusResult", () => {
  it("5 consecutive status poll failures trigger one silent refresh", async () => {
    const clock = manualClock();
    const t = new ScriptedTransport();
    const mgr = new ConnectionManager(t, clock);
    const log = events(mgr);
    await mgr.connect(profile);
    t.outcomes = ["ok"];

    for (let i = 0; i < 5; i++) mgr.recordStatusResult(false);
    await drain();
    expect(log.filter((e) => e.kind === "lost")).toHaveLength(1);
    expect(mgr.isConnected).toBe(true); // refreshed right away (single ok outcome)

    // Successes keep the failure counter at zero afterwards.
    for (let i = 0; i < 4; i++) mgr.recordStatusResult(false);
    mgr.recordStatusResult(true);
    mgr.recordStatusResult(false);
    expect(log.filter((e) => e.kind === "lost")).toHaveLength(1);
  });
});
