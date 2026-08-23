/**
 * Golden tests for the connectivity timing contracts (audit RUSTPLUS_CONNECTIVITY §5, §8):
 * rate limiter 50/25s⁻¹/333ms, 5-consecutive-timeout detector with immediate-loss classes,
 * 2 s×2→60 s backoff, dual-path proxy connect + per-connection state resets.
 */
import { describe, expect, it } from "vitest";
import {
  BackoffPolicy,
  RateLimiter,
  TimeoutDetector,
  classifyError,
  type Clock,
} from "../src/main/services/rustplus/timing.js";
import {
  ConnectionCore,
  ProxyExhaustedError,
  type RustTransport,
} from "../src/main/services/rustplus/connection-core.js";

/** Deterministic clock: sleep() advances virtual time instantly. */
function manualClock(start = 0): Clock & { advance: (ms: number) => void; now_: number } {
  const c = {
    now_: start,
    now() {
      return this.now_;
    },
    async sleep(ms: number) {
      this.now_ += ms;
    },
    advance(ms: number) {
      this.now_ += ms;
    },
  };
  return c as unknown as Clock & { advance: (ms: number) => void; now_: number };
}

describe("RateLimiter (legacy L52-94)", () => {
  it("allows a full burst of 50 immediately", async () => {
    const rl = new RateLimiter(manualClock());
    for (let i = 0; i < 50; i++) await rl.acquire();
    expect(rl.peek()).toBeCloseTo(0, 5);
  });

  it("the 51st acquire waits in 333 ms steps until refill", async () => {
    const clock = manualClock();
    const rl = new RateLimiter(clock);
    for (let i = 0; i < 50; i++) await rl.acquire();

    let resolved = false;
    const p = rl.acquire().then(() => {
      resolved = true;
    });
    await Promise.resolve(); // let the limiter enter its first wait
    expect(resolved).toBe(false);
    expect(clock.now_).toBe(333); // exactly one wait step so far

    await p;
    expect(resolved).toBe(true);
    // Refill over 333 ms at 25/s ≈ 8.25 tokens → consumed 1.
    expect(rl.peek()).toBeGreaterThan(7);
    expect(rl.peek()).toBeLessThan(8.3);
  });

  it("never exceeds the capacity of 50 no matter how long it idles", () => {
    const clock = manualClock(0);
    const rl = new RateLimiter(clock);
    clock.advance(1_000_000);
    expect(rl.peek()).toBe(50);
  });
});

describe("TimeoutDetector (legacy CheckConnectionLost L96-136)", () => {
  it("immediate-loss error classes fire ConnectionLost instantly and reset the counter", () => {
    const d = new TimeoutDetector();
    // Build up consecutive timeouts first — an immediate error must still fire instantly.
    for (let i = 0; i < 4; i++) expect(d.record(new Error("request timed out"))).toBe(false);

    expect(d.record(new Error("connection closed"))).toBe(true);
    expect(d.current).toBe(0); // counter reset on immediate loss
  });

  it("inner cause chain is walked like the legacy InnerException walk", () => {
    const d = new TimeoutDetector();
    const wrapped: Error & { cause?: unknown } = new Error("SendRequestAsync failed");
    wrapped.cause = new Error("Unable to read data from the transport connection");
    expect(d.record(wrapped)).toBe(true);
  });

  it("fires only after 5 CONSECUTIVE timeouts; a success resets", () => {
    const d = new TimeoutDetector();
    for (let i = 0; i < 4; i++) expect(d.record(new Error("the operation timed out"))).toBe(false);
    expect(d.record(new Error("task was canceled"))).toBe(true); // 5th consecutive

    const d2 = new TimeoutDetector();
    for (let i = 0; i < 3; i++) d2.record(new Error("timed out"));
    d2.success(); // valid response resets (L7470)
    expect(d2.current).toBe(0);
    // Counter restarts from zero: 4 more timeouts still survive, the 5th fires.
    for (let i = 0; i < 4; i++) expect(d2.record(new Error("timed out"))).toBe(false);
    expect(d2.record(new Error("timed out"))).toBe(true);
  });

  it("unrelated errors are 'other' and neither fire nor reset", () => {
    const d = new TimeoutDetector();
    d.record(new Error("some protobuf weirdness"));
    expect(d.current).toBe(0);
    expect(classifyError(new Error("some protobuf weirdness"))).toBe("other");
  });
});

describe("BackoffPolicy (legacy OnConnectionLost)", () => {
  it("yields 2 s ×2 capped at 60 s and resets", () => {
    const b = new BackoffPolicy();
    const seq: number[] = [];
    for (let i = 0; i < 7; i++) seq.push(b.next());
    expect(seq).toEqual([2000, 4000, 8000, 16000, 32000, 60000, 60000]);
    b.reset();
    expect(b.peek).toBe(2000);
  });
});

describe("ConnectionCore (legacy ConnectAsync L5860-5944)", () => {
  const profile = { host: "1.2.3.4", port: 28082, steamId64: "765", playerToken: "tok", UseFacepunchProxy: false };

  it("preferred path succeeds → connected with proxy choice recorded", async () => {
    const attempts: boolean[] = [];
    const transport: RustTransport = {
      async connect(opts) {
        attempts.push(opts.useProxy);
      },
      async disconnect() {},
    };
    const core = new ConnectionCore(transport, manualClock());
    await core.connect(profile);
    expect(core.isConnected).toBe(true);
    expect(core.state.activeProxy).toBe("direct");
    expect(attempts).toEqual([false]);
  });

  it("falls back to the opposite path when the preferred one fails", async () => {
    const attempts: boolean[] = [];
    const transport: RustTransport = {
      async connect(opts) {
        attempts.push(opts.useProxy);
        if (opts.useProxy) throw new Error("proxy unreachable"); // preferred path fails
      },
      async disconnect() {},
    };
    const core = new ConnectionCore(transport, manualClock());
    await core.connect({ ...profile, UseFacepunchProxy: true }); // proxy preferred
    expect(attempts).toEqual([true, false]); // proxy tried first, then direct
    expect(core.state.activeProxy).toBe("direct");
  });

  it("both paths failing → legacy German error message", async () => {
    const transport: RustTransport = {
      async connect() {
        throw new Error("nope");
      },
      async disconnect() {},
    };
    const core = new ConnectionCore(transport, manualClock());
    await expect(core.connect(profile)).rejects.toThrowError(/Rust\+ nicht erreichbar \(direkt & Proxy\)/);
    await expect(core.connect(profile)).rejects.toBeInstanceOf(ProxyExhaustedError);
  });

  it("every connect begins with full teardown of the previous socket", async () => {
    const calls: string[] = [];
    const transport: RustTransport = {
      async connect() {
        calls.push("connect");
      },
      async disconnect() {
        calls.push("disconnect");
      },
    };
    const core = new ConnectionCore(transport, manualClock());
    await core.connect(profile);
    core.markSubscribed(42);
    await core.connect(profile);
    expect(calls).toEqual(["connect", "disconnect", "connect"]); // forced clean slate
    expect(core.needsSubscribeOnce(42)).toBe(true); // subscription state is per-connection
  });

  it("disconnect clears all per-connection state even when the socket hangs (2 s cap)", async () => {
    const clock = manualClock();
    const transport: RustTransport = {
      async connect() {},
      disconnect(): Promise<void> {
        return new Promise(() => undefined); // hung socket, never settles
      },
    };
    const core = new ConnectionCore(transport, clock);
    await core.connect(profile);
    core.markSubscribed(7);
    await core.disconnect();
    // Virtual time accumulated both guards: 7 s connect cap (guard still elapsed) + 2 s disconnect cap.
    expect(clock.now_).toBe(9000);
    expect(core.isConnected).toBe(false);
    expect(core.state.subOnce.size).toBe(0);
    expect(core.state.teamChatPrimed).toBe(false);
    expect(core.state.activeProxy).toBeNull();
  });

  it("hard caps a hanging connect attempt at 7 s", async () => {
    const clock = manualClock();
    const transport: RustTransport = {
      connect(): Promise<void> {
        return new Promise(() => undefined); // never settles
      },
      async disconnect() {},
    };
    const core = new ConnectionCore(transport, clock);
    await expect(core.connect(profile)).rejects.toThrowError(/direkt & Proxy/);
    expect(clock.now_).toBeGreaterThanOrEqual(7000); // both attempts hit the hard cap
  });
});
