/**
 * Transport + protocol + subscription tests against a fake rustplus.js instance (no live server).
 * Pins: handshake event/timeout semantics, response error mapping, raw request contracts from the
 * 2.5.0 proto, subscribe-once-per-connection priming with 100 ms gaps.
 */
import { describe, expect, it } from "vitest";
import { EventEmitter } from "node:events";
import {
  RustPlusJsTransport,
  type RustPlusInstance,
} from "../src/main/services/rustplus/rustplus-js-transport.js";
import { ProtocolError, makeProtocol, rq, request, responseError } from "../src/main/services/rustplus/protocol.js";
import { SubscriptionOrchestrator, PRIME_GAP_MS } from "../src/main/services/rustplus/subscriptions.js";
import type { Clock } from "../src/main/services/rustplus/timing.js";

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

/** Fake of rustplus.js@2.5.0 surface: EventEmitter lifecycle + scripted sendRequestAsync. */
class FakeRustPlus extends EventEmitter implements RustPlusInstance {
  connected = false;
  sent: Array<Record<string, unknown>> = [];
  /** Scripted responder; default: success with empty response. */
  responder?: (data: Record<string, unknown>) => Promise<unknown>;
  hangHandshake = false;

  connect(): void {
    if (this.hangHandshake) return; // never emits
    queueMicrotask(() => {
      this.connected = true;
      this.emit("connected");
    });
  }

  disconnect(): void {
    if (!this.connected) return;
    this.connected = false;
    this.emit("disconnected");
  }

  isConnected(): boolean {
    return this.connected;
  }

  async sendRequestAsync(data: Record<string, unknown>): Promise<unknown> {
    this.sent.push(data);
    if (this.responder) return this.responder(data);
    // Default success envelope mirrors protobufjs decode shape.
    return { response: { seq: 1, error: 0 } };
  }
}

const factoryOf = (inst: FakeRustPlus) => () => inst;

describe("RustPlusJsTransport", () => {
  const opts = { host: "1.2.3.4", port: 28082, steamId64: "765", playerToken: "tok", useProxy: false, probeTimeoutMs: 6000 };

  it("resolves once the socket reports connected and keeps the instance for requests", async () => {
    const inst = new FakeRustPlus();
    const t = new RustPlusJsTransport(factoryOf(inst), manualClock());
    await t.connect(opts);
    expect(t.isConnected()).toBe(true);
    await expect(t.current.sendRequestAsync(rq.getInfo())).resolves.toBeDefined();
    await t.disconnect();
    expect(t.isConnected()).toBe(false);
  });

  it("rejects when the websocket errors during handshake", async () => {
    const inst = new FakeRustPlus();
    inst.connect = () => queueMicrotask(() => inst.emit("error", new Error("ECONNREFUSED")));
    const t = new RustPlusJsTransport(factoryOf(inst), manualClock());
    await expect(t.connect(opts)).rejects.toThrowError(/ECONNREFUSED/);
  });

  it("rejects on handshake timeout when the instance never emits", async () => {
    const clock = manualClock();
    const inst = new FakeRustPlus();
    inst.hangHandshake = true;
    const t = new RustPlusJsTransport(factoryOf(inst), clock);
    await expect(t.connect(opts)).rejects.toThrowError(/timed out .*handshake/);
    expect(clock.now_).toBe(6000);
  });

  it("re-emits disconnected from the live instance only after a successful connect", async () => {
    const inst = new FakeRustPlus();
    const t = new RustPlusJsTransport(factoryOf(inst), manualClock());
    const seen: string[] = [];
    t.events.on("disconnected", () => seen.push("disconnected"));
    await t.connect(opts);
    inst.emit("disconnected"); // underlying close
    expect(seen).toEqual(["disconnected"]);
  });

  it("every connect tears down any previous instance first", async () => {
    const first = new FakeRustPlus();
    const second = new FakeRustPlus();
    let n = 0;
    const t = new RustPlusJsTransport(() => (n++ === 0 ? first : second), manualClock());
    await t.connect(opts);
    await t.connect(opts);
    expect(first.connected).toBe(false);
    expect(second.connected).toBe(true);
  });
});

describe("protocol contracts (2.5.0 proto fields)", () => {
  it("maps AppError enum: 0/null → ok, others → ProtocolError carrying the code", async () => {
    expect(responseError({ response: { error: 0 } })).toBeNull();
    expect(responseError({ response: {} })).toBeNull();
    expect(responseError({})).toBeNull();

    const failing: Pick<RustPlusInstance, "sendRequestAsync"> = {
      async sendRequestAsync() {
        return { response: { error: 3 } };
      },
    };
    await expect(request(failing, rq.getInfo())).rejects.toThrowError(ProtocolError);
  });

  it("sends exact raw field shapes", async () => {
    const inst = new FakeRustPlus();
    const p = makeProtocol(inst);
    await p.send(rq.subscribeEntity(123));
    await p.toggleSwitch(777, true);
    await p.send(rq.promoteToLeader("76561198"));
    expect(inst.sent[0]).toEqual({ entityId: 123, setSubscription: { value: true } });
    expect(inst.sent[1]).toEqual({ entityId: 777, setEntityValue: { entityId: 777, value: true } });
    expect(inst.sent[2]).toEqual({ promoteToLeader: { steamId: "76561198" } });
  });

  it("switchState reads payload.value and returns null when absent", async () => {
    const inst = new FakeRustPlus();
    inst.responder = async (data) =>
      "getEntityInfo" in data
        ? { response: { error: 0, entityInfo: { payload: { value: 1 } } } }
        : { response: { error: 0 } };
    const p = makeProtocol(inst);
    expect(await p.switchState(42)).toBe(true);

    inst.responder = async () => ({ response: { error: 0 } });
    expect(await p.switchState(42)).toBeNull();
  });
});

describe("SubscriptionOrchestrator (PrimeSubscriptionsAsync parity)", () => {
  it("subscribes then pokes each entity sequentially with 100 ms gaps, marks once", async () => {
    const sleeps: number[] = [];
    const subs = new Set<number>();
    const core = {
      needsSubscribeOnce: (id: number) => !subs.has(id),
      markSubscribed: (id: number) => void subs.add(id),
    };
    const calls: string[] = [];
    const orch = new SubscriptionOrchestrator({
      core,
      send: async (data) => {
        // Record the request-type key, not the shared entityId field.
        calls.push(Object.keys(data).find((k) => k !== "entityId")!);
        return {};
      },
      sleep: async (ms) => {
        sleeps.push(ms);
      },
    });

    await orch.prime([10, 20, 30]);
    expect(calls).toEqual([
      "setSubscription", "getEntityInfo", // 10
      "setSubscription", "getEntityInfo", // 20
      "setSubscription", "getEntityInfo", // 30
    ]);
    expect(sleeps).toEqual([PRIME_GAP_MS, PRIME_GAP_MS]); // gap between, not after last

    calls.length = 0;
    await orch.prime([10, 20, 30]); // same connection → all skipped
    expect(calls).toEqual([]);
  });

  it("a failed poke does not fail priming; subscription still recorded", async () => {
    const subs = new Set<number>();
    const core = {
      needsSubscribeOnce: (id: number) => !subs.has(id),
      markSubscribed: (id: number) => void subs.add(id),
    };
    const orch = new SubscriptionOrchestrator({
      core,
      send: async (data) => {
        if ("getEntityInfo" in data) throw new Error("poke timeout");
        return {};
      },
      sleep: async () => undefined,
    });
    await expect(orch.prime([5])).resolves.toBeUndefined();
    expect(subs.has(5)).toBe(true);
  });

  it("ensure() subscribes exactly once per connection", async () => {
    const subs = new Set<number>();
    const core = {
      needsSubscribeOnce: (id: number) => !subs.has(id),
      markSubscribed: (id: number) => void subs.add(id),
    };
    let count = 0;
    const orch = new SubscriptionOrchestrator({
      core,
      send: async () => {
        count++;
        return {};
      },
      sleep: async () => undefined,
    });
    expect(await orch.ensure(9)).toBe(true);
    expect(await orch.ensure(9)).toBe(false);
    expect(count).toBe(1);
  });
});
