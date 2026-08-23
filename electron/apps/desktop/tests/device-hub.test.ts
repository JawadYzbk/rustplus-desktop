/**
 * DeviceEventHub golden tests — smart-switch state from broadcasts, storage snapshots with sticky
 * tool-cupboard flag, 150 ms trigger delay, 1500 ms TC follow-up pull, reset semantics.
 * Timers are virtual: a queue-backed scheduler advances instantly under test control.
 */
import { describe, expect, it } from "vitest";
import {
  DeviceEventHub,
  extractEntityChanged,
} from "../src/main/services/rustplus/device-hub.js";

/** Virtual timer rig: registered timers run only when the test fires due ones. */
function virtualTimers() {
  let nextId = 1;
  const pending = new Map<number, { fn: () => void; at: number }>();
  let nowMs = 1000;
  return {
    now: () => nowMs,
    setTimer: (fn: () => void, ms: number) => {
      const id = nextId++;
      pending.set(id, { fn, at: nowMs + ms });
      return id;
    },
    clearTimer: (h: unknown) => void pending.delete(h as number),
    /** Fires every timer due within `advanceMs`, in time order. */
    async tick(advanceMs: number): Promise<void> {
      const deadline = nowMs + advanceMs;
      for (;;) {
        const due = [...pending.values()].filter((t) => t.at <= deadline).sort((a, b) => a.at - b.at)[0];
        if (!due) break;
        // remove first (id lookup by identity)
        for (const [id, t] of pending) {
          if (t === due) {
            pending.delete(id);
            break;
          }
        }
        nowMs = Math.max(nowMs, due.at);
        due.fn();
        await Promise.resolve();
        await Promise.resolve();
      }
      nowMs = deadline;
    },
    pendingCount: () => pending.size,
  };
}

type Hub = DeviceEventHub & { deps_: ReturnType<typeof makeDeps> };
function makeDeps(sendResponses: Record<number, unknown> = {}) {
  const vt = virtualTimers();
  const sends: number[] = [];
  const send = async (data: Record<string, unknown>) => {
    const entityId = data["entityId"] as number;
    sends.push(entityId);
    const payload = sendResponses[entityId];
    if (payload === undefined) throw new Error("no scripted entityInfo");
    return { entityInfo: { payload } };
  };
  return { ...vt, send, sends, sendResponses };
}

const eventsOf = (hub: DeviceEventHub) => {
  const evts: unknown[] = [];
  hub.on("event", (e) => evts.push(e));
  return evts;
};

describe("extractEntityChanged", () => {
  it("pulls broadcast.entityChanged and ignores other messages", () => {
    const ok = extractEntityChanged({
      broadcast: { entityChanged: { entityId: 7, payload: { value: true } } },
    });
    expect(ok).toEqual({ entityId: 7, payload: { value: true } });

    expect(extractEntityChanged({ broadcast: { teamChanged: {} } })).toBeNull();
    expect(extractEntityChanged(null)).toBeNull();
    expect(extractEntityChanged({ broadcast: { entityChanged: { payload: {} } } })).toBeNull();
  });
});

describe("DeviceEventHub — smart switches", () => {
  it("boolean value broadcasts emit deviceState immediately", async () => {
    const deps = makeDeps();
    const hub = new DeviceEventHub(deps) as Hub;
    const evts = eventsOf(hub);

    await hub.handleEntityChanged(7, { value: true });
    await hub.handleEntityChanged(7, { value: false });

    expect(evts).toEqual([
      { kind: "deviceState", entityId: 7, on: true, deviceType: "SmartSwitch" },
      { kind: "deviceState", entityId: 7, on: false, deviceType: "SmartSwitch" },
    ]);
  });
});

describe("DeviceEventHub — storage snapshots", () => {
  it("entityInfo response builds snapshot; TC gets sticky flag and 1500ms follow-up pull", async () => {
    const deps = makeDeps();
    const hub = new DeviceEventHub(deps) as Hub;
    const evts = eventsOf(hub);

    hub.handleEntityInfoResponse(42, {
      hasProtection: true,
      protectionExpiry: 7200,
      capacity: 30,
      items: [{ itemId: -1779184422, quantity: 250, itemIsBlueprint: false }],
    });

    // Snapshot emitted synchronously; follow-up pull scheduled but not yet fired.
    expect(evts).toHaveLength(1);
    const e0 = evts[0] as { kind: string; entityId: number; snapshot: Record<string, unknown> };
    expect(e0.kind).toBe("storageSnapshot");
    expect(e0.snapshot.isToolCupboard).toBe(true);
    expect(e0.snapshot.upkeepSeconds).toBe(7200); // protectionExpiry → upkeep
    expect(e0.snapshot.items).toHaveLength(1);

    await deps.tick(2000);
    expect(deps.sends).toEqual([42]); // exactly one follow-up getEntityInfo
  });

  it("sticky TC: an event without hasProtection must NOT demote a known TC to a box", () => {
    const deps = makeDeps();
    const hub = new DeviceEventHub(deps) as Hub;
    const snaps: Array<{ isToolCupboard: boolean; upkeepSeconds: number | null }> = [];
    hub.on("event", (e) => e.kind === "storageSnapshot" && snaps.push(e.snapshot));

    hub.handleEntityInfoResponse(99, { hasProtection: true, protectionExpiry: 3600 });
    // Second event lacks the optional field entirely.
    hub.handleEntityInfoResponse(99, { items: [] });

    expect(snaps[0]!.isToolCupboard).toBe(true);
    expect(snaps[1]!.isToolCupboard).toBe(true); // sticky!
    expect(snaps[1]!.upkeepSeconds).toBeNull(); // legacy keeps upkeep null when expiry absent
  });

  it("boxes never carry upkeep; non-TC without prior cache stays box", () => {
    const deps = makeDeps();
    const hub = new DeviceEventHub(deps) as Hub;
    const snaps: Array<Record<string, unknown>> = [];
    hub.on("event", (e) => e.kind === "storageSnapshot" && snaps.push(e.snapshot));

    hub.handleEntityInfoResponse(5, { items: [{ itemId: 1, quantity: 3, itemIsBlueprint: false }] });
    expect(snaps[0]).toMatchObject({ isToolCupboard: false, upkeepSeconds: null, capacity: null });
    expect(deps.pendingCount()).toBe(0); // no follow-up pulls for boxes
  });

  it("storage-ish broadcast waits 150ms then pulls fresh entityInfo", async () => {
    const deps = makeDeps({
      11: { hasProtection: false, items: [{ itemId: 2, quantity: 1, itemIsBlueprint: false }] },
    });
    const hub = new DeviceEventHub(deps) as Hub;

    // Fire-and-forget: the handler parks on the virtual 150 ms timer until tick() advances.
    const pending = hub.handleEntityChanged(11, { items: [], hasProtection: false });
    expect(deps.sends).toEqual([]); // nothing yet inside the 150 ms window

    await deps.tick(200);
    await pending;
    expect(deps.sends).toEqual([11]);
  }, 5000);

  it("reset() clears cache and cancels scheduled pulls", async () => {
    const deps = makeDeps();
    const hub = new DeviceEventHub(deps) as Hub;
    hub.handleEntityInfoResponse(77, { hasProtection: true, protectionExpiry: 60 });
    expect(deps.pendingCount()).toBe(1);

    hub.reset();
    expect(deps.pendingCount()).toBe(0);
    expect(hub.tryGetCachedStorage(77)).toBeUndefined();

    await deps.tick(5000);
    expect(deps.sends).toEqual([]); // cancelled pull never ran
  });
});
