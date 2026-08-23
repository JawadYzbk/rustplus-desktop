/**
 * Push bridge golden tests — envelope shape, destroyed-window filtering, multi-target fan-out,
 * send-failure tolerance.
 */
import { describe, expect, it } from "vitest";
import { createPushBridge, type PushTarget } from "../src/main/push-bridge.js";

interface FakeTarget extends PushTarget {
  sent: Array<{ channel: string; payload: unknown }>;
  setDestroyed(v: boolean): void;
}

function fakeTarget(): FakeTarget {
  let destroyed = false;
  const t: FakeTarget = {
    sent: [],
    isDestroyed: () => destroyed,
    send: (channel, payload) => void t.sent.push({ channel, payload }),
    setDestroyed: (v) => {
      destroyed = v;
    },
  };
  return t;
}

describe("createPushBridge", () => {
  it("wraps events in the { stream, event } envelope on the conn/push channel", () => {
    const t = fakeTarget();
    const bridge = createPushBridge(() => [t]);
    bridge("conn", { kind: "connected", proxy: "direct" });
    expect(t.sent).toEqual([
      { channel: "conn/push", payload: { stream: "conn", event: { kind: "connected", proxy: "direct" } } },
    ]);
  });

  it("skips destroyed targets and tolerates throwing sends", () => {
    const dead = fakeTarget();
    dead.setDestroyed(true);
    const boom = fakeTarget();
    boom.send = () => {
      throw new Error("window vanished");
    };
    const alive = fakeTarget();
    const bridge = createPushBridge(() => [dead, boom, alive]);
    expect(() => bridge("device", { kind: "deviceState" })).not.toThrow();
    expect(dead.sent).toHaveLength(0);
    expect(alive.sent).toHaveLength(1);
  });

  it("fans out to every live window", () => {
    const a = fakeTarget();
    const b = fakeTarget();
    const bridge = createPushBridge(() => [a, b]);
    bridge("poll", { kind: "status", status: { players: 1, maxPlayers: 2, queue: 0, timeString: null } });
    expect(a.sent).toHaveLength(1);
    expect(b.sent).toHaveLength(1);
  });

  it("no-op when there are no windows at all", () => {
    const bridge = createPushBridge(() => []);
    expect(() => bridge("poll", { kind: "team", team: {} })).not.toThrow();
  });
});
