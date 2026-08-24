/**
 * PollService + server-switch golden tests — cadences (10s/5s/2s), status gating, watchdog feed,
 * stop semantics, switch ordering. Loops run against the REAL clock with short injected intervals
 * (a virtual clock would let the loop chain become an infinite microtask sequence).
 */
import { describe, expect, it } from "vitest";
import {
  PollService,
  POLL_INTERVALS,
  toMapMarkers,
  toMapSnapshot,
  toServerStatus,
  formatGameTime,
  switchServer,
  type PollManagerLike,
} from "../src/main/services/rustplus/poll-service.js";

const SHORT = { statusMs: 40, teamMs: 30, markersMs: 20 } as const;

/** Request channel → proto AppResponse content key (info/teamInfo/mapMarkers/time). */
const RESPONSE_KEY: Record<string, string> = {
  getInfo: "info",
  getTime: "time",
  getMap: "map",
  getTeamInfo: "teamInfo",
  getMapMarkers: "mapMarkers",
};

interface FakeMgr extends PollManagerLike {
  responses: Record<string, unknown>;
  statusResults: boolean[];
  sends: string[];
}

function fakeMgr(over: Partial<FakeMgr> = {}): FakeMgr {
  const mgr: FakeMgr = {
    isConnected: true,
    responses: {},
    statusResults: [],
    sends: [],
    ...over,
    send: undefined as never,
    recordStatusResult: undefined as never,
  };
  mgr.send = async (data) => {
    const key = Object.keys(data)[0]!;
    mgr.sends.push(key);
    const res = mgr.responses[key];
    if (res === undefined) throw new Error(`no scripted response for ${key}`);
    // Manager-level send resolves with the response CONTENT under the PROTO key
    // (request getInfo → response info, etc.).
    return { [RESPONSE_KEY[key]!]: res };
  };
  mgr.recordStatusResult = (ok) => void mgr.statusResults.push(ok);
  return mgr;
}

describe("PollService", () => {
  it("runs the legacy cadence constants", () => {
    expect(POLL_INTERVALS).toEqual({ statusMs: 10_000, teamMs: 5_000, markersMs: 2_000 });
  });

  it("emits parsed status/team/markers events and feeds the watchdog on success", async () => {
    const mgr = fakeMgr({
      responses: {
        getInfo: { players: 42, maxPlayers: 200, queuedPlayers: 7 },
        getTime: { time: 12.566666 },
        getTeamInfo: { members: [] },
        getMapMarkers: { markers: [] },
      },
    });
    const svc = new PollService(mgr, undefined, SHORT); // real clock
    const kinds: string[] = [];
    const statuses: Array<{ players: number; timeString: string | null }> = [];
    svc.on("poll", (e) => {
      kinds.push(e.kind);
      if (e.kind === "status") statuses.push(e.status);
    });

    svc.start();
    await new Promise((r) => setTimeout(r, 180));
    svc.stop();

    expect(kinds).toContain("status");
    expect(kinds).toContain("team");
    expect(kinds).toContain("markers");
    expect(statuses[0]).toEqual({ players: 42, maxPlayers: 200, queue: 7, timeString: "12:34" });
    expect(mgr.statusResults).toContain(true);
    expect(svc.isRunning).toBe(false);

    // After stop() the loops wind down: sends must not grow unboundedly.
    const sendsAtStop = mgr.sends.length;
    await new Promise((r) => setTimeout(r, 120));
    expect(mgr.sends.length).toBeLessThan(sendsAtStop + 6);
  }, 5000);

  it("loads one map snapshot and normalizes the live marker payload", async () => {
    const mgr = fakeMgr({
      responses: {
        getMap: { width: 6000, height: 6000, oceanMargin: 1000, jpgImage: Uint8Array.from([0xff, 0xd8]), monuments: [{ token: "launch_site", x: 1200, y: 3400 }] },
        getInfo: { mapSize: 4000, players: 1, maxPlayers: 10 },
      },
    });
    const svc = new PollService(mgr, undefined, SHORT);
    const maps: unknown[] = [];
    svc.on("poll", (event) => event.kind === "map" && maps.push(event.map));

    svc.start();
    await new Promise((resolve) => setTimeout(resolve, 60));
    svc.stop();

    expect(maps).toEqual([{
      width: 6000,
      height: 6000,
      worldSize: 4000,
      oceanMargin: 1000,
      imageBase64: "/9g=",
      monuments: [{ token: "launch_site", x: 1200, y: 3400 }],
    }]);
    expect(mgr.sends).toContain("getMap");
    expect(toMapMarkers({ markers: [{ id: 8, type: 1, x: 123, y: 456, steamId: 7656119, name: "Ada" }] })).toEqual([{
      id: 8,
      type: "Player",
      x: 123,
      y: 456,
      steamId: "7656119",
      rotation: null,
      radius: null,
      alpha: null,
      name: "Ada",
    }]);
    expect(toMapSnapshot(null)).toBeNull();
  }, 5000);

  it("status failure keeps values silent and records false for the watchdog", async () => {
    // No scripted responses → every poll fails.
    const mgr = fakeMgr({ responses: {} });
    const svc = new PollService(mgr, undefined, SHORT);
    const events: unknown[] = [];
    svc.on("poll", (e) => events.push(e));

    svc.start();
    await new Promise((r) => setTimeout(r, 120));
    svc.stop();

    expect(events).toEqual([]); // nothing emitted on failures
    expect(mgr.statusResults.every((r) => r === false)).toBe(true);
    expect(mgr.statusResults.length).toBeGreaterThanOrEqual(2);
  }, 5000);

  it("loops idle while disconnected and resume once connected", async () => {
    const mgr = fakeMgr({
      isConnected: false,
      responses: { getInfo: { players: 1, maxPlayers: 10 }, getTime: { time: 9.0 } },
    });
    const svc = new PollService(mgr, undefined, { ...SHORT, markersMs: 60 });
    const statuses: unknown[] = [];
    svc.on("poll", (e) => e.kind === "status" && statuses.push(e.status));

    svc.start();
    await new Promise((r) => setTimeout(r, 90));
    expect(statuses).toHaveLength(0);

    (mgr as { isConnected: boolean }).isConnected = true;
    await new Promise((r) => setTimeout(r, 120));
    svc.stop();
    expect(statuses.length).toBeGreaterThanOrEqual(1);
    expect(statuses[0]).toEqual({ players: 1, maxPlayers: 10, queue: 0, timeString: "09:00" });
  }, 5000);

  it("toServerStatus gates on players>=0 like the legacy st.Players check", () => {
    expect(toServerStatus({ players: -1, maxPlayers: 100 })).toBeNull();
    expect(toServerStatus(undefined)).toBeNull();
    expect(toServerStatus({ players: 0, maxPlayers: 100 })).toEqual({
      players: 0,
      maxPlayers: 100,
      queue: 0,
      timeString: null,
    });
  });

  it("formatGameTime ports the TryReadTimeHHMM numeric branches", () => {
    expect(formatGameTime(12.566666)).toBe("12:34"); // hours float
    expect(formatGameTime(9)).toBe("09:00");
    expect(formatGameTime(23.999)).toBe("00:00"); // m==60 rolls into next hour (C# parity)
    expect(formatGameTime(750)).toBe("12:30"); // minutes branch 0..1440
    expect(formatGameTime(-1)).toBeNull();
    expect(formatGameTime("12:34")).toBeNull(); // strings not accepted here
  });
});

describe("switchServer", () => {
  it("disconnects the old connection before connecting the next profile", async () => {
    const calls: string[] = [];
    const mgr = {
      disconnect: async () => {
        calls.push("disconnect");
        return {};
      },
      connect: async (p: { host: string }) => {
        calls.push(`connect:${p.host}`);
        return {};
      },
    };
    await switchServer(mgr, { host: "5.6.7.8", port: 28082, steamId64: "765", playerToken: "tok" });
    expect(calls).toEqual(["disconnect", "connect:5.6.7.8"]);
  });

  it("propagates connect failures after tearing down", async () => {
    const mgr = {
      disconnect: async () => undefined,
      connect: async () => {
        throw new Error("Rust+ nicht erreichbar (direkt & Proxy)");
      },
    };
    await expect(
      switchServer(mgr, { host: "1.1.1.1", port: 28015, steamId64: "765", playerToken: "t" }),
    ).rejects.toThrowError(/nicht erreichbar/);
  });
});
