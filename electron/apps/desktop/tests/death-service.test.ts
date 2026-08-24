import { appendFileSync, mkdtempSync, rmSync } from "node:fs";
import { tmpdir } from "node:os";
import { join } from "node:path";
import { afterEach, describe, expect, it } from "vitest";
import { DeathLogStore, DeathService, DeathTracker, summarizeDeaths, type DeathEntry } from "../src/main/services/deaths/death-service.js";

const dirs: string[] = [];
const tempStore = (): DeathLogStore => {
  const dir = mkdtempSync(join(tmpdir(), "rpd-deaths-"));
  dirs.push(dir);
  return new DeathLogStore(dir);
};

afterEach(() => {
  while (dirs.length) rmSync(dirs.pop()!, { recursive: true, force: true });
});

describe("DeathTracker", () => {
  it("establishes a baseline, detects timestamp advances once, and derives survival from respawn", () => {
    const tracker = new DeathTracker();
    const member = (dead: boolean, deathTime: number | null, spawnTime: number | null = null) => ({ steamId: "76561198000000001", name: "Ada", online: true, dead, x: null, y: null, deathTime, spawnTime, grid: null, locationType: null, locationName: null });
    const classify = () => ({ grid: null, locationType: "open" as const, locationName: null });

    expect(tracker.observe([member(true, 100)], classify, 200)).toEqual([]);
    expect(tracker.observe([member(true, 100)], classify, 201)).toEqual([]);
    expect(tracker.observe([member(false, 100, 120)], classify, 202)).toEqual([]);
    const [death] = tracker.observe([member(true, 150)], classify, 203);
    expect(death).toMatchObject({ steamId: "76561198000000001", deathTime: 150, spawnTime: 120, locationType: "open" });
  });

  it("uses base notes before the supplied location hint and computes the Rust grid", () => {
    const store = tempStore();
    const service = new DeathService(store, () => null);
    const team = (dead: boolean, deathTime: number) => ({
      worldSize: 3000,
      mapNotes: [{ icon: 2, x: 100, y: 100, label: "Compound" }],
      members: [{ steamId: "76561198000000002", name: "Bea", dead, deathTime, x: 110, y: 100, locationType: "monument", locationName: "Harbor" }],
    });

    expect(service.observeTeam("server-1", team(true, 10))).toEqual([]);
    const [death] = service.observeTeam("server-1", team(true, 11));
    expect(death).toMatchObject({ locationType: "base", locationName: "Compound", grid: "A19" });
    expect(store.loadEntries("server-1")).toHaveLength(1);
  });
});

describe("death statistics", () => {
  it("matches the legacy grouped summary and ignores malformed JSONL", () => {
    const now = Math.floor(Date.now() / 1000);
    const entries: DeathEntry[] = [
      { victim: "Ada", diedAt: now, spawnAt: now - 120, grid: "A1", type: "base", location: "Compound" },
      { victim: "Ada", diedAt: now - 3600, spawnAt: now - 3660, grid: "A1", type: "open", location: "Open" },
      { victim: "Bea", diedAt: now - 7200, spawnAt: null, grid: "B2", type: "monument", location: "Harbor" },
    ];
    const summary = summarizeDeaths(entries);
    expect(summary.total).toBe(3);
    expect(summary.victims).toBe(2);
    expect(summary.avgSurvival).toBe("1m 30s");
    expect(summary.longestSurvival).toBe("2m 0s");
    expect(summary.deadliestPlace).toBe("Compound");
    expect(summary.deadliestGrid).toBe("A1 (2)");
    expect(summary.byArea[0]).toMatchObject({ type: "base", deaths: 1 });
    expect(summary.byVictim[0]).toMatchObject({ victim: "Ada", deaths: 2 });
    expect(summary.deathsPerDay).toHaveLength(14);

    const store = tempStore();
    const path = store.pathFor("server-2");
    appendFileSync(path, "{not json}\n");
    appendFileSync(path, `${JSON.stringify({ name: "Ada", died_at: now, spawn_at: now - 60, location_type: "base", location_name: "Compound", grid: "A1" })}\n`);
    expect(store.loadEntries("server-2")).toHaveLength(1);
  });
});
