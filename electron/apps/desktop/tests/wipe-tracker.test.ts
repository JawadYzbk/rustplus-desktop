import { describe, expect, it, vi } from "vitest";
import { mkdtempSync, appendFileSync, readdirSync, rmSync } from "node:fs";
import { tmpdir } from "node:os";
import { join } from "node:path";
import { CloudApiClient } from "../src/main/services/cloud/cloud-api-client.js";
import { PlayerWipeTrackerEngine } from "../src/main/services/wipe-tracker/engine.js";
import { buildInsights } from "../src/main/services/wipe-tracker/insights.js";
import type { PlayerObservation } from "../src/main/services/wipe-tracker/models.js";
import { projectMapPoint } from "../src/main/services/wipe-tracker/models.js";
import { buildWipeKey, PlayerWipeTrackerService } from "../src/main/services/wipe-tracker/service.js";
import { PlayerWipeTrackerStore } from "../src/main/services/wipe-tracker/store.js";

const STEAM_ID = "76561198000000001";
const observation = (timestampUtc: Date, over: Partial<PlayerObservation> = {}): PlayerObservation => ({
  steamId: STEAM_ID,
  name: "Player",
  timestampUtc,
  sessionId: "s1",
  isConnected: true,
  snapshotValid: true,
  online: true,
  dead: false,
  afk: false,
  x: 0,
  y: 0,
  locationType: "open",
  locationName: null,
  grid: "A1",
  spawnTime: null,
  deathTime: null,
  ...over,
});

describe("PlayerWipeTrackerEngine", () => {
  it("establishes a baseline without elapsed coverage and classifies priorities", () => {
    const now = new Date("2026-01-01T00:00:00Z");
    const engine = new PlayerWipeTrackerEngine();
    expect(engine.observe(observation(now))).toBe(true);
    expect(engine.summarize().coverageSeconds).toBe(0);
    expect(PlayerWipeTrackerEngine.classify(observation(now, { snapshotValid: false }))).toBe("unknown");
    expect(PlayerWipeTrackerEngine.classify(observation(now, { online: false }))).toBe("offline");
    expect(PlayerWipeTrackerEngine.classify(observation(now, { dead: true }))).toBe("dead");
    expect(PlayerWipeTrackerEngine.classify(observation(now, { afk: true }))).toBe("afk");
    expect(PlayerWipeTrackerEngine.classify(observation(now, { x: 20 }), 20)).toBe("moving");
  });

  it("turns a reconnect gap into unknown time and excludes it from distance", () => {
    const start = new Date("2026-01-01T00:00:00Z");
    const engine = new PlayerWipeTrackerEngine();
    engine.observe(observation(start, { x: 0 }));
    engine.observe(observation(new Date(start.getTime() + 5_000), { x: 5 }));
    engine.observe(observation(new Date(start.getTime() + 120_000), { x: 500, sessionId: "s2" }));
    const summary = engine.summarize();
    expect(summary.unknownSeconds).toBeGreaterThan(60);
    expect(summary.estimatedDistance).toBeLessThan(100);
  });

  it("projects world corners through a padded uniform map image", () => {
    const projection = { viewWidth: 800, viewHeight: 500, imageWidth: 1000, imageHeight: 1000, worldRectX: 100, worldRectY: 100, worldRectWidth: 800, worldRectHeight: 800, worldSize: 4000 };
    expect(projectMapPoint(projection, 0, 4000)).toEqual({ x: 200, y: 50 });
    expect(projectMapPoint(projection, 4000, 0)).toEqual({ x: 600, y: 450 });
  });

  it("derives monument, blind-gap, peak-hour, and current-state insights", () => {
    const start = new Date("2026-01-01T12:00:00Z");
    const engine = new PlayerWipeTrackerEngine();
    const observations: PlayerObservation[] = [];
    const add = (seconds: number, x: number, over: Partial<PlayerObservation>): void => {
      const item = observation(new Date(start.getTime() + seconds * 1000), { x, y: 100, ...over });
      observations.push(item);
      engine.observe(item);
    };
    for (const [seconds, x] of [[0, 100], [10, 112], [20, 124], [30, 136], [40, 148]] as const) add(seconds, x, { locationType: "monument", locationName: "Launch Site" });
    add(50, 400, {});
    add(60, 412, {});
    add(180, 420, { sessionId: "s2" });
    add(190, 432, { sessionId: "s2" });
    const summary = engine.summarize();
    const insights = buildInsights(observations, engine.getSegments(), summary, new Date(start.getTime() + 200_000));
    expect(insights.firstSeenUtc).toEqual(start);
    expect(insights.topMonument).toBe("Launch Site");
    expect(insights.topMonumentVisits).toBe(1);
    expect(insights.topMonumentSeconds).toBeGreaterThanOrEqual(30);
    expect(insights.longestBlindGapSeconds).toBeGreaterThanOrEqual(100);
    expect(insights.currentState).toBe("stationary");
    expect(insights.peakHourLocal).not.toBeNull();
  });
});

describe("PlayerWipeTrackerStore", () => {
  it("skips corrupt JSONL lines, deduplicates observations, and deletes its root", () => {
    const directory = mkdtempSync(join(tmpdir(), "rpd-wipe-"));
    const store = new PlayerWipeTrackerStore(directory);
    const item = { schemaVersion: 1 as const, kind: "observation" as const, observation: observation(new Date("2026-01-01T00:00:00Z")) };
    expect(store.append("server", "wipe", STEAM_ID, item)).toBe(true);
    expect(store.append("server", "wipe", STEAM_ID, item)).toBe(true);
    const file = join(directory, "server", "wipe", `${STEAM_ID}.jsonl`);
    appendFileSync(file, "not json\n", "utf8");
    expect(store.load("server", "wipe", STEAM_ID)).toHaveLength(1);
    expect(readdirSync(join(directory, "server", "wipe"))).toContain(`${STEAM_ID}.jsonl`);
    store.deleteAll();
    expect(store.storageBytes()).toBe(0);
    rmSync(directory, { recursive: true, force: true });
  });

  it("keeps a wipe map inside its wipe directory", () => {
    const directory = mkdtempSync(join(tmpdir(), "rpd-wipe-map-"));
    const store = new PlayerWipeTrackerStore(directory);
    const expected = { pngBytes: new Uint8Array([1, 2, 3]), worldSize: 4500, worldRectX: 10, worldRectY: 20, worldRectWidth: 900, worldRectHeight: 900 };
    store.saveWipeMap("server", "wipe-a", expected);
    expect(store.loadWipeMap("server", "wipe-a")).toEqual(expected);
    expect(store.loadWipeMap("server", "wipe-b")).toBeNull();
    store.deleteAll();
    rmSync(directory, { recursive: true, force: true });
  });
});

describe("PlayerWipeTrackerService", () => {
  it("builds the Laravel day payload with a lowercase SHA-256 checksum", async () => {
    const directory = mkdtempSync(join(tmpdir(), "rpd-wipe-service-"));
    const fetchImpl = vi.fn<typeof fetch>().mockResolvedValue(new Response(JSON.stringify({ data: { ok: true } }), { status: 200 }));
    const cloud = new CloudApiClient({ baseUrl: "https://cloud.test/api", clientVersion: "8.1.0", token: () => "token", fetchImpl });
    const store = new PlayerWipeTrackerStore(directory);
    const service = new PlayerWipeTrackerService(store, {
      enabled: () => true,
      cloudBackupEnabled: () => false,
      capabilities: () => ({ planCode: "development", isTrackerAvailable: true, canTrackTeam: true, canUseCloudSync: true, canUseAdvancedViews: true, canUseRouteReplay: true, canExport: true, maxTrackedPlayers: 32, retainedWipes: 12, cloudRetentionDays: 365, fetchedAt: new Date().toISOString() }),
    }, cloud);
    const start = new Date("2026-01-01T00:00:00Z");
    service.startConnection("server-28015", start, "map", STEAM_ID, "session-1");
    service.observe(observation(start, { locationType: "monument", locationName: "Launch Site" }));
    service.observe(observation(new Date(start.getTime() + 10_000), { x: 12, locationType: "monument", locationName: "Launch Site" }));
    const request = service.buildCloudDay(STEAM_ID, "2026-01-01", "Player");
    expect(request?.player_steam_id).toBe(STEAM_ID);
    expect(request?.payload.observations).toHaveLength(2);
    expect(request?.checksum).toMatch(/^[a-f0-9]{64}$/);
    expect(service.getPlayer(STEAM_ID)?.observations).toHaveLength(2);
    expect(service.getPlayer(STEAM_ID)?.segments.length).toBeGreaterThan(0);
    await service.dispose();
    rmSync(directory, { recursive: true, force: true });
  });

  it("lists, restores, and deletes Laravel archives through the documented routes", async () => {
    const directory = mkdtempSync(join(tmpdir(), "rpd-wipe-cloud-"));
    const wipeStartedAt = "2026-01-01T00:00:00.000Z";
    const wipeKey = buildWipeKey("server-28015", new Date(wipeStartedAt), "map");
    const fetchImpl = vi.fn<typeof fetch>().mockImplementation(async (input, init) => {
      const url = String(input);
      if (init?.method === "DELETE") return new Response(JSON.stringify({ data: { deleted: url.endsWith("/player-wipe-tracker") ? 2 : true } }), { status: 200 });
      if (url.endsWith("/player-wipe-tracker/wipes")) return new Response(JSON.stringify({ data: [{ id: "archive-1", wipe_key: wipeKey, wipe_started_at: wipeStartedAt, server: { server_key: "server-28015", name: "Test server" }, player_count: 1, stored_bytes: 1234, players: [{ player_steam_id: STEAM_ID, day_count: 1 }] }] }), { status: 200 });
      if (url.includes("/players/")) return new Response(JSON.stringify({ data: [{ player_steam_id: STEAM_ID, player_name: "Player", day: "2026-01-01", payload: { schema_version: 1, generated_at: wipeStartedAt, observation_sessions: ["cloud-session"], observations: [{ timestamp: wipeStartedAt, x: 10, y: 20, state: "stationary", location_type: "open", location_name: null, grid: "A1", event: null }] } }] }), { status: 200 });
      if (url.includes("/wipes/archive-1")) return new Response(JSON.stringify({ data: { id: "archive-1", wipe_key: wipeKey, wipe_started_at: wipeStartedAt, server: { server_key: "server-28015", name: "Test server" }, player_count: 1, stored_bytes: 1234, players: [{ player_steam_id: STEAM_ID, day_count: 1 }] } }), { status: 200 });
      if ((input as URL).toString().includes("player-wipe-tracker")) return new Response(JSON.stringify({ data: { deleted: true } }), { status: 200 });
      return new Response("not found", { status: 404 });
    });
    const cloud = new CloudApiClient({ baseUrl: "https://cloud.test/api", clientVersion: "8.1.0", token: () => "token", fetchImpl });
    const store = new PlayerWipeTrackerStore(directory);
    const service = new PlayerWipeTrackerService(store, {
      enabled: () => true,
      cloudBackupEnabled: () => false,
      capabilities: () => ({ planCode: "supporter", isTrackerAvailable: true, canTrackTeam: true, canUseCloudSync: true, canUseAdvancedViews: true, canUseRouteReplay: true, canExport: true, maxTrackedPlayers: 32, retainedWipes: 12, cloudRetentionDays: 365, fetchedAt: wipeStartedAt }),
    }, cloud);
    service.startConnection("server-28015", new Date(wipeStartedAt), "map", STEAM_ID, "session-1");
    await expect(service.getCloudArchives()).resolves.toHaveLength(1);
    await expect(service.restoreCloudArchive("archive-1")).resolves.toMatchObject({ archiveId: "archive-1", players: 1, days: 1, observations: 1, isCurrentWipe: true });
    expect(service.getPlayer(STEAM_ID)?.observationCount).toBe(1);
    await expect(service.deleteCloudArchive("archive-1")).resolves.toBe(true);
    await expect(service.deleteAllCloud()).resolves.toBe(2);
    await service.dispose();
    rmSync(directory, { recursive: true, force: true });
  });
});
