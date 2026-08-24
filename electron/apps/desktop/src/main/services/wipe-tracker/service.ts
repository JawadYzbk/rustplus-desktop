import { createHash } from "node:crypto";
import type { PlayerWipeCapabilities } from "../cloud/cloud-service.js";
import { CloudApiClient, CloudApiError } from "../cloud/cloud-api-client.js";
import { buildInsights } from "./insights.js";
import { PlayerWipeTrackerEngine, MAX_CONTINUITY_GAP_SECONDS } from "./engine.js";
import {
  type CloudDayUploadRequest,
  type CloudTrackerDayPayload,
  type PlayerActivityState,
  type PlayerObservation,
  type TrackerInsights,
  type TrackerPersistedObservation,
  type TrackerSummary,
  type TrackerWipeMap,
} from "./models.js";
import { observationsFromStore, PlayerWipeTrackerStore } from "./store.js";

export const FREE_WIPE_CAPABILITIES: PlayerWipeCapabilities = {
  planCode: "free",
  isTrackerAvailable: false,
  canTrackTeam: false,
  canUseCloudSync: false,
  canUseAdvancedViews: false,
  canUseRouteReplay: false,
  canExport: false,
  maxTrackedPlayers: 1,
  retainedWipes: 1,
  cloudRetentionDays: 0,
  fetchedAt: new Date(0).toISOString(),
};

export interface WipeTrackerSettings {
  enabled: () => boolean;
  cloudBackupEnabled: () => boolean;
  capabilities: () => PlayerWipeCapabilities;
}

export interface WipePlayerDto {
  steamId: string;
  name: string;
  summary: TrackerSummary;
  insights: TrackerInsights;
  observationCount: number;
  observations: WipeReplayPoint[];
  segments: WipeReplaySegment[];
}

export interface WipeReplayPoint {
  timestampUtc: Date;
  x: number | null;
  y: number | null;
  state: PlayerActivityState;
  locationType: PlayerObservation["locationType"];
  locationName: string | null;
  grid: string | null;
  event: "death" | "respawn" | null;
  sessionId: string;
}

export interface WipeReplaySegment {
  startUtc: Date;
  endUtc: Date;
  state: PlayerActivityState;
}

export interface CloudArchivePlayer {
  steamId: string;
  dayCount: number;
}

export interface CloudArchiveSummary {
  id: string;
  serverKey: string;
  serverName: string;
  wipeKey: string;
  wipeStartedAtUtc: Date | null;
  firstObservedAtUtc: Date | null;
  lastObservedAtUtc: Date | null;
  playerCount: number | null;
  storedBytes: number | null;
  players: CloudArchivePlayer[];
}

export interface CloudRestoreResult {
  archiveId: string;
  players: number;
  days: number;
  observations: number;
  isCurrentWipe: boolean;
}

/** Coordinates Rust+ team snapshots, local JSONL history, pure engines, and Laravel day uploads. */
export class PlayerWipeTrackerService {
  private readonly engines = new Map<string, PlayerWipeTrackerEngine>();
  private readonly cloudQueue: CloudSyncQueue;
  private serverKey: string | null = null;
  private wipeKey: string | null = null;
  private sessionId: string | null = null;
  private wipeStartedAtUtc: Date | null = null;
  private ownSteamId = "";

  constructor(
    private readonly store: PlayerWipeTrackerStore,
    private readonly settings: WipeTrackerSettings,
    private readonly cloud: CloudApiClient,
  ) {
    this.cloudQueue = new CloudSyncQueue((request) => uploadDay(this.cloud, request));
  }

  get currentServerKey(): string | null { return this.serverKey; }
  get currentWipeKey(): string | null { return this.wipeKey; }
  get currentSessionId(): string | null { return this.sessionId; }
  get trackedPlayers(): string[] { return [...this.engines.keys()]; }

  startConnection(serverKey: string, wipeTimeUtc: Date | null, mapIdentity: string | null, ownSteamId: string, sessionId?: string): void {
    this.serverKey = serverKey;
    this.wipeKey = buildWipeKey(serverKey, wipeTimeUtc, mapIdentity);
    this.wipeStartedAtUtc = wipeTimeUtc ? new Date(wipeTimeUtc.getTime()) : null;
    this.sessionId = sessionId?.trim() || cryptoRandomId();
    this.ownSteamId = ownSteamId;
    this.engines.clear();
    for (const steamId of this.store.loadPlayerIds(serverKey, this.wipeKey)) {
      if (canTrack(steamId, this.ownSteamId, this.settings.capabilities())) this.engines.set(steamId, this.loadEngine(steamId));
    }
  }

  observe(observation: PlayerObservation): void {
    if (!this.settings.enabled() || !this.serverKey || !this.wipeKey || !this.settings.capabilities().isTrackerAvailable || !canTrack(observation.steamId, this.ownSteamId, this.settings.capabilities())) return;
    const normalized = { ...observation, sessionId: this.sessionId ?? observation.sessionId };
    if (!this.sessionId) this.sessionId = normalized.sessionId;
    let engine = this.engines.get(normalized.steamId);
    if (!engine) {
      engine = this.loadEngine(normalized.steamId);
      this.engines.set(normalized.steamId, engine);
    }
    if (!engine.observe(normalized)) return;
    const persisted: TrackerPersistedObservation = { schemaVersion: 1, kind: "observation", observation: normalized };
    this.store.append(this.serverKey, this.wipeKey, normalized.steamId, persisted);
    const capabilities = this.settings.capabilities();
    if (this.settings.cloudBackupEnabled() && capabilities.canUseCloudSync) {
      const request = this.buildCloudDay(normalized.steamId, dateKey(normalized.timestampUtc), normalized.name);
      if (request) this.cloudQueue.enqueue(request);
    }
  }

  /** Normalizes the raw getTeamInfo member shape already used by the automation pipeline. */
  observeTeam(team: unknown): void {
    const members = (team as { members?: unknown })?.members;
    if (!Array.isArray(members)) return;
    for (const raw of members) {
      const member = (raw ?? {}) as Record<string, unknown>;
      const steamId = String(member.steamId ?? member.steamId64 ?? "");
      if (!/^\d{17}$/.test(steamId)) continue;
      const locationType = member.locationType === "monument" || member.locationType === "base" || member.locationType === "open" ? member.locationType : "unknown";
      this.observe({
        steamId,
        name: typeof member.name === "string" && member.name ? member.name : steamId,
        timestampUtc: new Date(),
        sessionId: this.sessionId ?? cryptoRandomId(),
        isConnected: true,
        snapshotValid: true,
        online: member.online === true,
        dead: member.dead === true || member.isDead === true,
        afk: member.afk === true || member.isAfk === true,
        x: numberOrNull(member.x),
        y: numberOrNull(member.y),
        locationType,
        locationName: stringOrNull(member.locationName),
        grid: stringOrNull(member.grid),
        spawnTime: numberOrNull(member.spawnTime),
        deathTime: numberOrNull(member.deathTime),
      });
    }
  }

  disconnect(timestampUtc = new Date()): void {
    if (!this.serverKey || !this.wipeKey) return;
    for (const [steamId, engine] of this.engines) {
      const last = engine.lastObservation;
      engine.endSession(timestampUtc);
      if (last && timestampUtc > last.timestampUtc) {
        this.store.append(this.serverKey, this.wipeKey, steamId, {
          schemaVersion: 1,
          kind: "observation",
          observation: { ...last, timestampUtc, isConnected: false, snapshotValid: false },
        });
      }
    }
    this.sessionId = null;
  }

  getPlayer(steamId: string, nowUtc = new Date()): WipePlayerDto | null {
    if (!this.serverKey || !this.wipeKey) return null;
    const observations = this.getObservations(steamId);
    if (observations.length === 0) return null;
    const engine = this.engines.get(steamId) ?? this.loadEngine(steamId);
    const summary = engine.summarize();
    return { steamId, name: observations.at(-1)?.name ?? steamId, summary, insights: buildInsights(observations, engine.getSegments(), summary, nowUtc), observationCount: observations.length, observations: buildReplayPoints(observations), segments: engine.getSegments().map((segment) => ({ startUtc: segment.startUtc, endUtc: segment.endUtc, state: segment.state })) };
  }

  getPlayers(nowUtc = new Date()): WipePlayerDto[] {
    const ids = new Set([...this.engines.keys()]);
    if (this.serverKey && this.wipeKey) for (const id of this.store.loadPlayerIds(this.serverKey, this.wipeKey)) ids.add(id);
    return [...ids].map((id) => this.getPlayer(id, nowUtc)).filter((player): player is WipePlayerDto => player !== null);
  }

  getObservations(steamId: string): PlayerObservation[] {
    return this.serverKey && this.wipeKey ? observationsFromStore(this.store, this.serverKey, this.wipeKey, steamId) : [];
  }

  buildCloudDay(steamId: string, day: string, playerName: string | null): CloudDayUploadRequest | null {
    if (!this.serverKey || !this.wipeKey || !/^\d{4}-\d{2}-\d{2}$/.test(day)) return null;
    const observations = this.getObservations(steamId).filter((item) => dateKey(item.timestampUtc) === day);
    if (observations.length === 0) return null;
    const cloudObservations = [] as CloudTrackerDayPayload["observations"];
    let previous: PlayerObservation | null = null;
    for (const observation of observations) {
      const displacement = previous?.x !== null && previous?.y !== null && observation.x !== null && observation.y !== null && previous ? Math.hypot(previous.x - observation.x, previous.y - observation.y) : 0;
      const continuity = Boolean(previous && previous.sessionId === observation.sessionId && observation.timestampUtc > previous.timestampUtc && secondsBetween(previous.timestampUtc, observation.timestampUtc) <= MAX_CONTINUITY_GAP_SECONDS && previous.isConnected && previous.snapshotValid && observation.isConnected && observation.snapshotValid);
      const state = !continuity && previous ? "unknown" : PlayerWipeTrackerEngine.classify(observation, displacement);
      const event = previous && !previous.dead && observation.dead ? "death" : previous && previous.dead && !observation.dead ? "respawn" : null;
      cloudObservations.push({ timestamp: observation.timestampUtc.toISOString(), x: observation.x, y: observation.y, state, location_type: observation.locationType, location_name: observation.locationName, grid: observation.grid, event });
      previous = observation;
    }
    const payload: CloudTrackerDayPayload = { schema_version: 1, generated_at: new Date().toISOString(), observation_sessions: [...new Set(observations.map((item) => item.sessionId))], observations: cloudObservations };
    const serialized = JSON.stringify(payload);
    return {
      server_key: this.serverKey,
      wipe_key: this.wipeKey,
      wipe_started_at: this.wipeStartedAtUtc?.toISOString() ?? null,
      player_steam_id: steamId,
      player_name: playerName,
      day,
      schema_version: 1,
      payload,
      checksum: createHash("sha256").update(serialized, "utf8").digest("hex"),
    };
  }

  async getCloudArchives(): Promise<CloudArchiveSummary[]> {
    if (!this.settings.capabilities().canUseCloudSync) return [];
    const payload = await this.cloud.request<unknown>("GET", "player-wipe-tracker/wipes");
    return Array.isArray(payload) ? payload.map(parseArchive).filter((archive): archive is CloudArchiveSummary => archive !== null) : [];
  }

  async restoreCloudArchive(archiveId: string): Promise<CloudRestoreResult> {
    if (!this.settings.capabilities().canUseCloudSync) throw new Error("Cloud restore is available on premium plans only.");
    const archivePayload = await this.cloud.request<unknown>("GET", `player-wipe-tracker/wipes/${encodeURIComponent(archiveId)}`);
    const archive = parseArchive(archivePayload);
    if (!archive || !archive.serverKey || !archive.wipeKey) throw new Error("The cloud archive is missing its server or wipe identity.");

    let players = 0;
    let days = 0;
    let observations = 0;
    const restored = new Set<string>();
    for (const player of archive.players) {
      const dayPayload = await this.cloud.request<unknown>("GET", `player-wipe-tracker/wipes/${encodeURIComponent(archive.id)}/players/${encodeURIComponent(player.steamId)}`);
      const playerDays = Array.isArray(dayPayload) ? dayPayload.map(parseRestoreDay).filter((day): day is RestoreDay => day !== null) : [];
      if (playerDays.length === 0) continue;
      players += 1;
      days += playerDays.length;
      restored.add(player.steamId);
      for (const day of playerDays) observations += this.importCloudDay(archive, day);
    }
    const isCurrentWipe = archive.serverKey === this.serverKey && archive.wipeKey === this.wipeKey;
    if (isCurrentWipe) for (const steamId of restored) this.engines.set(steamId, this.loadEngine(steamId));
    return { archiveId: archive.id, players, days, observations, isCurrentWipe };
  }

  async deleteCloudArchive(archiveId: string): Promise<boolean> {
    const payload = await this.cloud.request<unknown>("DELETE", `player-wipe-tracker/wipes/${encodeURIComponent(archiveId)}`);
    return record(payload)["deleted"] === true;
  }

  async deleteAllCloud(): Promise<number> {
    const payload = await this.cloud.request<unknown>("DELETE", "player-wipe-tracker");
    const deleted = record(payload)["deleted"];
    return typeof deleted === "number" && Number.isFinite(deleted) ? Math.max(0, Math.trunc(deleted)) : 0;
  }

  storageBytes(): number { return this.store.storageBytes(); }
  hasCurrentWipeMap(): boolean { return this.serverKey !== null && this.wipeKey !== null && this.store.hasWipeMap(this.serverKey, this.wipeKey); }
  saveCurrentWipeMap(map: TrackerWipeMap): void { if (this.serverKey && this.wipeKey && !this.hasCurrentWipeMap()) this.store.saveWipeMap(this.serverKey, this.wipeKey, map); }
  loadCurrentWipeMap(): TrackerWipeMap | null { return this.serverKey && this.wipeKey ? this.store.loadWipeMap(this.serverKey, this.wipeKey) : null; }
  deleteWipe(serverKey: string, wipeKey: string): void { this.store.deleteWipe(serverKey, wipeKey); }
  deleteAll(): void { this.store.deleteAll(); }
  async dispose(): Promise<void> { await this.cloudQueue.dispose(); }

  private loadEngine(steamId: string): PlayerWipeTrackerEngine {
    const engine = new PlayerWipeTrackerEngine();
    if (this.serverKey && this.wipeKey) for (const item of observationsFromStore(this.store, this.serverKey, this.wipeKey, steamId)) engine.observe(item);
    return engine;
  }

  private importCloudDay(archive: CloudArchiveSummary, day: RestoreDay): number {
    const sessionId = day.payload.observation_sessions.find((session) => session.trim()) ?? `cloud:${archive.id}:${day.day}`;
    const playerName = day.playerName?.trim() || day.playerSteamId;
    let imported = 0;
    for (const point of day.payload.observations) {
      const timestamp = new Date(point.timestamp);
      if (Number.isNaN(timestamp.getTime())) continue;
      const state = parseState(point.state);
      const connected = state !== "unknown";
      const observation: PlayerObservation = {
        steamId: day.playerSteamId,
        name: playerName,
        timestampUtc: timestamp,
        sessionId,
        isConnected: connected,
        snapshotValid: connected,
        online: state !== "offline" && state !== "unknown",
        dead: state === "dead",
        afk: state === "afk",
        x: point.x,
        y: point.y,
        locationType: parseLocation(point.location_type),
        locationName: point.location_name,
        grid: point.grid,
        spawnTime: null,
        deathTime: null,
      };
      this.store.append(archive.serverKey, archive.wipeKey, day.playerSteamId, { schemaVersion: 1, kind: "observation", observation });
      imported += 1;
    }
    return imported;
  }
}

class CloudSyncQueue {
  private readonly pending = new Map<string, CloudDayUploadRequest>();
  private readonly scheduled = new Set<string>();
  private readonly queue: string[] = [];
  private running = false;
  private disposed = false;

  constructor(private readonly upload: (request: CloudDayUploadRequest) => Promise<number>) {}

  enqueue(request: CloudDayUploadRequest): boolean {
    const key = `${request.server_key}|${request.wipe_key}|${request.player_steam_id}|${request.day}`;
    this.pending.set(key, request);
    if (this.scheduled.has(key)) return true;
    if (this.queue.length >= 64) {
      this.pending.delete(key);
      return false;
    }
    this.scheduled.add(key);
    this.queue.push(key);
    void this.work();
    return true;
  }

  private async work(): Promise<void> {
    if (this.running) return;
    this.running = true;
    try {
      while (!this.disposed && this.queue.length > 0) {
        const key = this.queue.shift()!;
        const request = this.pending.get(key);
        this.pending.delete(key);
        if (!request) {
          this.scheduled.delete(key);
          continue;
        }
        for (let attempt = 0; attempt < 3; attempt += 1) {
          let status = 599;
          try { status = await this.upload(request); } catch { /* best effort; next attempt */ }
          if ((status >= 200 && status < 300) || status === 409 || status === 403 || status === 422) break;
          if (attempt < 2) await delay(250 * (2 ** attempt));
        }
        if (this.pending.has(key) && !this.disposed) this.queue.push(key);
        else this.scheduled.delete(key);
      }
    } finally {
      this.running = false;
    }
  }

  async dispose(): Promise<void> {
    this.disposed = true;
    this.queue.length = 0;
    this.pending.clear();
    this.scheduled.clear();
  }
}

async function uploadDay(client: CloudApiClient, request: CloudDayUploadRequest): Promise<number> {
  try {
    await client.request("PUT", "player-wipe-tracker/days", { json: request });
    return 200;
  } catch (error) {
    return error instanceof CloudApiError ? error.status : 599;
  }
}

interface RestoreDay {
  playerSteamId: string;
  playerName: string | null;
  day: string;
  payload: CloudTrackerDayPayload;
}

function parseArchive(value: unknown): CloudArchiveSummary | null {
  const item = record(value);
  const id = stringValue(item["id"]);
  if (!id) return null;
  const server = record(item["server"]);
  const players = Array.isArray(item["players"]) ? item["players"].map((player) => {
    const row = record(player);
    const steamId = stringValue(row["player_steam_id"]);
    return steamId ? { steamId, dayCount: integerValue(row["day_count"]) } : null;
  }).filter((player): player is CloudArchivePlayer => player !== null) : [];
  return {
    id,
    serverKey: stringValue(server["server_key"]) ?? "",
    serverName: stringValue(server["name"]) ?? stringValue(server["server_key"]) ?? "Unknown server",
    wipeKey: stringValue(item["wipe_key"]) ?? "",
    wipeStartedAtUtc: dateValue(item["wipe_started_at"]),
    firstObservedAtUtc: dateValue(item["first_observed_at"]),
    lastObservedAtUtc: dateValue(item["last_observed_at"]),
    playerCount: nullableIntegerValue(item["player_count"]),
    storedBytes: nullableIntegerValue(item["stored_bytes"]),
    players,
  };
}

function parseRestoreDay(value: unknown): RestoreDay | null {
  const item = record(value);
  const playerSteamId = stringValue(item["player_steam_id"]);
  const day = stringValue(item["day"]);
  const payload = parseDayPayload(item["payload"]);
  return playerSteamId && day && payload ? { playerSteamId, playerName: stringValue(item["player_name"]), day, payload } : null;
}

function parseDayPayload(value: unknown): CloudTrackerDayPayload | null {
  const payload = record(value);
  if (!Array.isArray(payload["observations"])) return null;
  const observations = payload["observations"].map((point) => {
    const item = record(point);
    const timestamp = stringValue(item["timestamp"]);
    if (!timestamp) return null;
    return {
      timestamp,
      x: numberOrNull(item["x"]),
      y: numberOrNull(item["y"]),
      state: parseState(item["state"]),
      location_type: parseLocation(item["location_type"]),
      location_name: stringOrNull(item["location_name"]),
      grid: stringOrNull(item["grid"]),
      event: item["event"] === "death" || item["event"] === "respawn" ? item["event"] : null,
    };
  }).filter((point): point is CloudTrackerDayPayload["observations"][number] => point !== null);
  return { schema_version: 1, generated_at: stringValue(payload["generated_at"]) ?? "", observation_sessions: Array.isArray(payload["observation_sessions"]) ? payload["observation_sessions"].filter((value): value is string => typeof value === "string") : [], observations };
}

function parseState(value: unknown): PlayerActivityState {
  switch (String(value ?? "").trim().toLowerCase()) {
    case "moving": return "moving";
    case "stationary": return "stationary";
    case "afk": return "afk";
    case "dead": return "dead";
    case "offline": return "offline";
    default: return "unknown";
  }
}

function parseLocation(value: unknown): PlayerObservation["locationType"] {
  const normalized = String(value ?? "").trim().toLowerCase();
  return normalized === "monument" || normalized === "base" || normalized === "open" ? normalized : "unknown";
}

function record(value: unknown): Record<string, unknown> { return typeof value === "object" && value !== null && !Array.isArray(value) ? value as Record<string, unknown> : {}; }
function stringValue(value: unknown): string | null { return typeof value === "string" && value ? value : null; }
function numberOrNull(value: unknown): number | null { return typeof value === "number" && Number.isFinite(value) ? value : null; }
function integerValue(value: unknown): number { return typeof value === "number" && Number.isFinite(value) ? Math.max(0, Math.trunc(value)) : 0; }
function nullableIntegerValue(value: unknown): number | null { return typeof value === "number" && Number.isFinite(value) ? Math.trunc(value) : null; }
function dateValue(value: unknown): Date | null { if (typeof value !== "string") return null; const parsed = new Date(value); return Number.isNaN(parsed.getTime()) ? null : parsed; }

function buildReplayPoints(observations: readonly PlayerObservation[]): WipeReplayPoint[] {
  const points: WipeReplayPoint[] = [];
  let previous: PlayerObservation | null = null;
  for (const observation of observations) {
    const displacement = previous?.x !== null && previous?.y !== null && observation.x !== null && observation.y !== null && previous ? Math.hypot(previous.x - observation.x, previous.y - observation.y) : 0;
    const continuity = Boolean(previous && previous.sessionId === observation.sessionId && observation.timestampUtc > previous.timestampUtc && secondsBetween(previous.timestampUtc, observation.timestampUtc) <= MAX_CONTINUITY_GAP_SECONDS && previous.isConnected && previous.snapshotValid && observation.isConnected && observation.snapshotValid);
    const state = !continuity && previous ? "unknown" : PlayerWipeTrackerEngine.classify(observation, displacement);
    const event = previous && !previous.dead && observation.dead ? "death" : previous && previous.dead && !observation.dead ? "respawn" : null;
    points.push({ timestampUtc: observation.timestampUtc, x: observation.x, y: observation.y, state, locationType: observation.locationType, locationName: observation.locationName, grid: observation.grid, event, sessionId: observation.sessionId });
    previous = observation;
  }
  return points;
}

export function buildWipeKey(serverKey: string, wipeTimeUtc: Date | null, mapIdentity: string | null): string {
  const normalized = wipeTimeUtc?.toISOString() ?? "unknown";
  return createHash("sha256").update(`${serverKey.trim()}|${normalized}|${mapIdentity?.trim() || "unknown"}`, "utf8").digest("hex");
}

function canTrack(steamId: string, ownSteamId: string, capabilities: PlayerWipeCapabilities): boolean {
  return /^\d{17}$/.test(steamId) && (steamId === ownSteamId || capabilities.canTrackTeam) && (steamId === ownSteamId || capabilities.maxTrackedPlayers > 1);
}
function dateKey(value: Date): string { return value.toISOString().slice(0, 10); }
function secondsBetween(start: Date, end: Date): number { return (end.getTime() - start.getTime()) / 1000; }
function stringOrNull(value: unknown): string | null { return typeof value === "string" ? value : null; }
function cryptoRandomId(): string { return `${Date.now().toString(36)}-${Math.random().toString(36).slice(2)}`; }
function delay(ms: number): Promise<void> { return new Promise((resolve) => setTimeout(resolve, ms)); }
