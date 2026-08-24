/**
 * PollService — headless port of the legacy UI-timer polling cadences:
 *  - server status every 10 s via getInfo; failures feed ConnectionManager.recordStatusResult
 *    (5 consecutive → silent refresh there) and keep last known values on error (parity);
 *  - team list every 5 s via getTeamInfo (MainWindow.Team.Core.cs:253);
 *  - map markers every 2 s via getMapMarkers (MainWindow.Map.Markers.cs:601).
 * Loops tick only while the connection is up; all timing rides the injected clock.
 */
import { EventEmitter } from "node:events";
import { Buffer } from "node:buffer";
import { rq } from "./protocol.js";
import type { Clock } from "./timing.js";
import { realClock } from "./timing.js";

export interface PollManagerLike {
  readonly isConnected: boolean;
  send(data: Record<string, unknown>, timeoutMs?: number): Promise<Record<string, unknown>>;
  recordStatusResult(ok: boolean): void;
}

export interface ServerStatus {
  players: number;
  maxPlayers: number;
  queue: number;
  timeString: string | null;
}

export interface MapMonument {
  x: number;
  y: number;
  token: string | null;
}

export interface MapSnapshot {
  width: number;
  height: number;
  worldSize: number;
  oceanMargin: number;
  imageBase64: string | null;
  monuments: MapMonument[];
}

export interface MapMarker {
  id: number;
  type: string;
  x: number;
  y: number;
  steamId: string | null;
  rotation: number | null;
  radius: number | null;
  alpha: number | null;
  name: string | null;
}

export type PollEvents =
  | { kind: "status"; status: ServerStatus }
  | { kind: "team"; team: Record<string, unknown> }
  | { kind: "map"; map: MapSnapshot }
  | { kind: "markers"; markers: MapMarker[] };

export const POLL_INTERVALS = {
  statusMs: 10_000,
  teamMs: 5_000,
  markersMs: 2_000,
} as const;

/** AppInfo payload → legacy ServerStatus shape (players>=0 gate preserved).
 * Response keys are the PROTO names (info/teamInfo/mapMarkers), not the request names. */
export function toServerStatus(info: unknown): ServerStatus | null {
  const i = (info ?? {}) as Record<string, unknown>;
  const num = (v: unknown): number => (typeof v === "number" ? v : -1);
  const players = num(i["players"]);
  if (!(players >= 0)) return null; // legacy: st != null && st.Players >= 0
  const firstNum = (...keys: string[]): number => {
    for (const k of keys) {
      const v = i[k];
      if (typeof v === "number") return v;
    }
    return 0;
  };
  return {
    players,
    maxPlayers: firstNum("maxPlayers"),
    queue: firstNum("queuedPlayers", "queue", "queueSize"),
    timeString: typeof i["time"] === "string" ? (i["time"] as string) : null,
  };
}

/** Game-time float → HH:MM — port of TryReadTimeHHMM's numeric branches (RustPlusClientReal L5219-5246). */
export function formatGameTime(t: unknown): string | null {
  if (typeof t !== "number" || !Number.isFinite(t)) return null;
  const pad = (n: number): string => String(n).padStart(2, "0");
  const toHHMM = (h: number, m: number): string => `${pad(((h % 24) + 24) % 24)}:${pad(((m % 60) + 60) % 60)}`;
  if (t >= 0 && t < 24) {
    let h = Math.floor(t);
    let m = Math.round((t - h) * 60);
    if (m === 60) {
      h = (h + 1) % 24;
      m = 0;
    }
    return toHHMM(h, m);
  }
  if (t >= 0 && t <= 1440) {
    const h = Math.floor(t / 60);
    const m = Math.round(t % 60);
    return toHHMM(h, m);
  }
  return null;
}

/** Request-name → proto AppResponse content key (asymmetry preserved deliberately). */
const RESPONSE_KEY: Record<string, string> = {
  getInfo: "info",
  getTime: "time",
  getMap: "map",
  getTeamInfo: "teamInfo",
  getMapMarkers: "mapMarkers",
};

const MARKER_TYPES = [
  "Undefined",
  "Player",
  "Explosion",
  "VendingMachine",
  "CH47",
  "CargoShip",
  "Crate",
  "GenericRadius",
  "PatrolHelicopter",
] as const;

function object(value: unknown): Record<string, unknown> {
  return value !== null && typeof value === "object" ? value as Record<string, unknown> : {};
}

function finite(value: unknown, fallback = 0): number {
  return typeof value === "number" && Number.isFinite(value) ? value : fallback;
}

function positive(value: unknown, fallback = 0): number {
  const result = finite(value, fallback);
  return result > 0 ? result : fallback;
}

function bytesToBase64(value: unknown): string | null {
  if (typeof value === "string") return value;
  if (value instanceof Uint8Array) return Buffer.from(value).toString("base64");
  if (Array.isArray(value) && value.every((item) => typeof item === "number")) {
    return Buffer.from(value as number[]).toString("base64");
  }
  const data = object(value).data;
  if (Array.isArray(data) && data.every((item) => typeof item === "number")) {
    return Buffer.from(data as number[]).toString("base64");
  }
  return null;
}

/** Normalize the raw AppMap message before it crosses the main/renderer boundary. */
export function toMapSnapshot(value: unknown, worldSize = 0): MapSnapshot | null {
  if (value === null || typeof value !== "object") return null;
  const raw = object(value);
  const monuments = Array.isArray(raw.monuments) ? raw.monuments.flatMap((value): MapMonument[] => {
    const monument = object(value);
    const x = finite(monument.x, Number.NaN);
    const y = finite(monument.y, Number.NaN);
    if (!Number.isFinite(x) || !Number.isFinite(y)) return [];
    return [{ x, y, token: typeof monument.token === "string" && monument.token ? monument.token : null }];
  }) : [];
  return {
    width: positive(raw.width),
    height: positive(raw.height),
    worldSize: positive(worldSize),
    oceanMargin: Math.max(0, finite(raw.oceanMargin)),
    imageBase64: bytesToBase64(raw.jpgImage),
    monuments,
  };
}

/** Keep the live marker stream small and stable for the renderer. */
export function toMapMarkers(value: unknown): MapMarker[] {
  const raw = object(value);
  if (!Array.isArray(raw.markers)) return [];
  return raw.markers.flatMap((value, index): MapMarker[] => {
    const marker = object(value);
    const typeValue = marker.type;
    const type = typeof typeValue === "number" ? MARKER_TYPES[typeValue] ?? "Undefined" : typeof typeValue === "string" ? typeValue : "Undefined";
    const x = finite(marker.x, Number.NaN);
    const y = finite(marker.y, Number.NaN);
    if (!Number.isFinite(x) || !Number.isFinite(y)) return [];
    const rawId = finite(marker.id, index);
    const id = rawId > 0 ? rawId : index;
    return [{
      id,
      type,
      x,
      y,
      steamId: marker.steamId === undefined || marker.steamId === null ? null : String(marker.steamId),
      rotation: typeof marker.rotation === "number" && Number.isFinite(marker.rotation) ? marker.rotation : null,
      radius: typeof marker.radius === "number" && Number.isFinite(marker.radius) ? marker.radius : null,
      alpha: typeof marker.alpha === "number" && Number.isFinite(marker.alpha) ? marker.alpha : null,
      name: typeof marker.name === "string" && marker.name ? marker.name : null,
    }];
  });
}

export class PollService extends EventEmitter {
  private running = false;
  private worldSize = 0;

  constructor(
    private readonly mgr: PollManagerLike,
    private readonly clock: Clock = realClock,
    private readonly intervals: { statusMs: number; teamMs: number; markersMs: number } = POLL_INTERVALS,
  ) {
    super();
  }

  start(): void {
    if (this.running) return;
    this.running = true;
    this.worldSize = 0;
    void this.loop("status", this.intervals.statusMs, () => this.pollStatus());
    void this.loop("team", this.intervals.teamMs, () => this.pollTeam());
    void this.loop("markers", this.intervals.markersMs, () => this.pollMarkers());
    // The map image is a connection snapshot; dynamic markers continue on their 2 s loop.
    void this.pollMap().catch(() => undefined);
  }

  stop(): void {
    this.running = false;
  }

  get isRunning(): boolean {
    return this.running;
  }

  private async loop(name: string, intervalMs: number, fn: () => Promise<void>): Promise<void> {
    // First pass runs immediately (legacy timers fire their first Tick right away too).
    while (this.running) {
      try {
        if (this.mgr.isConnected) await fn();
      } catch {
        /* keep last known values (parity) */
      }
      let remaining = intervalMs;
      while (remaining > 0 && this.running) {
        const step = Math.min(100, remaining);
        await this.clock.sleep(step);
        remaining -= step;
      }
    }
    void name;
  }

  private async pollStatus(): Promise<void> {
    try {
      // Legacy GetServerStatusAsync issues BOTH getInfo and getTime per tick.
      const res = await this.mgr.send(rq.getInfo());
      const timeRes = await this.mgr.send(rq.getTime()).catch(() => ({}) as Record<string, unknown>);
      const info = res[RESPONSE_KEY["getInfo"]!];
      const infoRecord = object(info);
      this.worldSize = positive(infoRecord.mapSize ?? infoRecord.worldSize, this.worldSize);
      const status = toServerStatus(info);
      if (!status) throw new Error("invalid status payload");
      if (!status.timeString) {
        const timePayload = (timeRes[RESPONSE_KEY["getTime"]!] ?? null) as { time?: unknown } | null;
        status.timeString = formatGameTime(timePayload?.time);
      }
      this.mgr.recordStatusResult(true);
      this.emit("poll", { kind: "status", status } satisfies PollEvents);
    } catch {
      this.mgr.recordStatusResult(false);
      // No event — last known values stay visible (legacy parity).
    }
  }

  private async pollTeam(): Promise<void> {
    const res = await this.mgr.send(rq.getTeamInfo());
    const team = res[RESPONSE_KEY["getTeamInfo"]!];
    if (team && typeof team === "object") {
      this.emit("poll", { kind: "team", team } as PollEvents);
    }
  }

  private async pollMarkers(): Promise<void> {
    const res = await this.mgr.send(rq.getMapMarkers());
    const markers = toMapMarkers(res[RESPONSE_KEY["getMapMarkers"]!]);
    this.emit("poll", { kind: "markers", markers } satisfies PollEvents);
  }

  private async pollMap(): Promise<void> {
    if (!this.running || !this.mgr.isConnected) return;
    const res = await this.mgr.send(rq.getMap());
    if (this.worldSize <= 0) {
      const infoRes = await this.mgr.send(rq.getInfo()).catch(() => ({}) as Record<string, unknown>);
      const info = object(infoRes[RESPONSE_KEY["getInfo"]!]);
      this.worldSize = positive(info.mapSize ?? info.worldSize, this.worldSize);
    }
    const map = toMapSnapshot(res[RESPONSE_KEY["getMap"]!], this.worldSize);
    if (map && this.running && this.mgr.isConnected) this.emit("poll", { kind: "map", map } satisfies PollEvents);
  }
}

/** Minimal profile contract for switching (matches ConnectionProfileRef). */
export interface SwitchProfile {
  host: string;
  port: number;
  steamId64: string;
  playerToken: string;
  UseFacepunchProxy?: boolean;
}

/** Switch servers: full teardown of the old connection (subscription/chat state dies with it),
 *  then connect the next profile. Legacy PerformConnectAsync semantics — no same-profile shortcut. */
export async function switchServer<M extends { disconnect(): Promise<unknown>; connect(p: SwitchProfile): Promise<unknown> }>(
  mgr: M,
  profile: SwitchProfile,
): Promise<unknown> {
  await mgr.disconnect();
  return mgr.connect(profile);
}
