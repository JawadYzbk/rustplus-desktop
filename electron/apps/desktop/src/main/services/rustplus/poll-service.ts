/**
 * PollService — headless port of the legacy UI-timer polling cadences:
 *  - server status every 10 s via getInfo; failures feed ConnectionManager.recordStatusResult
 *    (5 consecutive → silent refresh there) and keep last known values on error (parity);
 *  - team list every 5 s via getTeamInfo (MainWindow.Team.Core.cs:253);
 *  - map markers every 2 s via getMapMarkers (MainWindow.Map.Markers.cs:601).
 * Loops tick only while the connection is up; all timing rides the injected clock.
 */
import { EventEmitter } from "node:events";
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

export type PollEvents =
  | { kind: "status"; status: ServerStatus }
  | { kind: "team"; team: Record<string, unknown> }
  | { kind: "markers"; markers: Record<string, unknown> };

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
  getTeamInfo: "teamInfo",
  getMapMarkers: "mapMarkers",
};

export class PollService extends EventEmitter {
  private running = false;

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
    void this.loop("status", this.intervals.statusMs, () => this.pollStatus());
    void this.loop("team", this.intervals.teamMs, () => this.pollTeam());
    void this.loop("markers", this.intervals.markersMs, () => this.pollMarkers());
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
      const status = toServerStatus(res[RESPONSE_KEY["getInfo"]!]);
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
    const markers = res[RESPONSE_KEY["getMapMarkers"]!];
    if (markers && typeof markers === "object") {
      this.emit("poll", { kind: "markers", markers } as PollEvents);
    }
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
