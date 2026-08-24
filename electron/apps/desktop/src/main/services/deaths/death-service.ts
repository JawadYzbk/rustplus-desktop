import { appendFileSync, existsSync, mkdirSync, readFileSync, unlinkSync } from "node:fs";
import { basename, join } from "node:path";

export type DeathLocationType = "monument" | "base" | "open";

export interface DeathRecord {
  steamId: string;
  name: string | null;
  deathTime: number;
  spawnTime: number | null;
  x: number | null;
  y: number | null;
  grid: string | null;
  locationType: DeathLocationType;
  locationName: string | null;
}

export interface DeathEntry {
  victim: string;
  diedAt: number;
  spawnAt: number | null;
  grid: string | null;
  type: DeathLocationType;
  location: string;
}

export interface DeathStatsSummary {
  total: number;
  victims: number;
  avgSurvival: string;
  longestSurvival: string;
  peakHour: string;
  deadliestPlace: string;
  deadliestGrid: string;
  byArea: Array<{ name: string; type: DeathLocationType; deaths: number; percent: number }>;
  byVictim: Array<{ victim: string; deaths: number; avgSurvival: string }>;
  byLocation: Array<{ location: string; type: DeathLocationType; deaths: number }>;
  recent: Array<{ victim: string; type: DeathLocationType; location: string; grid: string; died: string }>;
  deathsPerDay: Array<{ day: string; count: number }>;
}

export interface DeathStatsFilters {
  search?: string;
  player?: string;
  type?: "all" | DeathLocationType;
  range?: "all" | "24h" | "7d";
}

interface RawDeath {
  name?: string;
  died_at?: number;
  spawn_at?: number | null;
  grid?: string | null;
  location_type?: string;
  location_name?: string | null;
}

interface TeamMember {
  steamId: string;
  name: string | null;
  online: boolean;
  dead: boolean;
  x: number | null;
  y: number | null;
  deathTime: number | null;
  spawnTime: number | null;
  grid: string | null;
  locationType: DeathLocationType | null;
  locationName: string | null;
}

interface BaseNote {
  x: number;
  y: number;
  label: string;
}

const BASE_RADIUS = 90;
const HOME_ICON = 2;

/** C# DeathTracker parity: first snapshot establishes history; later edges emit once. */
export class DeathTracker {
  private readonly lastDeathTime = new Map<string, number>();
  private readonly spawnObserved = new Map<string, number>();
  private readonly wasDead = new Map<string, boolean>();
  private baselineEstablished = false;

  reset(): void {
    this.lastDeathTime.clear();
    this.spawnObserved.clear();
    this.wasDead.clear();
    this.baselineEstablished = false;
  }

  observe(members: readonly TeamMember[], classify: (member: TeamMember) => Pick<DeathRecord, "grid" | "locationType" | "locationName">, now = Math.floor(Date.now() / 1000)): DeathRecord[] {
    const deaths: DeathRecord[] = [];
    for (const member of members) {
      if (!member.steamId) continue;
      const previousDead = this.wasDead.get(member.steamId) === true;
      if (previousDead && !member.dead) this.spawnObserved.set(member.steamId, member.spawnTime ?? now);
      this.wasDead.set(member.steamId, member.dead);

      const deathTime = member.deathTime && member.deathTime > 0 ? member.deathTime : now;
      const isNewDeath = member.deathTime && member.deathTime > 0
        ? member.deathTime > (this.lastDeathTime.get(member.steamId) ?? 0)
        : member.dead && !previousDead;
      if (!isNewDeath) continue;
      this.lastDeathTime.set(member.steamId, deathTime);
      if (!this.baselineEstablished) continue;

      const location = classify(member);
      deaths.push({
        steamId: member.steamId,
        name: member.name,
        deathTime,
        spawnTime: member.spawnTime && member.spawnTime > 0 ? member.spawnTime : this.spawnObserved.get(member.steamId) ?? null,
        x: member.x,
        y: member.y,
        ...location,
      });
    }
    this.baselineEstablished = true;
    return deaths;
  }
}

/** Local JSONL store; malformed historical lines are ignored like the WPF reader. */
export class DeathLogStore {
  constructor(private readonly root: string, private readonly legacyRoot: string | null = null) {}

  pathFor(serverKey: string): string {
    const safe = serverKey.replace(/[<>:"/\\|?*\u0000-\u001f]/g, "_").trim() || "server";
    return join(this.root, `${safe}.jsonl`);
  }

  append(serverKey: string, death: DeathRecord): void {
    try {
      mkdirSync(this.root, { recursive: true });
      appendFileSync(this.pathFor(serverKey), `${JSON.stringify({
        steam_id: death.steamId,
        name: death.name,
        died_at: death.deathTime,
        spawn_at: death.spawnTime,
        x: death.x,
        y: death.y,
        grid: death.grid,
        location_type: death.locationType,
        location_name: death.locationName,
      })}\n`, "utf8");
    } catch {
      // A local write failure must not break the team poll.
    }
  }

  loadEntries(serverKey: string | null): DeathEntry[] {
    if (!serverKey) return [];
    return this.readRaw(serverKey).map((raw) => ({
      victim: raw.name || "Unknown",
      diedAt: numberOr(raw.died_at, 0),
      spawnAt: numberOrNull(raw.spawn_at),
      grid: stringOrNull(raw.grid),
      type: normalizeType(raw.location_type),
      location: raw.location_name?.trim() || areaName(normalizeType(raw.location_type)),
    }));
  }

  clear(serverKey: string | null): void {
    if (!serverKey) return;
    for (const path of [this.pathFor(serverKey), this.legacyPathFor(serverKey)].filter((value): value is string => value !== null)) {
      try {
        if (existsSync(path)) unlinkSync(path);
      } catch {
        // A failed delete must not crash the stats view.
      }
    }
  }

  private legacyPathFor(serverKey: string): string | null {
    if (!this.legacyRoot) return null;
    return join(this.legacyRoot, basename(this.pathFor(serverKey)));
  }

  private readRaw(serverKey: string): RawDeath[] {
    const paths = [this.pathFor(serverKey), this.legacyPathFor(serverKey)].filter((value): value is string => value !== null);
    const result: RawDeath[] = [];
    for (const path of paths) {
      if (!existsSync(path)) continue;
      let lines = "";
      try {
        lines = readFileSync(path, "utf8");
      } catch {
        continue;
      }
      for (const line of lines.split(/\r?\n/)) {
        if (!line.trim()) continue;
        try {
          const value = JSON.parse(line) as RawDeath;
          if (value && typeof value === "object") result.push(value);
        } catch {
          // Keep one bad line from hiding the rest of the log.
        }
      }
    }
    return result;
  }
}

export class DeathService {
  private readonly tracker = new DeathTracker();
  private serverKey: string | null = null;

  constructor(private readonly store: DeathLogStore, private readonly worldSize: () => number | null = () => null) {}

  get currentServerKey(): string | null { return this.serverKey; }

  observeTeam(serverKey: string, team: unknown): DeathRecord[] {
    if (serverKey !== this.serverKey) {
      this.tracker.reset();
      this.serverKey = serverKey;
    }
    const rawTeam = asRecord(team);
    const members = arrayOf(rawTeam?.members).map(normalizeMember).filter((member): member is TeamMember => member !== null);
    const bases = baseNotes(rawTeam);
    const mapWorldSize = numberOrNull(rawTeam?.worldSize ?? rawTeam?.mapSize) ?? this.worldSize();
    const deaths = this.tracker.observe(members, (member) => {
      const base = nearestBase(bases, member.x, member.y);
      const type = base ? "base" : member.locationType ?? "open";
      return {
        locationType: type,
        locationName: base?.label ?? member.locationName,
        grid: member.grid ?? gridLabel(member.x, member.y, mapWorldSize),
      };
    });
    for (const death of deaths) this.store.append(serverKey, death);
    return deaths;
  }

  disconnect(): void {
    this.tracker.reset();
  }

  stats(filters: DeathStatsFilters = {}): { serverKey: string | null; players: string[]; summary: DeathStatsSummary } {
    const allEntries = this.store.loadEntries(this.serverKey);
    const query = filters.search?.trim().toLocaleLowerCase();
    const cutoff = filters.range === "24h" ? Math.floor(Date.now() / 1000) - 86400 : filters.range === "7d" ? Math.floor(Date.now() / 1000) - 604800 : null;
    const entries = allEntries.filter((entry) => {
      if (query && !`${entry.victim} ${entry.location} ${entry.grid ?? ""}`.toLocaleLowerCase().includes(query)) return false;
      if (filters.player && filters.player !== "all" && entry.victim !== filters.player) return false;
      if (filters.type && filters.type !== "all" && entry.type !== filters.type) return false;
      return cutoff === null || entry.diedAt >= cutoff;
    });
    return { serverKey: this.serverKey, players: [...new Set(allEntries.map((entry) => entry.victim))].sort((a, b) => a.localeCompare(b)), summary: summarizeDeaths(entries) };
  }

  clear(): boolean {
    this.store.clear(this.serverKey);
    return true;
  }
}

export function summarizeDeaths(entries: readonly DeathEntry[]): DeathStatsSummary {
  if (entries.length === 0) return emptySummary();

  const byVictimMap = new Map<string, DeathEntry[]>();
  const byLocationMap = new Map<string, { type: DeathLocationType; deaths: number }>();
  const byAreaMap = new Map<DeathLocationType, number>();
  for (const entry of entries) {
    byVictimMap.set(entry.victim, [...(byVictimMap.get(entry.victim) ?? []), entry]);
    const place = byLocationMap.get(entry.location) ?? { type: entry.type, deaths: 0 };
    place.deaths += 1;
    byLocationMap.set(entry.location, place);
    byAreaMap.set(entry.type, (byAreaMap.get(entry.type) ?? 0) + 1);
  }
  const byVictim = [...byVictimMap.entries()].map(([victim, values]) => ({ victim, deaths: values.length, avgSurvival: averageSurvival(values) })).sort(descCountThenText("deaths", "victim"));
  const byLocation = [...byLocationMap.entries()].map(([location, value]) => ({ location, ...value })).sort(descCountThenText("deaths", "location"));
  const byArea = [...byAreaMap.entries()].map(([type, deaths]) => ({ name: areaName(type), type, deaths, percent: Math.round(100 * deaths / entries.length) })).sort(descCountThenText("deaths", "name"));
  const grids = entries.filter((entry) => entry.grid && entry.grid !== "off-grid");
  const deadliestGrid = topCount(grids.map((entry) => entry.grid as string));
  const peakHourCounts = new Map<number, number>();
  for (const entry of entries) {
    const hour = new Date(entry.diedAt * 1000).getHours();
    peakHourCounts.set(hour, (peakHourCounts.get(hour) ?? 0) + 1);
  }
  const peakHour = [...peakHourCounts.entries()].sort((a, b) => b[1] - a[1] || a[0] - b[0])[0];
  const survivals = entries.filter((entry) => entry.spawnAt !== null && entry.diedAt > entry.spawnAt).map((entry) => entry.diedAt - (entry.spawnAt as number));
  const longest = survivals.length ? Math.max(...survivals) : 0;

  const today = new Date();
  const dayCounts = new Map(entries.map((entry) => [localDateKey(entry.diedAt), 0]));
  for (const entry of entries) dayCounts.set(localDateKey(entry.diedAt), (dayCounts.get(localDateKey(entry.diedAt)) ?? 0) + 1);
  const deathsPerDay = Array.from({ length: 14 }, (_, index) => {
    const day = new Date(today.getFullYear(), today.getMonth(), today.getDate() - 13 + index);
    const key = localDateKey(Math.floor(day.getTime() / 1000));
    return { day: key, count: dayCounts.get(key) ?? 0 };
  });

  return {
    total: entries.length,
    victims: byVictim.length,
    avgSurvival: averageSurvival(entries),
    longestSurvival: longest > 0 ? formatDuration(longest) : "—",
    peakHour: peakHour ? `${String(peakHour[0]).padStart(2, "0")}:00 (${peakHour[1]})` : "—",
    deadliestPlace: byLocation[0]?.location ?? "—",
    deadliestGrid: deadliestGrid ? `${deadliestGrid.key} (${deadliestGrid.count})` : "—",
    byArea,
    byVictim,
    byLocation,
    recent: [...entries].sort((a, b) => b.diedAt - a.diedAt).slice(0, 100).map((entry) => ({ victim: entry.victim, type: entry.type, location: entry.location, grid: entry.grid ?? "—", died: new Date(entry.diedAt * 1000).toLocaleString() })),
    deathsPerDay,
  };
}

function emptySummary(): DeathStatsSummary {
  return { total: 0, victims: 0, avgSurvival: "—", longestSurvival: "—", peakHour: "—", deadliestPlace: "—", deadliestGrid: "—", byArea: [], byVictim: [], byLocation: [], recent: [], deathsPerDay: [] };
}

function normalizeMember(value: unknown): TeamMember | null {
  const raw = asRecord(value);
  if (!raw) return null;
  const steamId = String(raw.steamId ?? raw.steamId64 ?? raw.userId ?? raw.playerId ?? "");
  if (!/^\d+$/.test(steamId) || steamId === "0") return null;
  const position = asRecord(raw.position ?? raw.pos);
  let dead = booleanOrUndefined(raw.dead ?? raw.isDead);
  if (!dead && raw.alive !== undefined) dead = !Boolean(raw.alive);
  const lifeState = numberOrNull(raw.lifeState ?? raw.lifestate);
  if (!dead && (lifeState === 1 || lifeState === 2)) dead = true;
  if (!dead && Boolean(raw.wounded ?? raw.isWounded)) dead = true;
  const locationType = raw.locationType === "monument" || raw.locationType === "base" || raw.locationType === "open" ? raw.locationType : null;
  return {
    steamId,
    name: typeof raw.name === "string" ? raw.name : typeof raw.displayName === "string" ? raw.displayName : null,
    online: Boolean(raw.online ?? raw.isOnline),
    dead: dead === true,
    x: numberOrNull(raw.x ?? position?.x),
    y: numberOrNull(raw.y ?? position?.y),
    deathTime: positiveOrNull(raw.deathTime),
    spawnTime: positiveOrNull(raw.spawnTime),
    grid: typeof raw.grid === "string" && raw.grid ? raw.grid : null,
    locationType,
    locationName: typeof raw.locationName === "string" ? raw.locationName : null,
  };
}

function baseNotes(team: Record<string, unknown> | null): BaseNote[] {
  const notes = [...arrayOf(team?.leaderMapNotes), ...arrayOf(team?.mapNotes)];
  return notes.flatMap((value) => {
    const note = asRecord(value);
    if (!note || numberOr(note.icon ?? note.type, -1) !== HOME_ICON) return [];
    const x = numberOrNull(note.x);
    const y = numberOrNull(note.y);
    return x === null || y === null ? [] : [{ x, y, label: typeof note.label === "string" && note.label.trim() ? note.label : "Base" }];
  });
}

function nearestBase(notes: readonly BaseNote[], x: number | null, y: number | null): BaseNote | null {
  if (x === null || y === null) return null;
  let best: BaseNote | null = null;
  let bestDistance = Number.POSITIVE_INFINITY;
  for (const note of notes) {
    const distance = Math.hypot(x - note.x, y - note.y);
    if (distance <= BASE_RADIUS && distance < bestDistance) { best = note; bestDistance = distance; }
  }
  return best;
}

function gridLabel(x: number | null, y: number | null, worldSize: number | null): string | null {
  if (x === null || y === null || !worldSize || worldSize <= 0) return null;
  const cells = Math.max(1, Math.ceil(worldSize / 150));
  const col = Math.max(0, Math.min(cells - 1, Math.floor(x / 150)));
  const row = Math.max(0, Math.min(cells - 1, Math.floor((worldSize - y) / 150)));
  return `${columnLabel(col)}${row}`;
}

function columnLabel(index: number): string {
  let value = index + 1;
  let result = "";
  while (value > 0) { value -= 1; result = String.fromCharCode(65 + value % 26) + result; value = Math.floor(value / 26); }
  return result;
}

function normalizeType(value: string | undefined): DeathLocationType { return value === "monument" || value === "base" ? value : "open"; }
function areaName(type: DeathLocationType): string { return type === "monument" ? "Monument" : type === "base" ? "Base" : "Open"; }
function numberOr(value: unknown, fallback: number): number { const parsed = numberOrNull(value); return parsed === null ? fallback : parsed; }
function numberOrNull(value: unknown): number | null { const parsed = typeof value === "number" ? value : typeof value === "string" && value.trim() ? Number(value) : NaN; return Number.isFinite(parsed) ? parsed : null; }
function positiveOrNull(value: unknown): number | null { const parsed = numberOrNull(value); return parsed !== null && parsed > 0 ? Math.floor(parsed) : null; }
function stringOrNull(value: unknown): string | null { return typeof value === "string" && value ? value : null; }
function booleanOrUndefined(value: unknown): boolean | undefined { return typeof value === "boolean" ? value : typeof value === "string" ? value.toLowerCase() === "true" ? true : value.toLowerCase() === "false" ? false : undefined : undefined; }
function asRecord(value: unknown): Record<string, unknown> | null { return value !== null && typeof value === "object" && !Array.isArray(value) ? value as Record<string, unknown> : null; }
function arrayOf(value: unknown): unknown[] { return Array.isArray(value) ? value : []; }
function localDateKey(unixSeconds: number): string { const date = new Date(unixSeconds * 1000); return `${date.getFullYear()}-${String(date.getMonth() + 1).padStart(2, "0")}-${String(date.getDate()).padStart(2, "0")}`; }
function formatDuration(seconds: number): string { const total = Math.floor(seconds); if (total < 60) return `${total}s`; const minutes = Math.floor(total / 60); const remainder = total % 60; if (minutes < 60) return `${minutes}m ${remainder}s`; return `${Math.floor(minutes / 60)}h ${minutes % 60}m`; }
function averageSurvival(entries: readonly DeathEntry[]): string { const values = entries.filter((entry) => entry.spawnAt !== null && entry.diedAt > entry.spawnAt).map((entry) => entry.diedAt - (entry.spawnAt as number)); return values.length ? formatDuration(Math.floor(values.reduce((sum, value) => sum + value, 0) / values.length)) : "—"; }
function topCount(values: readonly string[]): { key: string; count: number } | null { const counts = new Map<string, number>(); for (const value of values) counts.set(value, (counts.get(value) ?? 0) + 1); const top = [...counts.entries()].sort((a, b) => b[1] - a[1])[0]; return top ? { key: top[0], count: top[1] } : null; }
function descCountThenText<T extends { deaths: number }>(countKey: "deaths", textKey: keyof T): (a: T, b: T) => number { return (a, b) => b[countKey] - a[countKey] || String(a[textKey]).localeCompare(String(b[textKey])); }
