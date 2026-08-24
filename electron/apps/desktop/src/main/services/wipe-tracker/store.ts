import { appendFileSync, existsSync, mkdirSync, readdirSync, readFileSync, rmSync, statSync, writeFileSync } from "node:fs";
import { join } from "node:path";
import { parseObservation, serializeObservation, type PlayerObservation, type TrackerPersistedObservation, type TrackerWipeMap } from "./models.js";

/** Append-only JSONL storage isolated by server, wipe, and player. */
export class PlayerWipeTrackerStore {
  private readonly root: string;

  constructor(rootDirectory: string) {
    this.root = rootDirectory;
    mkdirSync(this.root, { recursive: true });
  }

  append(serverKey: string, wipeKey: string, steamId: string, item: TrackerPersistedObservation): boolean {
    try {
      const path = this.filePath(serverKey, wipeKey, steamId);
      mkdirSync(join(this.root, safe(serverKey), safe(wipeKey)), { recursive: true });
      appendFileSync(path, `${serializeObservation(item)}\n`, "utf8");
      return true;
    } catch {
      return false;
    }
  }

  appendAsync(serverKey: string, wipeKey: string, steamId: string, item: TrackerPersistedObservation): Promise<void> {
    if (!this.append(serverKey, wipeKey, steamId, item)) return Promise.reject(new Error("failed to append wipe tracker observation"));
    return Promise.resolve();
  }

  load(serverKey: string, wipeKey: string, steamId: string): TrackerPersistedObservation[] {
    const path = this.filePath(serverKey, wipeKey, steamId);
    if (!existsSync(path)) return [];
    const result: TrackerPersistedObservation[] = [];
    const seen = new Set<string>();
    for (const line of readFileSync(path, "utf8").split(/\r?\n/)) {
      const item = line.trim() ? parseObservation(line) : null;
      if (!item) continue;
      const key = `${item.kind}|${item.observation.sessionId}|${item.observation.timestampUtc.toISOString()}`;
      if (!seen.has(key)) {
        seen.add(key);
        result.push(item);
      }
    }
    return result.sort((a, b) => a.observation.timestampUtc.getTime() - b.observation.timestampUtc.getTime());
  }

  loadPlayerIds(serverKey: string, wipeKey: string): string[] {
    const directory = join(this.root, safe(serverKey), safe(wipeKey));
    if (!existsSync(directory)) return [];
    return readdirSync(directory)
      .filter((name) => name.endsWith(".jsonl"))
      .map((name) => name.slice(0, -6))
      .filter((name) => /^\d{1,20}$/.test(name));
  }

  storageBytes(): number { return totalBytes(this.root); }

  hasWipeMap(serverKey: string, wipeKey: string): boolean {
    const directory = join(this.root, safe(serverKey), safe(wipeKey));
    return existsSync(join(directory, "map.png")) && existsSync(join(directory, "map.json"));
  }

  saveWipeMap(serverKey: string, wipeKey: string, map: TrackerWipeMap): void {
    const directory = join(this.root, safe(serverKey), safe(wipeKey));
    mkdirSync(directory, { recursive: true });
    writeFileSync(join(directory, "map.png"), map.pngBytes);
    writeFileSync(join(directory, "map.json"), JSON.stringify({ worldSize: map.worldSize, worldRectX: map.worldRectX, worldRectY: map.worldRectY, worldRectWidth: map.worldRectWidth, worldRectHeight: map.worldRectHeight }));
  }

  loadWipeMap(serverKey: string, wipeKey: string): TrackerWipeMap | null {
    const directory = join(this.root, safe(serverKey), safe(wipeKey));
    try {
      if (!this.hasWipeMap(serverKey, wipeKey)) return null;
      const metadata = JSON.parse(readFileSync(join(directory, "map.json"), "utf8")) as Record<string, unknown>;
      const number = (key: string): number => typeof metadata[key] === "number" ? metadata[key] as number : NaN;
      const map = { pngBytes: new Uint8Array(readFileSync(join(directory, "map.png"))), worldSize: number("worldSize"), worldRectX: number("worldRectX"), worldRectY: number("worldRectY"), worldRectWidth: number("worldRectWidth"), worldRectHeight: number("worldRectHeight") };
      return Object.values(map).every((value) => value instanceof Uint8Array || (typeof value === "number" && Number.isFinite(value))) ? map : null;
    } catch { return null; }
  }

  deleteWipe(serverKey: string, wipeKey: string): void { rmSync(join(this.root, safe(serverKey), safe(wipeKey)), { recursive: true, force: true }); }

  deleteAll(): void {
    rmSync(this.root, { recursive: true, force: true });
    mkdirSync(this.root, { recursive: true });
  }

  private filePath(serverKey: string, wipeKey: string, steamId: string): string {
    return join(this.root, safe(serverKey), safe(wipeKey), `${safe(steamId)}.jsonl`);
  }
}

function safe(value: string): string {
  const result = value.replace(/[<>:"/\\|?*]/g, "_").trim();
  return (result || "unknown").slice(0, 120);
}

function totalBytes(path: string): number {
  if (!existsSync(path)) return 0;
  const stat = statSync(path);
  if (stat.isFile()) return stat.size;
  return readdirSync(path).reduce((sum, item) => sum + totalBytes(join(path, item)), 0);
}

export function observationsFromStore(store: PlayerWipeTrackerStore, serverKey: string, wipeKey: string, steamId: string): PlayerObservation[] {
  return store.load(serverKey, wipeKey, steamId).filter((item) => item.kind === "observation").map((item) => item.observation);
}
