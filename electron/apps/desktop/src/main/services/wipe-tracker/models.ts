export type PlayerActivityState = "moving" | "stationary" | "afk" | "dead" | "offline" | "unknown";
export type TrackerLocationType = "monument" | "base" | "open" | "unknown";

export interface PlayerObservation {
  steamId: string;
  name: string;
  timestampUtc: Date;
  sessionId: string;
  isConnected: boolean;
  snapshotValid: boolean;
  online: boolean;
  dead: boolean;
  afk: boolean;
  x: number | null;
  y: number | null;
  locationType: TrackerLocationType;
  locationName: string | null;
  grid: string | null;
  spawnTime: number | null;
  deathTime: number | null;
}

export interface TrackerSegment {
  startUtc: Date;
  endUtc: Date;
  state: PlayerActivityState;
  locationType: TrackerLocationType;
  locationName: string | null;
  grid: string | null;
  startX: number | null;
  startY: number | null;
  endX: number | null;
  endY: number | null;
  sessionId: string;
}

export interface MonumentVisit {
  name: string;
  startUtc: Date;
  endUtc: Date;
  entryX: number | null;
  entryY: number | null;
  exitX: number | null;
  exitY: number | null;
}

export interface TrackerSummary {
  coverageSeconds: number;
  unknownSeconds: number;
  movingSeconds: number;
  stationarySeconds: number;
  afkSeconds: number;
  deadSeconds: number;
  offlineSeconds: number;
  estimatedDistance: number;
  deaths: number;
  monumentVisits: MonumentVisit[];
}

export interface TrackerInsights {
  firstSeenUtc: Date | null;
  lastSeenUtc: Date | null;
  sessionCount: number;
  topMonument: string | null;
  topMonumentSeconds: number;
  topMonumentVisits: number;
  longestBlindGapSeconds: number;
  longestBlindGapStartUtc: Date | null;
  peakHourLocal: number | null;
  peakHourActiveSeconds: number;
  currentState: PlayerActivityState;
  currentLocationType: TrackerLocationType;
  currentLocationName: string | null;
  currentGrid: string | null;
  currentAsOfUtc: Date | null;
  isLikelyOnline: boolean;
}

export interface TrackerWipeMap {
  pngBytes: Uint8Array;
  worldSize: number;
  worldRectX: number;
  worldRectY: number;
  worldRectWidth: number;
  worldRectHeight: number;
}

export interface TrackerMapProjection {
  viewWidth: number;
  viewHeight: number;
  imageWidth: number;
  imageHeight: number;
  worldRectX: number;
  worldRectY: number;
  worldRectWidth: number;
  worldRectHeight: number;
  worldSize: number;
}

export function projectMapPoint(projection: TrackerMapProjection, worldX: number, worldY: number): { x: number; y: number } {
  if (projection.viewWidth <= 0 || projection.viewHeight <= 0 || projection.imageWidth <= 0 || projection.imageHeight <= 0 || projection.worldRectWidth <= 0 || projection.worldRectHeight <= 0 || projection.worldSize <= 0) return { x: 0, y: 0 };
  const scale = Math.min(projection.viewWidth / projection.imageWidth, projection.viewHeight / projection.imageHeight);
  const imageLeft = (projection.viewWidth - projection.imageWidth * scale) / 2;
  const imageTop = (projection.viewHeight - projection.imageHeight * scale) / 2;
  const sourceX = projection.worldRectX + worldX / projection.worldSize * projection.worldRectWidth;
  const sourceY = projection.worldRectY + (1 - worldY / projection.worldSize) * projection.worldRectHeight;
  return { x: imageLeft + sourceX * scale, y: imageTop + sourceY * scale };
}

export interface TrackerPersistedObservation {
  schemaVersion: number;
  kind: "observation";
  observation: PlayerObservation;
}

export interface CloudTrackerObservation {
  timestamp: string;
  x: number | null;
  y: number | null;
  state: PlayerActivityState;
  location_type: TrackerLocationType;
  location_name: string | null;
  grid: string | null;
  event: "death" | "respawn" | null;
}

export interface CloudTrackerDayPayload {
  schema_version: 1;
  generated_at: string;
  observation_sessions: string[];
  observations: CloudTrackerObservation[];
}

export interface CloudDayUploadRequest {
  server_key: string;
  wipe_key: string;
  wipe_started_at: string | null;
  player_steam_id: string;
  player_name: string | null;
  day: string;
  schema_version: 1;
  payload: CloudTrackerDayPayload;
  checksum: string;
}

export function serializeObservation(item: TrackerPersistedObservation): string {
  return JSON.stringify({
    schemaVersion: item.schemaVersion,
    kind: item.kind,
    observation: { ...item.observation, timestampUtc: item.observation.timestampUtc.toISOString() },
  });
}

export function parseObservation(line: string): TrackerPersistedObservation | null {
  try {
    const raw = JSON.parse(line) as Record<string, unknown>;
    const observation = raw.observation as Record<string, unknown> | undefined;
    if (raw.schemaVersion !== 1 || raw.kind !== "observation" || !observation || typeof observation.timestampUtc !== "string") return null;
    const timestamp = new Date(observation.timestampUtc);
    if (Number.isNaN(timestamp.getTime())) return null;
    return {
      schemaVersion: 1,
      kind: "observation",
      observation: {
        steamId: String(observation.steamId ?? ""),
        name: String(observation.name ?? ""),
        timestampUtc: timestamp,
        sessionId: String(observation.sessionId ?? ""),
        isConnected: observation.isConnected === true,
        snapshotValid: observation.snapshotValid === true,
        online: observation.online === true,
        dead: observation.dead === true,
        afk: observation.afk === true,
        x: numberOrNull(observation.x),
        y: numberOrNull(observation.y),
        locationType: locationOrUnknown(observation.locationType),
        locationName: stringOrNull(observation.locationName),
        grid: stringOrNull(observation.grid),
        spawnTime: numberOrNull(observation.spawnTime),
        deathTime: numberOrNull(observation.deathTime),
      },
    };
  } catch {
    return null;
  }
}

function numberOrNull(value: unknown): number | null {
  return typeof value === "number" && Number.isFinite(value) ? value : null;
}

function stringOrNull(value: unknown): string | null {
  return typeof value === "string" ? value : null;
}

function locationOrUnknown(value: unknown): TrackerLocationType {
  return value === "monument" || value === "base" || value === "open" ? value : "unknown";
}
