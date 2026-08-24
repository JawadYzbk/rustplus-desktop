import type {
  MonumentVisit,
  PlayerActivityState,
  PlayerObservation,
  TrackerSegment,
  TrackerSummary,
} from "./models.js";

export const PERSIST_DISTANCE_WORLD_UNITS = 10;
export const MONUMENT_OUTSIDE_OBSERVATIONS_TO_EXIT = 2;
export const MINIMUM_MONUMENT_VISIT_SECONDS = 30;
export const REENTRY_MERGE_SECONDS = 60;
export const HEARTBEAT_SECONDS = 60;
export const MAX_CONTINUITY_GAP_SECONDS = 15;

/** Pure observation classifier and segment builder ported from PlayerWipeTrackerEngine.cs. */
export class PlayerWipeTrackerEngine {
  private readonly segments: TrackerSegment[] = [];
  private readonly visits: MonumentVisit[] = [];
  private last: PlayerObservation | null = null;
  private state: PlayerActivityState = "unknown";
  private stateStartUtc = new Date(0);
  private stateStartObservation: PlayerObservation | null = null;
  private activeMonument: string | null = null;
  private monumentStartUtc = new Date(0);
  private monumentEntryX: number | null = null;
  private monumentEntryY: number | null = null;
  private outsideMonumentObservations = 0;

  get lastObservation(): PlayerObservation | null { return this.last; }
  get estimatedDistance(): number { return this.distance; }
  get deaths(): number { return this.deathCount; }
  getSegments(): readonly TrackerSegment[] { return this.segments; }
  getMonumentVisits(): readonly MonumentVisit[] { return this.visits; }

  private distance = 0;
  private deathCount = 0;

  observe(input: PlayerObservation): boolean {
    const observation = { ...input, timestampUtc: new Date(input.timestampUtc.getTime()) };
    if (!observation.steamId || !observation.sessionId) return false;
    if (!this.last) {
      this.last = observation;
      this.state = classify(observation);
      this.stateStartUtc = observation.timestampUtc;
      this.stateStartObservation = observation;
      this.updateMonument(observation);
      return true;
    }

    if (observation.timestampUtc <= this.last.timestampUtc) return false;
    const previous = this.last;
    const gapSeconds = secondsBetween(previous.timestampUtc, observation.timestampUtc);
    const continuity = previous.sessionId === observation.sessionId && gapSeconds <= MAX_CONTINUITY_GAP_SECONDS &&
      previous.isConnected && previous.snapshotValid && observation.isConnected && observation.snapshotValid;
    if (!continuity) {
      this.closeState(previous.timestampUtc);
      this.segments.push({
        startUtc: previous.timestampUtc,
        endUtc: observation.timestampUtc,
        state: "unknown",
        locationType: "unknown",
        locationName: null,
        grid: null,
        startX: previous.x,
        startY: previous.y,
        endX: observation.x,
        endY: observation.y,
        sessionId: observation.sessionId,
      });
      this.state = classify(observation);
      this.stateStartUtc = observation.timestampUtc;
      this.stateStartObservation = observation;
      this.outsideMonumentObservations = 0;
      this.updateMonument(observation);
      this.last = observation;
      return true;
    }

    if (!previous.dead && observation.dead) this.deathCount += 1;
    const displacement = distanceFromLastPersisted(previous, observation);
    const nextState = PlayerWipeTrackerEngine.classify(observation, displacement);
    const stateChanged = nextState !== this.state || observation.locationType !== this.stateStartObservation?.locationType ||
      observation.locationName !== this.stateStartObservation?.locationName;
    if (stateChanged) {
      this.closeState(observation.timestampUtc);
      this.state = nextState;
      this.stateStartUtc = observation.timestampUtc;
      this.stateStartObservation = observation;
    }
    if (canMeasureDistance(previous, observation, gapSeconds)) this.distance += distance(previous, observation);
    this.updateMonument(observation);
    this.last = observation;
    return stateChanged || secondsBetween(this.stateStartUtc, observation.timestampUtc) >= HEARTBEAT_SECONDS || displacement >= PERSIST_DISTANCE_WORLD_UNITS;
  }

  endSession(timestampUtc: Date): void {
    if (!this.last) return;
    const end = timestampUtc > this.last.timestampUtc ? timestampUtc : this.last.timestampUtc;
    if (end <= this.last.timestampUtc) return;
    this.closeState(this.last.timestampUtc);
    this.segments.push({ startUtc: this.last.timestampUtc, endUtc: end, state: "unknown", locationType: "unknown", locationName: null, grid: null, startX: this.last.x, startY: this.last.y, endX: null, endY: null, sessionId: this.last.sessionId });
    this.closeMonument(end, this.last.x, this.last.y);
    this.last = null;
  }

  summarize(): TrackerSummary {
    const durations = new Map<PlayerActivityState, number>();
    for (const segment of this.segments) durations.set(segment.state, (durations.get(segment.state) ?? 0) + secondsBetween(segment.startUtc, segment.endUtc));
    const seconds = (state: PlayerActivityState): number => durations.get(state) ?? 0;
    return {
      coverageSeconds: seconds("moving") + seconds("stationary") + seconds("afk") + seconds("dead") + seconds("offline"),
      unknownSeconds: seconds("unknown"),
      movingSeconds: seconds("moving"),
      stationarySeconds: seconds("stationary"),
      afkSeconds: seconds("afk"),
      deadSeconds: seconds("dead"),
      offlineSeconds: seconds("offline"),
      estimatedDistance: this.distance,
      deaths: this.deathCount,
      monumentVisits: [...this.visits],
    };
  }

  static classify(observation: PlayerObservation, displacement = 0): PlayerActivityState {
    const state = classify(observation);
    return state === "stationary" && displacement >= PERSIST_DISTANCE_WORLD_UNITS ? "moving" : state;
  }

  private closeState(endUtc: Date): void {
    if (!this.stateStartObservation || endUtc <= this.stateStartUtc) return;
    const last = this.last ?? this.stateStartObservation;
    this.segments.push({ startUtc: this.stateStartUtc, endUtc, state: this.state, locationType: this.stateStartObservation.locationType, locationName: this.stateStartObservation.locationName, grid: this.stateStartObservation.grid, startX: this.stateStartObservation.x, startY: this.stateStartObservation.y, endX: last.x, endY: last.y, sessionId: this.stateStartObservation.sessionId });
  }

  private updateMonument(observation: PlayerObservation): void {
    if (!observation.isConnected || !observation.snapshotValid) return;
    if (observation.locationType === "monument" && observation.locationName?.trim()) {
      if (!this.activeMonument) {
        this.activeMonument = observation.locationName;
        this.monumentStartUtc = observation.timestampUtc;
        this.monumentEntryX = observation.x;
        this.monumentEntryY = observation.y;
      } else if (this.activeMonument !== observation.locationName) {
        this.closeMonument(observation.timestampUtc, observation.x, observation.y);
        this.activeMonument = observation.locationName;
        this.monumentStartUtc = observation.timestampUtc;
        this.monumentEntryX = observation.x;
        this.monumentEntryY = observation.y;
      }
      this.outsideMonumentObservations = 0;
      return;
    }
    if (!this.activeMonument) return;
    this.outsideMonumentObservations += 1;
    if (this.outsideMonumentObservations >= MONUMENT_OUTSIDE_OBSERVATIONS_TO_EXIT) this.closeMonument(observation.timestampUtc, observation.x, observation.y);
  }

  private closeMonument(endUtc: Date, exitX: number | null, exitY: number | null): void {
    if (!this.activeMonument) return;
    if (secondsBetween(this.monumentStartUtc, endUtc) >= MINIMUM_MONUMENT_VISIT_SECONDS) {
      const visit: MonumentVisit = { name: this.activeMonument, startUtc: this.monumentStartUtc, endUtc, entryX: this.monumentEntryX, entryY: this.monumentEntryY, exitX, exitY };
      const previous = this.visits.at(-1);
      if (previous && previous.name === visit.name && secondsBetween(previous.endUtc, visit.startUtc) <= REENTRY_MERGE_SECONDS) {
        this.visits[this.visits.length - 1] = { ...previous, endUtc: visit.endUtc, exitX: visit.exitX, exitY: visit.exitY };
      } else this.visits.push(visit);
    }
    this.activeMonument = null;
    this.outsideMonumentObservations = 0;
  }
}

export function classify(observation: PlayerObservation): PlayerActivityState {
  if (!observation.isConnected || !observation.snapshotValid) return "unknown";
  if (!observation.online) return "offline";
  if (observation.dead) return "dead";
  if (observation.afk) return "afk";
  return "stationary";
}

function secondsBetween(start: Date, end: Date): number { return (end.getTime() - start.getTime()) / 1000; }
function distance(a: PlayerObservation, b: PlayerObservation): number { return Math.hypot(a.x! - b.x!, a.y! - b.y!); }
function distanceFromLastPersisted(a: PlayerObservation, b: PlayerObservation): number { return a.x === null || a.y === null || b.x === null || b.y === null ? 0 : distance(a, b); }
function canMeasureDistance(a: PlayerObservation, b: PlayerObservation, gapSeconds: number): boolean {
  if (a.sessionId !== b.sessionId || !a.snapshotValid || !b.snapshotValid || !a.isConnected || !b.isConnected || a.x === null || a.y === null || b.x === null || b.y === null || a.dead !== b.dead || a.deathTime !== b.deathTime || a.spawnTime !== b.spawnTime || gapSeconds <= 0 || gapSeconds > MAX_CONTINUITY_GAP_SECONDS) return false;
  return distance(a, b) <= Math.max(100, gapSeconds * 35);
}
