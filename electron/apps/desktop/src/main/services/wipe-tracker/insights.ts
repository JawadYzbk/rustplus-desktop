import { PlayerWipeTrackerEngine } from "./engine.js";
import type { PlayerObservation, TrackerInsights, TrackerSegment, TrackerSummary } from "./models.js";

/** Pure glance-value derivation ported from TrackerInsightsBuilder.cs. */
export function buildInsights(observations: readonly PlayerObservation[], segments: readonly TrackerSegment[], summary: TrackerSummary, nowUtc = new Date()): TrackerInsights {
  if (observations.length === 0) return emptyInsights();
  const ordered = [...observations].sort((a, b) => a.timestampUtc.getTime() - b.timestampUtc.getTime());
  const first = ordered[0]!;
  const last = ordered.at(-1)!;
  const monumentGroups = new Map<string, { seconds: number; visits: number }>();
  for (const visit of summary.monumentVisits) {
    const current = monumentGroups.get(visit.name) ?? { seconds: 0, visits: 0 };
    current.seconds += secondsBetween(visit.startUtc, visit.endUtc);
    current.visits += 1;
    monumentGroups.set(visit.name, current);
  }
  const topMonument = [...monumentGroups.entries()].sort((a, b) => b[1].seconds - a[1].seconds)[0];
  let longestBlindGapSeconds = 0;
  let longestBlindGapStartUtc: Date | null = null;
  for (const segment of segments) {
    if (segment.state !== "unknown") continue;
    const seconds = secondsBetween(segment.startUtc, segment.endUtc);
    if (seconds > longestBlindGapSeconds) {
      longestBlindGapSeconds = seconds;
      longestBlindGapStartUtc = segment.startUtc;
    }
  }
  const hourBuckets = Array.from({ length: 24 }, () => 0);
  for (const segment of segments) {
    if (!(segment.state === "moving" || segment.state === "stationary" || segment.state === "afk")) continue;
    let cursor = new Date(segment.startUtc.getTime());
    const end = new Date(segment.endUtc.getTime());
    while (cursor < end) {
      const nextHour = new Date(cursor);
      nextHour.setMinutes(0, 0, 0);
      nextHour.setHours(nextHour.getHours() + 1);
      const sliceEnd = nextHour < end ? nextHour : end;
      const hour = cursor.getHours();
      hourBuckets[hour] = (hourBuckets[hour] ?? 0) + secondsBetween(cursor, sliceEnd);
      cursor = sliceEnd;
    }
  }
  const peakHourActiveSeconds = Math.max(...hourBuckets);
  const peakHourLocal = peakHourActiveSeconds > 0 ? hourBuckets.indexOf(peakHourActiveSeconds) : null;
  const currentState = PlayerWipeTrackerEngine.classify(last);
  return {
    firstSeenUtc: first.timestampUtc,
    lastSeenUtc: last.timestampUtc,
    sessionCount: new Set(ordered.map((item) => item.sessionId)).size,
    topMonument: topMonument?.[0] ?? null,
    topMonumentSeconds: topMonument?.[1].seconds ?? 0,
    topMonumentVisits: topMonument?.[1].visits ?? 0,
    longestBlindGapSeconds,
    longestBlindGapStartUtc,
    peakHourLocal,
    peakHourActiveSeconds,
    currentState,
    currentLocationType: last.locationType,
    currentLocationName: last.locationName,
    currentGrid: last.grid,
    currentAsOfUtc: last.timestampUtc,
    isLikelyOnline: last.isConnected && last.online && !last.dead && secondsBetween(last.timestampUtc, nowUtc) <= 180,
  };
}

function emptyInsights(): TrackerInsights {
  return { firstSeenUtc: null, lastSeenUtc: null, sessionCount: 0, topMonument: null, topMonumentSeconds: 0, topMonumentVisits: 0, longestBlindGapSeconds: 0, longestBlindGapStartUtc: null, peakHourLocal: null, peakHourActiveSeconds: 0, currentState: "unknown", currentLocationType: "unknown", currentLocationName: null, currentGrid: null, currentAsOfUtc: null, isLikelyOnline: false };
}

function secondsBetween(start: Date, end: Date): number { return (end.getTime() - start.getTime()) / 1000; }
