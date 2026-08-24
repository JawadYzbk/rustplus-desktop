import { wipeDeleteAllCloud, wipeDeleteCloudArchive, wipeGetCloudArchives, wipeGetMap, wipeGetPlayer, wipeGetStatus, wipeRestoreCloudArchive } from "@rpd/shared";
import type { PlayerWipeTrackerService } from "./services/wipe-tracker/service.js";

type WipePlayerDto = ReturnType<typeof wipeGetStatus["response"]["parse"]>["players"][number];
type WipeArchiveDto = ReturnType<typeof wipeGetCloudArchives["response"]["parse"]>["archives"][number];

export function buildWipeHandlers(tracker: PlayerWipeTrackerService): {
  "wipe/getStatus": () => ReturnType<typeof wipeGetStatus["response"]["parse"]>;
  "wipe/getPlayer": (request: ReturnType<typeof wipeGetPlayer["request"]["parse"]>) => ReturnType<typeof wipeGetPlayer["response"]["parse"]>;
  "wipe/getMap": () => ReturnType<typeof wipeGetMap["response"]["parse"]>;
  "wipe/getCloudArchives": () => ReturnType<typeof wipeGetCloudArchives["response"]["parse"]> | Promise<ReturnType<typeof wipeGetCloudArchives["response"]["parse"]>>;
  "wipe/restoreCloudArchive": (request: ReturnType<typeof wipeRestoreCloudArchive["request"]["parse"]>) => ReturnType<typeof wipeRestoreCloudArchive["response"]["parse"]> | Promise<ReturnType<typeof wipeRestoreCloudArchive["response"]["parse"]>>;
  "wipe/deleteCloudArchive": (request: ReturnType<typeof wipeDeleteCloudArchive["request"]["parse"]>) => ReturnType<typeof wipeDeleteCloudArchive["response"]["parse"]> | Promise<ReturnType<typeof wipeDeleteCloudArchive["response"]["parse"]>>;
  "wipe/deleteAllCloud": () => ReturnType<typeof wipeDeleteAllCloud["response"]["parse"]> | Promise<ReturnType<typeof wipeDeleteAllCloud["response"]["parse"]>>;
} {
  return {
    "wipe/getStatus": () => ({
      serverKey: tracker.currentServerKey,
      wipeKey: tracker.currentWipeKey,
      sessionId: tracker.currentSessionId,
      players: tracker.getPlayers().map(toPlayerDto),
    }),
    "wipe/getPlayer": ({ steamId }) => {
      const player = tracker.getPlayer(steamId);
      return { player: player ? toPlayerDto(player) : null };
    },
    "wipe/getMap": () => ({ map: toMapDto(tracker.loadCurrentWipeMap()) }),
    "wipe/getCloudArchives": async () => ({ archives: (await tracker.getCloudArchives()).map(toArchiveDto) }),
    "wipe/restoreCloudArchive": async ({ archiveId }) => tracker.restoreCloudArchive(archiveId),
    "wipe/deleteCloudArchive": async ({ archiveId }) => ({ deleted: await tracker.deleteCloudArchive(archiveId) }),
    "wipe/deleteAllCloud": async () => ({ deleted: await tracker.deleteAllCloud() }),
  };
}

function toPlayerDto(player: NonNullable<ReturnType<PlayerWipeTrackerService["getPlayer"]>>): WipePlayerDto {
  if (!player) throw new Error("wipe player is missing");
  return {
    steamId: player.steamId,
    name: player.name,
    observationCount: player.observationCount,
    summary: {
      ...player.summary,
      monumentVisits: player.summary.monumentVisits.map((visit) => ({ ...visit, startUtc: visit.startUtc.toISOString(), endUtc: visit.endUtc.toISOString() })),
    },
    insights: {
      ...player.insights,
      firstSeenUtc: player.insights.firstSeenUtc?.toISOString() ?? null,
      lastSeenUtc: player.insights.lastSeenUtc?.toISOString() ?? null,
      longestBlindGapStartUtc: player.insights.longestBlindGapStartUtc?.toISOString() ?? null,
      currentAsOfUtc: player.insights.currentAsOfUtc?.toISOString() ?? null,
    },
    observations: player.observations.map((point) => ({ ...point, timestampUtc: point.timestampUtc.toISOString() })),
    segments: player.segments.map((segment) => ({ ...segment, startUtc: segment.startUtc.toISOString(), endUtc: segment.endUtc.toISOString() })),
  };
}

function toMapDto(map: ReturnType<PlayerWipeTrackerService["loadCurrentWipeMap"]>): ReturnType<typeof wipeGetMap["response"]["parse"]>["map"] {
  if (!map) return null;
  const [imageWidth, imageHeight] = pngDimensions(map.pngBytes);
  return { pngBase64: Buffer.from(map.pngBytes).toString("base64"), imageWidth, imageHeight, worldSize: map.worldSize, worldRectX: map.worldRectX, worldRectY: map.worldRectY, worldRectWidth: map.worldRectWidth, worldRectHeight: map.worldRectHeight };
}

function pngDimensions(bytes: Uint8Array): [number, number] {
  if (bytes.length >= 24 && bytes[0] === 0x89 && bytes[1] === 0x50 && bytes[2] === 0x4e && bytes[3] === 0x47) return [bytes[16]! * 0x1000000 + bytes[17]! * 0x10000 + bytes[18]! * 0x100 + bytes[19]!, bytes[20]! * 0x1000000 + bytes[21]! * 0x10000 + bytes[22]! * 0x100 + bytes[23]!];
  return [1, 1];
}

function toArchiveDto(archive: Awaited<ReturnType<PlayerWipeTrackerService["getCloudArchives"]>>[number]): WipeArchiveDto {
  return {
    ...archive,
    wipeStartedAtUtc: archive.wipeStartedAtUtc?.toISOString() ?? null,
    firstObservedAtUtc: archive.firstObservedAtUtc?.toISOString() ?? null,
    lastObservedAtUtc: archive.lastObservedAtUtc?.toISOString() ?? null,
  };
}
