import { describe, expect, it } from "vitest";
import { normalizeTeamSnapshot, useConnectionStore } from "../src/renderer/src/stores/connection.js";

describe("normalizeTeamSnapshot", () => {
  it("keeps the Rust+ member fields needed by the team surface", () => {
    const team = normalizeTeamSnapshot({
      leaderSteamId: "76561198000000001",
      members: [
        { steamId: "76561198000000001", name: "Ada", online: true, x: 12.5, y: -4 },
        { steamId64: "76561198000000002", displayName: "Bea", isOnline: false, isDead: true, position: { x: 3, y: 9 } },
        { name: "invalid" },
      ],
    }, 123);

    expect(team).toEqual({
      leaderSteamId: "76561198000000001",
      receivedAt: 123,
      members: [
        { steamId: "76561198000000001", name: "Ada", online: true, dead: false, x: 12.5, y: -4 },
        { steamId: "76561198000000002", name: "Bea", online: false, dead: true, x: 3, y: 9 },
      ],
    });
  });

  it("returns a safe empty snapshot for missing payloads", () => {
    expect(normalizeTeamSnapshot(null, 7)).toEqual({ leaderSteamId: null, members: [], receivedAt: 7 });
  });

  it("keeps queued players from status pushes and defaults an omitted queue to zero", () => {
    useConnectionStore.getState().applyPush("poll", { kind: "status", status: { players: 42, maxPlayers: 200, queuedPlayers: 7, timeString: "12:34" } });
    expect(useConnectionStore.getState().status).toEqual({ players: 42, maxPlayers: 200, queuedPlayers: 7, timeString: "12:34" });
    useConnectionStore.getState().applyPush("poll", { kind: "status", status: { players: 1, maxPlayers: 10 } });
    expect(useConnectionStore.getState().status?.queuedPlayers).toBe(0);
  });
});
