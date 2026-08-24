import { describe, expect, it, vi } from "vitest";
import type { ConnSnapshotDto } from "@rpd/shared";
import { connectionHandlers } from "../src/main/channels.connection.js";

const snapshot: ConnSnapshotDto = {
  connected: true,
  activeProxy: "direct",
  host: "rust.example",
  port: 28082,
  consecutiveTimeouts: 0,
  teamChatPrimed: true,
  clanChatPrimed: true,
};

describe("profile-scoped connection IPC", () => {
  it("resolves the token in main and returns only the connection snapshot", async () => {
    const connect = vi.fn(async (profile: { playerToken: string }) => {
      expect(profile.playerToken).toBe("secret-token");
      return snapshot;
    });
    const handlers = connectionHandlers({
      connect,
      disconnect: async () => snapshot,
      snapshot: () => snapshot,
    }, {
      list: () => [{ Name: "Home", Host: "rust.example", Port: 28082, SteamId64: "76561198000000001" }],
      matchKey: (profile) => `${profile.Host}:${profile.Port}|${profile.SteamId64}`,
      tokenFor: () => "secret-token",
    });

    const result = await handlers["conn/connectProfile"]({ matchKey: "rust.example:28082|76561198000000001" });
    expect(result).toEqual(snapshot);
    expect(JSON.stringify(result)).not.toContain("secret-token");
    expect(connect).toHaveBeenCalledOnce();
  });
});
