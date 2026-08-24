/** Connection IPC handlers — thin adapters over the ConnectionManager instance.
 * The manager is injected (not imported) so the module stays testable without Electron. */
import type { HandlerMapOf } from "./ipc.js";
import type { ConnSnapshotDto, IpcChannels } from "@rpd/shared";

export interface ConnManagerLike {
  connect(profile: {
    host: string;
    port: number;
    steamId64: string;
    playerToken: string;
    UseFacepunchProxy?: boolean;
  }): Promise<ConnSnapshotDto>;
  disconnect(): Promise<ConnSnapshotDto>;
  snapshot(): ConnSnapshotDto;
}

export interface StoredProfileLike {
  Name: string;
  Host: string;
  Port: number;
  SteamId64: string;
  UseFacepunchProxy?: boolean;
}

export interface ProfileSecretStoreLike {
  list(): StoredProfileLike[];
  matchKey(profile: StoredProfileLike): string;
  tokenFor(matchKey: string): string;
}

export function connectionHandlers(
  manager: ConnManagerLike,
  profiles: ProfileSecretStoreLike,
): Pick<HandlerMapOf<IpcChannels>, "conn/connect" | "conn/connectProfile" | "conn/disconnect" | "conn/status"> {
  return {
    "conn/connect": async (req: { host: string; port: number; steamId64: string; playerToken: string; useProxy?: boolean }) => {
      const snap = await manager.connect({
        host: req.host,
        port: req.port,
        steamId64: req.steamId64,
        playerToken: req.playerToken,
        UseFacepunchProxy: req.useProxy === true ? true : undefined,
      });
      return snap;
    },
    "conn/connectProfile": async (req: { matchKey: string; useProxy?: boolean }) => {
      const profile = profiles.list().find((candidate) => profiles.matchKey(candidate) === req.matchKey);
      if (!profile) throw new Error("server profile not found");
      const playerToken = profiles.tokenFor(req.matchKey);
      if (!playerToken) throw new Error("server profile has no Rust+ player token; pair it again");
      return manager.connect({
        host: profile.Host,
        port: profile.Port,
        steamId64: profile.SteamId64,
        playerToken,
        UseFacepunchProxy: req.useProxy === true || profile.UseFacepunchProxy === true ? true : undefined,
      });
    },
    "conn/disconnect": async () => {
      await manager.disconnect();
      return manager.snapshot();
    },
    "conn/status": async () => manager.snapshot(),
  };
}
