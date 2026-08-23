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

export function connectionHandlers(
  manager: ConnManagerLike,
): Pick<HandlerMapOf<IpcChannels>, "conn/connect" | "conn/disconnect" | "conn/status"> {
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
    "conn/disconnect": async () => {
      await manager.disconnect();
      return manager.snapshot();
    },
    "conn/status": async () => manager.snapshot(),
  };
}
