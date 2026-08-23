/** Handlers for backup/* + reset/* channels (stage-3 closeout). */
import type { HandlerMapOf } from "./ipc.js";
import type { IpcChannels } from "@rpd/shared";
import type { BackupService } from "./services/backup-service.js";
import { performReset, type ResetContext } from "./services/reset-service.js";

export interface BackupChannelContext {
  backup: BackupService;
  reset: ResetContext;
}

export function buildBackupHandlers(ctx: BackupChannelContext): Pick<
  HandlerMapOf<IpcChannels>,
  "backup/create" | "backup/restore" | "reset/perform"
> {
  return {
    "backup/create": ({ password }) => ctx.backup.create(password ?? null),
    "backup/restore": ({ path, password }) => ctx.backup.restore(path, password ?? null),
    "reset/perform": ({ targets }) => performReset(ctx.reset, targets),
  };
}
