/**
 * Handlers for migrate/scan + migrate/run (stage-3 migrator M3).
 * Scan is strictly read-only; run writes only into the NEW root, never touching legacy files.
 */
import type { HandlerMapOf } from "./ipc.js";
import type { IpcChannels } from "@rpd/shared";
import type { LegacyMigrator } from "./services/legacy-migrator.js";

export interface MigrationChannelContext {
  migrator: LegacyMigrator;
}

/** IPC handlers for the `migrate/*` channels; the registry in @rpd/shared stays the contract. */
export function buildMigrationHandlers(ctx: MigrationChannelContext): Pick<HandlerMapOf<IpcChannels>, "migrate/scan" | "migrate/run"> {
  return {
    "migrate/scan": () => ctx.migrator.scan(),
    "migrate/run": () => ctx.migrator.run(),
  };
}
