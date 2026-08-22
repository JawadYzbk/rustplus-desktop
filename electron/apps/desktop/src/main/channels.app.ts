import { app } from "electron";
import type { HandlerMapOf } from "./ipc.js";
import { APP_VERSION } from "@rpd/shared";
import type { IpcChannels } from "@rpd/shared";
import { logger } from "./logger.js";

export interface AppChannelContext {
  smokeMode: boolean;
}

/** Handlers for the stage-2 `app/*` channels; the registry in @rpd/shared stays the single contract. */
export function buildAppHandlers(ctx: AppChannelContext): HandlerMapOf<IpcChannels> {
  return {
    // Literal keys (not appGetInfo.name) keep per-channel contextual handler types.
    "app/getInfo": () => ({
      version: APP_VERSION,
      electron: process.versions.electron ?? "unknown",
      chrome: process.versions.chrome ?? "unknown",
      node: process.versions.node ?? "unknown",
      platform: process.platform,
      locale: app.getLocale(),
      smokeMode: ctx.smokeMode,
    }),
    "app/logFromRenderer": ({ level, scope, message }) => {
      logger.log(level, `renderer/${scope}`, message);
    },
  };
}
