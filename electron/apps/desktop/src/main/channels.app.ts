import { app } from "electron";
import type { HandlerMapOf } from "./ipc.js";
import type { IpcChannels, UiPrefs } from "@rpd/shared";
import { APP_VERSION } from "@rpd/shared";
import { logger } from "./logger.js";

export interface AppChannelContext {
  smokeMode: boolean;
  /** Persisted shell-preference store (ui-prefs.json). */
  uiPrefs: {
    get(): UiPrefs;
    set(patch: Partial<UiPrefs>): UiPrefs;
  };
}

/** IPC handlers for the `app/*` and `uiPrefs/*` channels; the registry in @rpd/shared stays the contract. */
export function buildAppHandlers(ctx: AppChannelContext): HandlerMapOf<IpcChannels> {
  return {
    // Literal keys (not def.name) keep per-channel contextual handler types.
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
    "uiPrefs/get": () => ctx.uiPrefs.get(),
    "uiPrefs/set": (patch) => ctx.uiPrefs.set(patch),
  };
}
