/**
 * Main-process entry.
 *
 * Stage-2 scope (MIGRATION_PROGRESS stage 2): secure boot skeleton — single instance, hardened window,
 * typed IPC wiring, rotating file logger, smoke-verification mode. Feature services land per stage.
 */
import { app, BrowserWindow } from "electron";
import { APP_NAME, APP_VERSION, ipcChannels } from "@rpd/shared";
import { logger } from "./logger.js";
import { createRegistrar } from "./ipc.js";
import { buildAppHandlers } from "./channels.app.js";
import { createMainWindow, getMainWindow } from "./window.js";

const isSmoke = process.env["RPD_SMOKE"] === "1";
const isDev = Boolean(process.env["ELECTRON_RENDERER_URL"]);

if (!app.requestSingleInstanceLock()) {
  // Second instance exits; the first instance receives `second-instance` (parity with the C# mutex +
  // named-pipe SHOWUI forwarding, audit UI_SHELL §6).
  app.quit();
} else {
  app.on("second-instance", () => {
    const win = getMainWindow();
    if (win) {
      if (win.isMinimized()) win.restore();
      win.show();
      win.focus();
    }
  });

  app.on("window-all-closed", () => {
    app.quit();
  });

  app.on("activate", () => {
    if (BrowserWindow.getAllWindows().length === 0) createMainWindow();
  });

  void app.whenReady().then(bootstrap);
}

function bootstrap(): void {
  logger.init();
  logger.info("app", `${APP_NAME} ${APP_VERSION} starting (dev=${isDev ? "1" : "0"} smoke=${isSmoke ? "1" : "0"})`);

  const registrar = createRegistrar(ipcChannels);
  registrar.register(buildAppHandlers({ smokeMode: isSmoke }));

  const win = createMainWindow();

  if (isSmoke) {
    const failsafe = setTimeout(() => {
      logger.error("smoke", "smoke run timed out before renderer finished loading");
      app.exit(1);
    }, 20_000);
    win.webContents.once("did-finish-load", () => {
      clearTimeout(failsafe);
      logger.info("smoke", "RPD_SMOKE_OK");
      setTimeout(() => app.exit(0), 250);
    });
  }
}
