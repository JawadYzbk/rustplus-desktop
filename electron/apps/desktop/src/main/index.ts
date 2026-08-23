/**
 * Main-process entry.
 *
 * Stage-2 scope (MIGRATION_PROGRESS stage 2): secure boot skeleton — single instance, hardened window,
 * typed IPC wiring, rotating file logger, smoke-verification mode. Feature services land per stage.
 */
import { app, BrowserWindow } from "electron";
import { join } from "node:path";
import { APP_NAME, APP_VERSION, ipcChannels } from "@rpd/shared";
import { logger } from "./logger.js";
import { createRegistrar } from "./ipc.js";
import { buildAppHandlers } from "./channels.app.js";
import { buildMigrationHandlers } from "./channels.migration.js";
import { LegacyMigrator, defaultLegacyRoots } from "./services/legacy-migrator.js";
import { SettingsStore } from "./stores/settings-store.js";
import { ProfilesStore } from "./stores/profiles-store.js";
import { SafeStorageSecretCodec } from "./stores/safe-storage-codec.js";
import {
  AlertTemplateStore,
  DeviceHotkeysStore,
  HotkeyOptionsStore,
  TrackedPlayersStore,
} from "./stores/legacy-stores.js";
import { TutorialProgressStore } from "./stores/tutorial-progress-store.js";
import { UiPrefsStore } from "./stores/ui-prefs-store.js";
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
  // Storage root fixed per ELECTRON_ARCHITECTURE §6 (%APPDATA%\RustPlusDesk-Electron) — independent of
  // the workspace package name that Electron would otherwise derive ("@rpd" in dev).
  app.setPath("userData", join(app.getPath("appData"), "RustPlusDesk-Electron"));

  logger.init();
  logger.info("app", `${APP_NAME} ${APP_VERSION} starting (dev=${isDev ? "1" : "0"} smoke=${isSmoke ? "1" : "0"})`);

  const userDataDir = app.getPath("userData");
  const storeLog = (scope: string) => (level: "warn" | "error", message: string) =>
    logger.log(level, scope, message);

  const uiPrefsStore = new UiPrefsStore(userDataDir, storeLog("store/ui-prefs"));
  const settingsStore = new SettingsStore(userDataDir, storeLog("store/settings"));
  const profilesStore = new ProfilesStore(userDataDir, new SafeStorageSecretCodec(), storeLog("store/profiles"));

  // %LOCALAPPDATA% has no first-class Electron path; env var is authoritative on Windows.
  const localAppData = process.env["LOCALAPPDATA"] ?? join(app.getPath("appData"), "..", "Local");
  const migrator = new LegacyMigrator(
    defaultLegacyRoots(app.getPath("appData"), localAppData),
    userDataDir,
    {
      settings: settingsStore,
      profiles: profilesStore,
      hotkeys: new DeviceHotkeysStore(userDataDir, storeLog("store/hotkeys")),
      hotkeyOptions: new HotkeyOptionsStore(userDataDir, storeLog("store/hotkey-options")),
      alerts: new AlertTemplateStore(userDataDir, storeLog("store/alerts")),
      trackedPlayers: new TrackedPlayersStore(userDataDir, storeLog("store/tracked-players")),
      tutorials: new TutorialProgressStore(userDataDir, storeLog("store/tutorials")),
    },
    (level, message) => logger.log(level, "migrator", message),
  );

  const registrar = createRegistrar(ipcChannels);
  registrar.register({
    ...buildAppHandlers({ smokeMode: isSmoke, uiPrefs: uiPrefsStore }),
    ...buildMigrationHandlers({ migrator }),
  });

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
