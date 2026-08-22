/**
 * Main-window factory.
 *
 * Parity anchors (audit UI_SHELL §1–2, §6): default 1750×814 / min 1000×700; custom titlebar implemented via
 * hidden native bar + Windows titleBarOverlay so our React header owns the top strip while keeping native
 * window controls; dark background matches AppBg token (#0F1013).
 */
import { BrowserWindow, shell } from "electron";
import { join } from "node:path";

let mainWindow: BrowserWindow | null = null;

export function getMainWindow(): BrowserWindow | null {
  return mainWindow && !mainWindow.isDestroyed() ? mainWindow : null;
}

export function createMainWindow(): BrowserWindow {
  if (mainWindow && !mainWindow.isDestroyed()) return mainWindow;

  const win = new BrowserWindow({
    width: 1750,
    height: 814,
    minWidth: 1000,
    minHeight: 700,
    show: false,
    titleBarStyle: "hidden",
    titleBarOverlay: process.platform === "win32" ? { height: 48 } : false,
    backgroundColor: "#0F1013",
    webPreferences: {
      preload: join(__dirname, "../preload/index.js"),
      contextIsolation: true,
      nodeIntegration: false,
      sandbox: true,
      spellcheck: false,
    },
  });

  mainWindow = win;

  win.once("ready-to-show", () => win.show());

  win.on("closed", () => {
    if (mainWindow === win) mainWindow = null;
  });

  // Hardening: no popups, no arbitrary navigation. External links go through the OS browser allowlist-style.
  win.webContents.setWindowOpenHandler(({ url }) => {
    void shell.openExternal(url);
    return { action: "deny" };
  });
  win.webContents.on("will-navigate", (event, url) => {
    const devUrl = process.env["ELECTRON_RENDERER_URL"];
    const allowed = devUrl ? url.startsWith(devUrl) : url.startsWith("file://");
    if (!allowed) event.preventDefault();
  });

  const devUrl = process.env["ELECTRON_RENDERER_URL"];
  if (devUrl) {
    void win.loadURL(devUrl);
  } else {
    void win.loadFile(join(__dirname, "../renderer/index.html"));
  }

  return win;
}
