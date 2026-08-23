import { describe, expect, it } from "vitest";
import { readFileSync } from "node:fs";
import { fileURLToPath } from "node:url";
import { dirname, join } from "node:path";
import {
  appGetInfo,
  APP_NAME,
  APP_VERSION,
  ipcChannels,
} from "../src/index.js";

const here = dirname(fileURLToPath(import.meta.url));

function lookup(name: string) {
  const def = (ipcChannels as Record<string, unknown>)[name];
  if (!def) throw new Error(`channel ${name} not registered`);
  return def as typeof appGetInfo;
}

describe("ipc registry", () => {
  it("exposes only snake-less dotted channel names", () => {
    for (const name of Object.keys(ipcChannels)) {
      expect(name).toMatch(/^[a-z][a-zA-Z0-9]*(\/[a-z][a-zA-Z0-9]*)+$/);
    }
  });

  it("validates a well-formed logFromRenderer payload", () => {
    const def = lookup("app/logFromRenderer");
    const parsed = def.request.safeParse({
      level: "info",
      scope: "shell",
      message: "hello",
    });
    expect(parsed.success).toBe(true);
  });

  it("rejects an out-of-contract logFromRenderer payload", () => {
    const def = lookup("app/logFromRenderer");
    const parsed = def.request.safeParse({
      level: "verbose",
      scope: "",
      message: "x".repeat(5000),
    });
    expect(parsed.success).toBe(false);
  });

  it("appGetInfo response schema accepts a realistic payload", () => {
    const parsed = appGetInfo.response.safeParse({
      version: APP_VERSION,
      electron: "33.0.0",
      chrome: "130.0.0.0",
      node: "20.18.0",
      platform: "win32",
      locale: "en-US",
      smokeMode: false,
    });
    expect(parsed.success).toBe(true);
  });
});

describe("app constants", () => {
  it("APP_VERSION matches the workspace package.json version", () => {
    const pkg = JSON.parse(readFileSync(join(here, "../package.json"), "utf8")) as { version: string };
    // shared package sits at packages/shared → workspace root is three levels up from tests/
    const root = join(here, "../../../package.json");
    const rootPkg = JSON.parse(readFileSync(root, "utf8")) as { name: string; version: string };
    expect(rootPkg.name).toBe("rustplusdesk-electron");
    expect(APP_VERSION).toBe(pkg.version);
    expect(APP_VERSION).toBe(rootPkg.version);
  });

  it("APP_NAME matches the legacy product name", () => {
    expect(APP_NAME).toBe("RustPlusDesk");
  });
});
