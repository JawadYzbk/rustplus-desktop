/**
 * Contract tests for the typed IPC registry — no Electron runtime required (fake registrar).
 * Exit criterion for stage 2: "IPC contract tests green".
 */
import { describe, expect, it, vi } from "vitest";
import { ipcChannels, type ConnSnapshotDto } from "@rpd/shared";
import { registerIpcHandlers, type IpcRegistrar } from "../src/main/ipc.js";

const connSnap = (over: Partial<ConnSnapshotDto> = {}): ConnSnapshotDto => ({
  connected: false,
  activeProxy: null,
  host: null,
  port: null,
  consecutiveTimeouts: 0,
  teamChatPrimed: false,
  clanChatPrimed: false,
  ...over,
});

function fakeRegistrar(): IpcRegistrar & { handlers: Map<string, (raw: unknown) => Promise<unknown>> } {
  const handlers = new Map<string, (raw: unknown) => Promise<unknown>>();
  return {
    handlers,
    handle(channel, listener) {
      if (handlers.has(channel)) throw new Error(`duplicate registration for ${channel}`);
      handlers.set(channel, listener);
    },
  };
}

const baseHandlers = () => ({
  "app/getInfo": () => ({
    version: "8.1.0",
    electron: "33.0.0",
    chrome: "130.0.0.0",
    node: "20.18.0",
    platform: "win32",
    locale: "en-US",
    smokeMode: false,
  }),
  "app/logFromRenderer": () => undefined,
  "cloud/login": () => ({
    signedIn: true as const,
    user: { id: "u", steamId: null, name: null, displayName: null, email: "x@example.com", providers: ["email"], hasPassword: true },
  }),
  "cloud/bootstrap": () => ({ signedIn: false, user: null, capabilities: null, error: null }),
  "cloud/logout": () => ({ signedIn: false as const }),
  "wipe/getStatus": () => ({ serverKey: null, wipeKey: null, sessionId: null, players: [] }),
  "wipe/getPlayer": () => ({ player: null }),
  "wipe/getMap": () => ({ map: null }),
  "wipe/getCloudArchives": () => ({ archives: [] }),
  "wipe/restoreCloudArchive": () => ({ archiveId: "a", players: 0, days: 0, observations: 0, isCurrentWipe: false }),
  "wipe/deleteCloudArchive": () => ({ deleted: true }),
  "wipe/deleteAllCloud": () => ({ deleted: 0 }),
  "deaths/getStats": () => ({ serverKey: null, players: [], summary: { total: 0, victims: 0, avgSurvival: "—", longestSurvival: "—", peakHour: "—", deadliestPlace: "—", deadliestGrid: "—", byArea: [], byVictim: [], byLocation: [], recent: [], deathsPerDay: [] } }),
  "deaths/clear": () => ({ cleared: true }),
  "settings/getWipe": () => ({ enabled: false, cloudBackupEnabled: false }),
  "settings/setWipe": () => ({ enabled: true, cloudBackupEnabled: false }),
  "uiPrefs/get": () => ({ sidebarPinned: true, sidebarWidth: 420 }),
  "uiPrefs/set": (patch: { sidebarPinned?: boolean; sidebarWidth?: number }) => ({
    sidebarPinned: patch.sidebarPinned ?? true,
    sidebarWidth: patch.sidebarWidth ?? 420,
  }),
  "migrate/scan": () => ({ roots: [], sources: [] }),
  "migrate/run": () => ({ startedAt: "", finishedAt: "", rows: [] }),
  "backup/create": () => ({ path: "", bytes: 0, encrypted: false }),
  "backup/restore": () => ({ restored: [], skipped: [] }),
  "reset/perform": ({ targets }: { targets: string[] }) => ({ performed: targets }),
  "conn/connect": () =>
    connSnap({
      connected: true,
      activeProxy: "direct",
      host: "1.2.3.4",
      port: 28082,
      teamChatPrimed: true,
      clanChatPrimed: true,
    }),
  "conn/disconnect": () => connSnap(),
  "conn/status": () => connSnap(),
  "profile/list": () => ({
    profiles: [
      { matchKey: "h:1|s", name: "n", host: "h", port: 1, steamId64: "s", deviceCount: 0 },
    ],
  }),
  "profile/pair": () => ({ activated: true, profile: { matchKey: "h:1|s", name: "n", host: "h", port: 1, steamId64: "76561198000000001", deviceCount: 0 } }),
  "profile/getDevices": () => ({ devices: [], found: false }),
  "profile/saveDevices": () => ({ saved: false }),
  "profile/activate": () => ({ activated: true }),
  "profile/exportDevices": () => ({ saved: false, canceled: true, path: null, bytes: 0 }),
  "profile/importPreview": () => ({ canceled: true, path: null, candidates: [] }),
  "profile/applyImport": () => ({ saved: false, imported: 0 }),
  "profile/deleteDevice": () => ({ removed: false, reason: "notFound" as const }),
  "deviceAutomation/getRules": () => ({ found: false, isActive: false, rules: [] }),
  "deviceAutomation/saveRules": () => ({ saved: true }),
  "raid/getData": () => ({ sources: [], targets: [] }),
  "raid/calculate": () => ({ methods: [], recommended: null, combination: [], resources: [], items: [] }),
  "recycler/getData": () => ({ items: [] }),
  "recycler/calculate": () => ({ outputs: [], wildSeconds: 0, safeSeconds: 0 }),
  "logic/status": () => ({
    activeKey: null,
    isRunning: false,
    currentRuleName: null,
    currentStepNumber: 0,
    currentStepType: null,
    pendingRules: [],
  }),
  "logic/stop": () => ({ stopped: true }),
  "logic/run": () => Promise.resolve({ accepted: true }),
  "logic/getRules": () => ({ found: false, isEngineActive: false, rules: [] }),
  "logic/saveRules": () => ({ saved: true }),
  "logic/getRule": () => ({ found: false, rule: null }),
  "logic/saveRule": () => ({ saved: true }),
  "logic/getTimers": () => ({ timers: [] }),
  "logic/addTimer": () => ({ ok: true, id: "x", reason: null }),
  "logic/removeTimer": () => ({ removed: true }),
});

describe("registerIpcHandlers", () => {
  it("registers exactly one validated handler per declared channel", () => {
    const registrar = fakeRegistrar();
    registerIpcHandlers(registrar, ipcChannels, baseHandlers());
    expect([...registrar.handlers.keys()].sort()).toEqual(Object.keys(ipcChannels).sort());
  });

  it("throws at startup when a declared channel has no handler", () => {
    const registrar = fakeRegistrar();
    const handlers = baseHandlers() as Record<string, unknown>;
    delete handlers["app/logFromRenderer"];
    expect(() =>
      registerIpcHandlers(registrar, ipcChannels, handlers as never),
    ).toThrowError(/declared but has no registered handler/);
  });

  it("wraps a valid call in an ok envelope with schema-validated data", async () => {
    const registrar = fakeRegistrar();
    registerIpcHandlers(registrar, ipcChannels, baseHandlers());
    const result = (await registrar.handlers.get("app/getInfo")!(undefined)) as {
      ok: boolean;
      data?: unknown;
    };
    expect(result.ok).toBe(true);
    expect((result.data as { platform: string }).platform).toBe("win32");
  });

  it("rejects an out-of-contract request with bad_request and never invokes the handler", async () => {
    const registrar = fakeRegistrar();
    const logSpy = vi.fn();
    registerIpcHandlers(registrar, ipcChannels, {
      ...baseHandlers(),
      "app/logFromRenderer": logSpy,
    });
    const result = (await registrar.handlers.get("app/logFromRenderer")!({
      level: "verbose", // not in enum
      scope: "",
      message: "",
    })) as { ok: boolean; error?: { code: string } };
    expect(result.ok).toBe(false);
    expect(result.error?.code).toBe("bad_request");
    expect(logSpy).not.toHaveBeenCalled();
  });

  it("converts a throwing handler into a handler_error envelope", async () => {
    const registrar = fakeRegistrar();
    registerIpcHandlers(registrar, ipcChannels, {
      ...baseHandlers(),
      "app/getInfo": () => {
        throw new Error("boom");
      },
    });
    const result = (await registrar.handlers.get("app/getInfo")!(undefined)) as {
      ok: boolean;
      error?: { code: string; message: string };
    };
    expect(result.ok).toBe(false);
    expect(result.error?.code).toBe("handler_error");
    expect(result.error?.message).toContain("boom");
  });

  it("rejects a handler response that breaches the response schema", async () => {
    const registrar = fakeRegistrar();
    registerIpcHandlers(registrar, ipcChannels, {
      ...baseHandlers(),
      "app/getInfo": () => ({ version: 42 }) as never, // wrong types on purpose (runtime schema must catch it)
    });
    const result = (await registrar.handlers.get("app/getInfo")!(undefined)) as {
      ok: boolean;
      error?: { code: string };
    };
    expect(result.ok).toBe(false);
    expect(result.error?.code).toBe("handler_error");
  });
});
