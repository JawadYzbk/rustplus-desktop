/**
 * Contract tests for the typed IPC registry — no Electron runtime required (fake registrar).
 * Exit criterion for stage 2: "IPC contract tests green".
 */
import { describe, expect, it, vi } from "vitest";
import { ipcChannels } from "@rpd/shared";
import { registerIpcHandlers, type IpcRegistrar } from "../src/main/ipc.js";

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
  "uiPrefs/get": () => ({ sidebarPinned: true, sidebarWidth: 420 }),
  "uiPrefs/set": (patch: { sidebarPinned?: boolean; sidebarWidth?: number }) => ({
    sidebarPinned: patch.sidebarPinned ?? true,
    sidebarWidth: patch.sidebarWidth ?? 420,
  }),
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
