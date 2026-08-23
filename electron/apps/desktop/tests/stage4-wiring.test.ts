/**
 * Stage-4 wiring golden tests — CLI runtime resolution (node/browser/cli-entry), fcm-register
 * browser retry parity, and ConnRuntime composition (broadcasts→hub, lifecycle→polls, event push).
 */
import { describe, expect, it } from "vitest";
import { EventEmitter } from "node:events";
import {
  findBundledNode,
  resolveCliEntry,
  findChromiumBrowser,
  browserRegisterEnvs,
} from "../src/main/services/rustplus/cli-runtime.js";
import { PairingListenerService } from "../src/main/services/rustplus/pairing-listener.js";
import {
  ConnRuntime,
  type ConnRuntimeDeps,
} from "../src/main/services/rustplus/conn-runtime.js";
import { DeviceEventHub } from "../src/main/services/rustplus/device-hub.js";
import { PollService } from "../src/main/services/rustplus/poll-service.js";
import { ConnectionManager } from "../src/main/services/rustplus/connection-manager.js";

// ---------------------------------------------------------------- findBundledNode

describe("findBundledNode", () => {
  const existsFor = (paths: string[]) => (p: string) => paths.includes(p.replace(/\\/g, "/"));

  it("prefers runtime/node-win-x64/node.exe, then flat variants, per base dir", () => {
    const found = findBundledNode(["C:\\app", "C:\\exe"], {
      exists: existsFor(["C:/exe/node-win-x64/node.exe"]),
    });
    expect(found?.replace(/\\/g, "/")).toBe("C:/exe/node-win-x64/node.exe");
  });

  it("walks up at most 5 levels as dev fallback", () => {
    const base = "C:/dev/proj/a/b/c";
    const deep = "C:/dev/runtime/node-win-x64/node.exe";
    const found = findBundledNode([base], { exists: existsFor([deep]) });
    // C:/dev/proj/a/b/c → up to 5 parents reaches C:/dev
    expect(found?.replace(/\\/g, "/")).toBe(deep);
  });

  it("returns null when nothing exists", () => {
    expect(findBundledNode(["C:\\nowhere"], { exists: () => false })).toBeNull();
  });
});

describe("resolveCliEntry", () => {
  it("returns the first existing candidate", () => {
    const root = "C:/cli";
    const found = resolveCliEntry(root, (p) => p.replace(/\\/g, "/") === `${root}/index.js`);
    expect(found?.replace(/\\/g, "/")).toBe(`${root}/index.js`);
  });
  it("returns null when no candidate exists", () => {
    expect(resolveCliEntry("C:/cli", () => false)).toBeNull();
  });
});

describe("findChromiumBrowser", () => {
  it("registry hit wins over filesystem probes and names the candidate", () => {
    const found = findChromiumBrowser(["C:/PF"], {
      exists: () => false,
      registryLookup: (exe) => (exe === "chrome.exe" ? "C:/reg/chrome.exe" : null),
    });
    expect(found).toEqual({ path: "C:/reg/chrome.exe", name: "Google Chrome" });
  });

  it("falls back to ProgramFiles-relative probes in preference order", () => {
    const found = findChromiumBrowser(["C:/PF"], {
      exists: (p) => p.replace(/\\/g, "/") === "C:/PF/Microsoft/Edge/Application/msedge.exe",
    });
    expect(found).toEqual({ path: "C:\\PF\\Microsoft\\Edge\\Application\\msedge.exe", name: "Microsoft Edge" });
  });

  it("onlyThese restricts candidates (second-attempt fallback parity)", () => {
    const found = findChromiumBrowser([], {
      exists: () => false,
      onlyThese: ["msedge.exe", "chrome.exe"],
      registryLookup: (exe) => (exe === "msedge.exe" ? "C:/edge.exe" : null),
    });
    expect(found).toEqual({ path: "C:/edge.exe", name: "Microsoft Edge" });
  });
});

// ------------------------------------------------- PairingListenerService browser retry

function fakeSpawner(codes: number[], envsSeen: Record<string, string>[]) {
  return {
    run: async (_args: string, env: Record<string, string>) => {
      envsSeen.push(env);
      return codes.shift() ?? 1;
    },
    start: (() => null) as never,
  };
}

describe("PairingListenerService — fcm-register browser retry parity", () => {
  function makeService(codes: number[], envsSeen: Record<string, string>[] = [], resolver?: () => ReturnType<typeof browserRegisterEnvs>[]) {
    return new PairingListenerService({
      nodeExe: "node.exe",
      cliEntry: "cli.js",
      configPath: "Z:/definitely-missing/fcm.json",
      spawner: fakeSpawner(codes, envsSeen) as never,
      browserResolver: resolver,
    });
  }

  it("retries with the fallback browser when the primary refuses to start", async () => {
    const envsSeen: Record<string, string>[] = [];
    const svc = makeService(
      [1, 0],
      envsSeen,
      () =>
        [
          { label: "Google Chrome", env: { PUPPETEER_EXECUTABLE_PATH: "C:/chrome.exe", CHROME_PATH: "C:/chrome.exe" } },
          { label: "Microsoft Edge", env: { PUPPETEER_EXECUTABLE_PATH: "C:/edge.exe", CHROME_PATH: "C:/edge.exe" } },
        ] as never,
    );
    await expect(svc.start()).resolves.toBeUndefined();
    expect(envsSeen.map((e) => e.PUPPETEER_EXECUTABLE_PATH)).toEqual(["C:/chrome.exe", "C:/edge.exe"]);
    expect(svc.isRunning).toBe(true);
  });

  it("aborts entirely when every browser attempt fails (legacy parity)", async () => {
    const envsSeen: Record<string, string>[] = [];
    const statuses: string[] = [];
    const svc = makeService([1, 1], envsSeen, () => [
      { label: "Google Chrome", env: {} },
    ] as never);
    svc.on("status", (s) => statuses.push(s));
    await expect(svc.start()).rejects.toThrowError(/fcm-register failed/);
    expect(statuses).toContain("stopped");
    expect(svc.isRunning).toBe(false);
  });

  it("reports the no-browser case without spawning anything", async () => {
    const envsSeen: Record<string, string>[] = [];
    const logs: string[] = [];
    const svc = new PairingListenerService({
      nodeExe: "node.exe",
      cliEntry: "cli.js",
      configPath: "Z:/definitely-missing/fcm.json",
      spawner: fakeSpawner([0], envsSeen) as never,
      log: (m) => logs.push(m),
      browserResolver: () => [],
    });
    await expect(svc.start()).rejects.toThrowError(/No Chromium-based browser found/);
    expect(envsSeen).toHaveLength(0);
    expect(logs.join("\n")).toContain("❌");
  });
});

// ------------------------------------------------------------------ ConnRuntime

describe("ConnRuntime wiring", () => {
  function rig() {
    const transport = { events: new EventEmitter() };
    // Bare-emitter stand-in for the manager: ConnRuntime only uses .on/.removeAllListeners.
    const managerEvents = new EventEmitter();
    const managerLike = managerEvents as unknown as ConnectionManager;

    const polls = new PollService(
      { isConnected: false, send: async () => ({}), recordStatusResult: () => undefined },
      undefined,
      { statusMs: 3_600_000, teamMs: 3_600_000, markersMs: 3_600_000 },
    );
    const hubSends: number[] = [];
    const hub = new DeviceEventHub({ send: async (d) => { hubSends.push(d["entityId"] as number); throw new Error("offline"); } });

    const deps: ConnRuntimeDeps = {
      transport,
      manager: managerLike,
      polls,
      hub,
    };
    const rt = new ConnRuntime(deps);
    return { transport, managerEvents, managerLike, polls, hub, hubSends, rt };
  }

  it("routes entityChanged broadcasts into the hub", async () => {
    const r = rig();
    const handled: Array<{ entityId: number; value?: unknown }> = [];
    (r.hub as unknown as { handleEntityChanged: unknown }).handleEntityChanged = async (
      entityId: number,
      payload: { value?: unknown },
    ) => void handled.push({ entityId, value: payload.value });

    r.rt.wire();
    r.transport.events.emit("message", { broadcast: { entityChanged: { entityId: 9, payload: { value: true } } } });
    await Promise.resolve();
    await Promise.resolve();

    expect(handled).toEqual([{ entityId: 9, value: true }]);
    r.polls.stop();
  });

  it("starts/stops polling and resets the hub across the connection lifecycle", () => {
    const r = rig();
    const resets: number[] = [];
    (r.hub as unknown as { reset: unknown }).reset = () => void resets.push(1);

    r.rt.wire();
    r.managerEvents.emit("connected", { proxy: "direct" });
    expect(r.polls.isRunning).toBe(true);
    expect(resets).toHaveLength(1);

    r.managerEvents.emit("lost", { reason: "x" });
    expect(r.polls.isRunning).toBe(false);

    r.managerEvents.emit("disconnected");
    expect(resets).toHaveLength(2);
  });

  it("forwards conn/poll/device streams as unified push events", () => {
    const r = rig();
    const pushes: Array<{ stream: string; event: { kind?: string } }> = [];
    r.rt.on("push", (p) => pushes.push(p as { stream: string; event: { kind?: string } }));

    r.rt.wire();
    r.managerEvents.emit("connected", { proxy: "proxy" });
    r.managerEvents.emit("reconnectingIn", { delayMs: 4000 });
    r.hub.emit("event", { kind: "deviceState", entityId: 1, on: true, deviceType: "SmartSwitch" });

    expect(pushes.some((p) => p.stream === "conn" && p.event.kind === "connected")).toBe(true);
    expect(pushes.some((p) => p.stream === "conn" && p.event.kind === "reconnectingIn")).toBe(true);
    expect(pushes.some((p) => p.stream === "device")).toBe(true);
  });
});
