/**
 * ConnRuntime — composition wiring for one connection lifecycle:
 *  - transport broadcast messages → DeviceEventHub (entityChanged extraction);
 *  - ConnectionManager connected/lost/disconnected → hub reset + PollService start/stop;
 *  - every manager/poll/hub event → forwarded to a listener (renderer push in bootstrap).
 * Pure EventEmitter plumbing — no Electron imports, fully scriptable in tests.
 */
import { EventEmitter } from "node:events";
import { extractEntityChanged } from "./device-hub.js";
import type { DeviceEventHub } from "./device-hub.js";
import type { PollService } from "./poll-service.js";
import type { ConnectionManager } from "./connection-manager.js";

export interface TransportEventSurface {
  readonly events: EventEmitter; // emits "message" with raw AppMessage broadcasts
}

export interface ConnRuntimeDeps {
  transport: TransportEventSurface;
  manager: ConnectionManager;
  polls: PollService;
  hub: DeviceEventHub;
}

export class ConnRuntime extends EventEmitter {
  private wired = false;

  constructor(private readonly deps: ConnRuntimeDeps) {
    super();
  }

  /** Idempotent: attaches all listeners once. */
  wire(): void {
    if (this.wired) return;
    this.wired = true;
    const { transport, manager, polls, hub } = this.deps;

    transport.events.on("message", (msg: unknown) => {
      const ec = extractEntityChanged(msg);
      if (!ec) return;
      void hub.handleEntityChanged(ec.entityId, ec.payload).catch(() => undefined);
    });

    manager.on("connected", () => {
      hub.reset(); // per-connection state dies with the connection
      polls.start();
    });
    manager.on("lost", () => {
      // Keep polling during silent refresh? No — legacy status loop keeps running, but the
      // watchdog already drives reconnection; stop poll sends while the socket is down.
      polls.stop();
    });
    manager.on("disconnected", () => {
      polls.stop();
      hub.reset();
    });

    polls.on("poll", (e: unknown) => this.emit("push", { stream: "poll", event: e }));
    hub.on("event", (e: unknown) => this.emit("push", { stream: "device", event: e }));
    for (const name of ["connecting", "connected", "lost", "disconnected", "reconnectingIn"] as const) {
      manager.on(name, (e: unknown) => this.emit("push", { stream: "conn", event: { kind: name, ...((e as object) ?? {}) } }));
    }
  }

  unwire(): void {
    if (!this.wired) return;
    const { transport, manager, polls, hub } = this.deps;
    transport.events.removeAllListeners("message");
    for (const name of ["connected", "lost", "disconnected"] as const) {
      manager.removeAllListeners(name);
    }
    for (const name of ["connecting", "reconnectingIn"]) {
      manager.removeAllListeners(name);
    }
    polls.removeAllListeners("poll");
    hub.removeAllListeners("event");
    this.wired = false;
  }
}
