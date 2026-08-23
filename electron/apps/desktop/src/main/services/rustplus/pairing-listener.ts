/**
 * Pairing listener service — port of PairingListenerRealProcess lifecycle (Option C hybrid, audit §9):
 * keeps rustplus-cli subprocesses for FCM register/listen while the persistent Rust+ socket runs
 * in-process. Semantics preserved:
 *  - register only when config missing or < 50 bytes; record FcmIssuedAt=now, FcmExpiresAt=+15 days;
 *  - long-running `fcm-listen`; stdout lines go through PairingLineParser;
 *  - unexpected exit → auto-restart after 3 s unless stopped; registration failure aborts (parity:
 *    legacy stops entirely rather than looping a broken browser automation).
 */
import { spawn, type ChildProcess } from "node:child_process";
import { existsSync, statSync } from "node:fs";
import { EventEmitter } from "node:events";
import { PairingLineParser, type ListenerEvent } from "./pairing-parser.js";

export const FCM_CONFIG_MIN_BYTES = 50;
export const RESTART_DELAY_MS = 3_000;
export const FCM_TOKEN_DAYS = 15;

export interface CliSpawner {
  /** Starts the long-running CLI; returns stdout lines via the handler. */
  start(args: string, onLine: (line: string) => void, onExit: (code: number | null) => void): ChildProcess;
  /** One-shot run for `fcm-register`; resolves with exit code. */
  run(args: string, env: Record<string, string>): Promise<number>;
}

export interface PairingListenerDeps {
  nodeExe: string;
  cliEntry: string;
  configPath: string;
  spawner: CliSpawner;
  log?: (message: string) => void;
}

export class PairingListenerService extends EventEmitter {
  private proc: ChildProcess | null = null;
  private running = false;
  private stopping = false;
  private readonly parser = new PairingLineParser();

  constructor(private readonly deps: PairingListenerDeps) {
    super();
  }

  get isRunning(): boolean {
    return this.running;
  }

  /**
   * Starts register-if-needed + listen. Emits: "status" (starting|registering|listening|stopped),
   * "event" (ListenerEvent), "registrationCompleted", "restartScheduled".
   * Resolves once the listener process is up (or rejects on registration failure).
   */
  async start(): Promise<void> {
    if (this.running && this.proc && this.proc.exitCode === null) return;
    this.stopping = false;
    this.emit("status", "starting");

    const needsRegister =
      !existsSync(this.deps.configPath) || statSync(this.deps.configPath).size < FCM_CONFIG_MIN_BYTES;

    if (needsRegister) {
      this.emit("status", "registering");
      this.deps.log?.("Starting one time registration (fcm-register) …");
      const code = await this.deps.spawner.run(
        "fcm-register",
        {}, // browser discovery env is resolved by the CLI itself in the Electron era
      );
      if (code !== 0) {
        // Legacy parity: stop entirely — do not loop broken browser automation.
        this.running = false;
        this.emit("status", "stopped");
        throw new Error(`fcm-register failed with exit code ${code}`);
      }
      this.deps.log?.("Registration completed.");
      this.emit("registrationCompleted", {
        issuedAt: Date.now(),
        expiresAt: Date.now() + FCM_TOKEN_DAYS * 24 * 60 * 60 * 1000,
      });
    }

    this.deps.log?.("Starting Listener (fcm-listen) …");
    this.proc = this.deps.spawner.start(
      "fcm-listen",
      (line) => {
        for (const evt of this.parser.feed(line)) {
          if (evt.kind === "listening") this.deps.log?.("Listening for FCM Notifications");
          this.emit("event", evt);
        }
      },
      (code) => {
        this.running = false;
        this.emit("status", "stopped");
        if (this.stopping) return;
        this.deps.log?.(`fcm-listen exited (${code}) – restarting in ${RESTART_DELAY_MS / 1000}s…`);
        this.emit("restartScheduled", RESTART_DELAY_MS);
        setTimeout(() => {
          if (!this.stopping) void this.start().catch(() => undefined);
        }, RESTART_DELAY_MS);
      },
    );
    this.running = true;
    this.emit("status", "listening");
  }

  async stop(): Promise<void> {
    this.stopping = true;
    try {
      this.proc?.kill();
    } catch {
      /* already dead */
    }
    this.proc = null;
    this.running = false;
    this.emit("status", "stopped");
  }
}
