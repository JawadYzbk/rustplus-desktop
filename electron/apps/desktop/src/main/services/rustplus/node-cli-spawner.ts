/**
 * NodeCliSpawner — real CliSpawner over node:child_process, replacing the legacy
 * ProcessStartInfo plumbing. Args are passed as an argv array (no shell), so quoting quirks from
 * the C# string-concatenation era cannot resurface.
 */
import { spawn, type ChildProcess } from "node:child_process";
import type { CliSpawner } from "./pairing-listener.js";

export interface NodeCliSpawnerDeps {
  nodeExe: string;
  cliEntry: string;
  workingDir: string;
}

export class NodeCliSpawner implements CliSpawner {
  constructor(private readonly deps: NodeCliSpawnerDeps) {}

  start(
    subcommand: string,
    onLine: (line: string) => void,
    onExit: (code: number | null) => void,
    extraEnv: Record<string, string> = {},
  ): ChildProcess {
    const child = spawn(this.deps.nodeExe, [this.deps.cliEntry, subcommand], {
      cwd: this.deps.workingDir,
      env: { ...process.env, ...extraEnv },
      stdio: ["ignore", "pipe", "pipe"],
      windowsHide: true,
    });
    child.stdout?.setEncoding("utf8");
    child.stderr?.setEncoding("utf8");
    child.stdout?.on("data", (chunk: string) => {
      for (const line of chunk.split(/\r?\n/)) if (line.length > 0) onLine(line);
    });
    child.stderr?.on("data", (chunk: string) => {
      for (const line of chunk.split(/\r?\n/)) if (line.length > 0) onLine(line);
    });
    child.on("close", (code) => onExit(code));
    return child;
  }

  /** One-shot fcm-register with browser env; resolves with the exit code (legacy parity). */
  run(subcommand: string, env: Record<string, string> = {}): Promise<number> {
    return new Promise((resolve, reject) => {
      const child = this.start(subcommand, () => undefined, (code) => {
        if (code === null) reject(new Error(`fcm-register process died without exit code`));
        else resolve(code);
      }, env);
      child.on("error", (err) => reject(err));
    });
  }
}
