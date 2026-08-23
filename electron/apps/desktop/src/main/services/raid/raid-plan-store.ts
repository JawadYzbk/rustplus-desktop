/**
 * Raid plan store — port of Services/Raid/RaidPlanStore.cs: JSON list persisted atomically
 * (tmp file + rename), missing/corrupt files load as empty.
 * Deviation note: the C# default path is %LocalAppData%/RustPlusDesk/raid-plan.json; here the
 * caller passes the userData path explicitly (index.ts wiring) so this module stays
 * importable outside Electron (vitest).
 */
import { existsSync, mkdirSync, readFileSync, renameSync, writeFileSync } from "node:fs";
import { dirname, join } from "node:path";
import type { RaidPlanEntry } from "./raid-models.js";

export class RaidPlanStore {
  private readonly path: string;

  constructor(path: string) {
    this.path = path;
  }

  /** Default location parity helper for production wiring (%LocalAppData% equivalent). */
  static defaultPath(userDataDir: string): string {
    return join(userDataDir, "raid-plan.json");
  }

  /** LoadAsync parity — corrupt/missing returns []. */
  load(): RaidPlanEntry[] {
    if (!existsSync(this.path)) return [];
    try {
      const parsed = JSON.parse(readFileSync(this.path, "utf-8")) as unknown;
      if (!Array.isArray(parsed)) return [];
      return parsed.filter(
        (e): e is RaidPlanEntry =>
          typeof e === "object" &&
          e !== null &&
          typeof (e as RaidPlanEntry).targetId === "number" &&
          typeof (e as RaidPlanEntry).quantity === "number" &&
          typeof (e as RaidPlanEntry).sourceId === "number",
      );
    } catch {
      return [];
    }
  }

  /** SaveAsync parity — write tmp then rename over the target. */
  save(entries: readonly RaidPlanEntry[]): void {
    const directory = dirname(this.path);
    mkdirSync(directory, { recursive: true });
    const temporaryPath = this.path + ".tmp";
    writeFileSync(temporaryPath, `${JSON.stringify(entries, null, 2)}\n`, "utf-8");
    renameSync(temporaryPath, this.path);
  }
}
