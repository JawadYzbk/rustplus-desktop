/**
 * Granular reset — port of PerformGranularResetAsync (Connection.Reset.cs:123-216, audit §4).
 *
 * Parity notes:
 *  - "cache" deletes the whole cache dir AND Overlays\ (the legacy side effect is preserved, warning surfaced).
 *  - "steam" clears SteamId64 in tracking settings.
 *  - "pairing" deletes rustplusjs-config.json (FCM creds; listener restart is runtime state, stage 4).
 *  - "profiles" clears the profile list.
 *  - "three_d" wipes 3DMaps\.
 *  - Connection hard-reset and cloud orphan purge are runtime/cloud-stage concerns (4 / 10) and are NOT here.
 */
import { existsSync, rmSync } from "node:fs";
import { join } from "node:path";
import type { ResetTarget } from "@rpd/shared";
import { SettingsStore } from "../stores/settings-store.js";
import { ProfilesStore } from "../stores/profiles-store.js";

export interface ResetContext {
  userDataDir: string;
  settings: SettingsStore;
  profiles: ProfilesStore;
  log?: (level: "warn", message: string) => void;
}

const OVERLAY_WARNING =
  "cache reset also deleted Overlays\\ (legacy parity side effect)";

export function performReset(ctx: ResetContext, targets: ResetTarget[]): { performed: string[] } {
  const performed: string[] = [];

  for (const target of targets) {
    switch (target) {
      case "profiles": {
        // Clear by removing every stored match key.
        for (const p of ctx.profiles.list()) {
          ctx.profiles.removeByMatchKey(`${p.Host}:${p.Port}|${p.SteamId64}`);
        }
        performed.push("profiles");
        break;
      }
      case "steam": {
        ctx.settings.patch({ SteamId64: "" });
        performed.push("steam");
        break;
      }
      case "pairing": {
        rmIfExists(join(ctx.userDataDir, "rustplusjs-config.json"));
        performed.push("pairing");
        break;
      }
      case "crosshairs": {
        rmIfExists(join(ctx.userDataDir, "custom_crosshairs.json"));
        performed.push("crosshairs");
        break;
      }
      case "cache": {
        rmIfExists(join(ctx.userDataDir, "cache"), true);
        rmIfExists(join(ctx.userDataDir, "Overlays"), true); // legacy side effect
        ctx.log?.("warn", OVERLAY_WARNING);
        performed.push("cache");
        break;
      }
      case "three_d": {
        rmIfExists(join(ctx.userDataDir, "3DMaps"), true);
        performed.push("three_d");
        break;
      }
    }
  }

  return { performed };
}

function rmIfExists(path: string, dir = false): void {
  if (!existsSync(path)) return;
  rmSync(path, { recursive: dir, force: true });
}
