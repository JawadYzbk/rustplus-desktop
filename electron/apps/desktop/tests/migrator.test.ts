/**
 * Legacy migrator (M3) tests — synthetic %APPDATA%\RustPlusDesk fixture tree, real fs.
 * Proves: transform correctness, seal-at-rest during import, verbatim copies, deferred reporting,
 * and that legacy files are never modified.
 */
import { afterEach, beforeEach, describe, expect, it } from "vitest";
import { existsSync, mkdirSync, mkdtempSync, readFileSync, rmSync, writeFileSync } from "node:fs";
import { tmpdir } from "node:os";
import { join } from "node:path";
import {
  TUTORIAL_STATUS,
  type MigrationRow,
} from "@rpd/shared";
import { SettingsStore } from "../src/main/stores/settings-store.js";
import { ProfilesStore } from "../src/main/stores/profiles-store.js";
import { PassthroughSecretCodec, SEALED_PREFIX } from "../src/main/stores/secret-codec.js";
import {
  AlertTemplateStore,
  DeviceHotkeysStore,
  HotkeyOptionsStore,
  TrackedPlayersStore,
} from "../src/main/stores/legacy-stores.js";
import { TutorialProgressStore } from "../src/main/stores/tutorial-progress-store.js";
import { LegacyMigrator } from "../src/main/services/legacy-migrator.js";

let legacyRoot = "";
let newRoot = "";
let legacySettingsBytes = "";

function writeLegacy(rel: string, content: string): void {
  const p = join(legacyRoot, rel);
  mkdirSync(join(p, ".."), { recursive: true });
  writeFileSync(p, content, "utf8");
}

beforeEach(() => {
  legacyRoot = mkdtempSync(join(tmpdir(), "rpd-legacy-"));
  newRoot = mkdtempSync(join(tmpdir(), "rpd-new-"));

  // --- settings: valid keys + one unknown + one invalid ---
  legacySettingsBytes = JSON.stringify({
    LastHost: "1.2.3.4",
    LastPort: 28082,
    SidebarWidth: 460,
    MapGridOpacity: 0.5,
    TelegramCallMsg: "Custom!",
    NotARealKey: true, // unknown → dropped w/ warning
    MapMonumentOpacity: "bogus", // invalid → reset to default w/ warning
  });
  writeLegacy("tracking_settings.json", legacySettingsBytes);

  // --- profiles: bare array w/ plaintext token (legacy format) ---
  writeLegacy(
    "profiles.json",
    JSON.stringify([
      {
        Name: "Main",
        Description: "",
        Host: "5.6.7.8",
        Port: 28082,
        SteamId64: "76561198000000000",
        PlayerToken: "plain-token",
        Devices: [{ Kind: "SmartSwitch", EntityId: 42 }],
      },
    ]),
  );

  writeLegacy("hotkeys.json", JSON.stringify({ "5.6.7.8:28082|76561198000000000": { "Ctrl+F1": [42] } }));
  writeLegacy("hotkey_options.json", JSON.stringify({ ParallelMode: true, ToggleDelayMs: 200 }));
  writeLegacy("custom_alerts.json", JSON.stringify({ "de-DE": { CargoSpawned: "Cargo!" } }));
  writeLegacy(
    "tracked_players.json",
    JSON.stringify([
      {
        BMId: "bm-1",
        Name: "Neelo",
        LastServerName: "US Main",
        GroupName: "",
        GroupColor: "",
        Sessions: [],
        IsBMOnly: false,
      },
    ]),
  );
  writeLegacy(
    "tutorial-progress.json",
    JSON.stringify({
      Tutorials: {
        basics: { TutorialId: "basics", TutorialVersion: 2, Status: TUTORIAL_STATUS.Completed, CompletedStepIds: ["s1"] },
      },
      Preferences: { FirstRunPromptDismissed: true },
    }),
  );
  writeLegacy("rustplusjs-config.json", JSON.stringify({ steam_id: "76561198000000000" }));
  mkdirSync(join(legacyRoot, "cache"), { recursive: true });
  writeLegacy(join("cache", "minimap_settings.json"), JSON.stringify({ zoom: 3 }));
});

afterEach(() => {
  rmSync(legacyRoot, { recursive: true, force: true });
  rmSync(newRoot, { recursive: true, force: true });
});

/** Deterministic sealing codec so tests can assert ciphertext-at-rest without Electron. */
class SealingCodec extends PassthroughSecretCodec {
  override seal(plaintext: string): string {
    return plaintext ? SEALED_PREFIX + Buffer.from(plaintext, "utf8").toString("base64") : plaintext;
  }
  override open(blob: string): string {
    if (!blob.startsWith(SEALED_PREFIX)) return blob;
    return Buffer.from(blob.slice(SEALED_PREFIX.length), "base64").toString("utf8");
  }
}

function buildMigrator(): LegacyMigrator {
  return new LegacyMigrator(
    { appData: legacyRoot, localData: join(newRoot, "_no_local"), deathsDir: join(newRoot, "_no_deaths") },
    newRoot,
    {
      settings: new SettingsStore(newRoot),
      profiles: new ProfilesStore(newRoot, new SealingCodec()),
      hotkeys: new DeviceHotkeysStore(newRoot),
      hotkeyOptions: new HotkeyOptionsStore(newRoot),
      alerts: new AlertTemplateStore(newRoot),
      trackedPlayers: new TrackedPlayersStore(newRoot),
      tutorials: new TutorialProgressStore(newRoot),
    },
  );
}

function row(rows: MigrationRow[], label: string): MigrationRow {
  const found = rows.find((r) => r.source === label);
  if (!found) throw new Error(`no row for ${label}; got ${rows.map((r) => r.source).join(", ")}`);
  return found;
}

describe("LegacyMigrator", () => {
  it("scan reports roots and per-source existence without touching anything", () => {
    const m = buildMigrator();
    const scan = m.scan();
    expect(scan.roots.find((r) => r.kind === "appData")?.exists).toBe(true);
    const profiles = scan.sources.find((s) => s.id === "profiles");
    expect(profiles?.exists).toBe(true);
    expect(profiles?.bytes).toBeGreaterThan(0);
    // Read-only: legacy file bytes unchanged after scan.
    expect(readFileSync(join(legacyRoot, "tracking_settings.json"), "utf8")).toBe(legacySettingsBytes);
  });

  it("run imports everything; settings merge keeps valid values, drops unknown, resets invalid", () => {
    const report = buildMigrator().run();

    expect(row(report.rows, "Tracking settings").status).toBe("warning");
    const settings = new SettingsStore(newRoot).all;
    expect(settings.LastHost).toBe("1.2.3.4");
    expect(settings.SidebarWidth).toBe(460);
    expect(settings.MapGridOpacity).toBe(0.5); // valid custom value kept
    expect(settings.TelegramCallMsg).toBe("Custom!");
    expect(settings.MapMonumentOpacity).toBe(1.0); // invalid → default
    const warnings = row(report.rows, "Tracking settings").warnings ?? [];
    expect(warnings.some((w) => w.includes("NotARealKey"))).toBe(true);
    expect(warnings.some((w) => w.startsWith("MapMonumentOpacity"))).toBe(true);

    // Profiles imported with token sealed at rest, unknown fields preserved.
    expect(row(report.rows, "Server profiles").status).toBe("migrated");
    const rawProfiles = readFileSync(join(newRoot, "profiles.json"), "utf8");
    expect(rawProfiles).not.toContain("plain-token");
    expect(rawProfiles).toContain(SEALED_PREFIX);
    const codec = new SealingCodec();
    const reloaded = new ProfilesStore(newRoot, codec);
    expect(reloaded.tokenFor("5.6.7.8:28082|76561198000000000")).toBe("plain-token");
    const storedDoc = JSON.parse(rawProfiles) as { profiles: Array<Record<string, unknown>> };
    expect(storedDoc.profiles[0]!["Devices"]).toEqual([{ Kind: "SmartSwitch", EntityId: 42 }]);

    expect(row(report.rows, "Device hotkeys").status).toBe("migrated");
    expect(new DeviceHotkeysStore(newRoot).all()["5.6.7.8:28082|76561198000000000"]).toEqual({ "Ctrl+F1": [42] });

    expect(row(report.rows, "Hotkey options").status).toBe("migrated");
    expect(new HotkeyOptionsStore(newRoot).get()).toEqual({ ParallelMode: true, ToggleDelayMs: 200 });

    expect(row(report.rows, "Custom alert templates").status).toBe("migrated");
    expect(new AlertTemplateStore(newRoot).all()["de-DE"]).toEqual({ CargoSpawned: "Cargo!" });

    expect(row(report.rows, "Tracked players").status).toBe("migrated");
    expect(new TrackedPlayersStore(newRoot).list()[0]!.Name).toBe("Neelo");

    expect(row(report.rows, "Tutorial progress").status).toBe("migrated");
    const tut = new TutorialProgressStore(newRoot);
    expect(tut.all()["basics"]!.Status).toBe(TUTORIAL_STATUS.Completed);
    expect(tut.preferences().FirstRunPromptDismissed).toBe(true);

    expect(row(report.rows, "FCM pairing config").status).toBe("copied");
    expect(existsSync(join(newRoot, "rustplusjs-config.json"))).toBe(true);

    expect(row(report.rows, "Cache files").status).toBe("copied");
    expect(JSON.parse(readFileSync(join(newRoot, "cache", "minimap_settings.json"), "utf8"))).toEqual({ zoom: 3 });

    // Deferred stage-owned data is always reported for the full picture.
    expect(row(report.rows, "Map overlays (Overlays\\)").status).toBe("deferred");
    expect(row(report.rows, "Death logs (RustPlusDesktop\\deaths)").status).toBe("deferred");
  });

  it("missing sources report 'missing' and never fail the run", () => {
    rmSync(join(legacyRoot, "hotkeys.json"), { force: true });
    rmSync(join(legacyRoot, "tracked_players.json"), { force: true });
    const report = buildMigrator().run();
    expect(row(report.rows, "Device hotkeys").status).toBe("missing");
    expect(row(report.rows, "Tracked players").status).toBe("missing");
    expect(report.rows.every((r) => r.status !== "failed")).toBe(true);
  });

  it("is idempotent — running twice yields the same end state", () => {
    const m = buildMigrator();
    m.run();
    const first = readFileSync(join(newRoot, "profiles.json"), "utf8");
    m.run();
    const second = readFileSync(join(newRoot, "profiles.json"), "utf8");
    expect(second).toBe(first);
    expect(new ProfilesStore(newRoot, new SealingCodec()).tokenFor("5.6.7.8:28082|76561198000000000")).toBe("plain-token");
  });
});
