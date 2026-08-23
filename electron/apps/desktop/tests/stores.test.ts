/**
 * Settings + profiles store tests — real filesystem (mkdtemp), passthrough secret codec.
 * Defaults are pinned against the audit-documented legacy values (DATA_STORES §2).
 */
import { afterEach, beforeEach, describe, expect, it } from "vitest";
import { mkdtempSync, existsSync, readFileSync, rmSync } from "node:fs";
import { tmpdir } from "node:os";
import { join } from "node:path";
import { TRACKING_SETTINGS_DEFAULTS } from "@rpd/shared";
import { SettingsStore } from "../src/main/stores/settings-store.js";
import { ProfilesStore } from "../src/main/stores/profiles-store.js";
import { PassthroughSecretCodec, SEALED_PREFIX } from "../src/main/stores/secret-codec.js";

let dir = "";

beforeEach(() => {
  dir = mkdtempSync(join(tmpdir(), "rpd-stores-"));
});

afterEach(() => {
  rmSync(dir, { recursive: true, force: true });
});

describe("SettingsStore", () => {
  it("returns exact legacy defaults on first run", () => {
    const store = new SettingsStore(dir);
    const s = store.all;
    // Spot-pins across every documented group (full catalog lives in the shared schema).
    expect(s.SidebarWidth).toBe(420);
    expect(s.SidebarPinned).toBe(true);
    expect(s.WindowWidth).toBe(1280);
    expect(s.WindowHeight).toBe(720);
    expect(s.WindowLeft).toBe("NaN");
    expect(s.AutoConnectEnabled).toBe(false);
    expect(s.BackgroundTrackingEnabled).toBe(true);
    expect(s.MapGridOpacity).toBe(0.7);
    expect(s.MapShowSteamMarkers).toBe(true);
    expect(s.AnnounceCargo).toBe(false);
    expect(s.ListenForServerEvents).toBe(true);
    expect(s.AfkAlertMinutes).toBe(5);
    expect(s.OfflineDeathAlertsEnabled).toBe(true);
    expect(s.AutoLoadShops).toBe(true);
    expect(s.GenericAlarmOverlayEnabled).toBe(true);
    expect(s.TelegramCallMsg).toBe("Alarm ausgeloest!");
    expect(s.TelegramCallLang).toBe("de-DE-Standard-A");
    expect(s.NotificationsRetentionDays).toBe(30);
    expect(s.LastCrosshairStyle).toBe("GreenDot");
    // Full-object sanity: defaults must satisfy the schema verbatim.
    expect(s).toEqual(TRACKING_SETTINGS_DEFAULTS);
  });

  it("persists a patch atomically and reloads it in a new store instance", () => {
    new SettingsStore(dir).patch({ MapGridOpacity: 0.5, SidebarWidth: 460 });
    const reloaded = new SettingsStore(dir).all;
    expect(reloaded.MapGridOpacity).toBe(0.5);
    expect(reloaded.SidebarWidth).toBe(460);
    // Untouched keys stay at defaults.
    expect(reloaded.AnnounceCargo).toBe(false);
    expect(existsSync(join(dir, "tracking_settings.json"))).toBe(true);
  });

  it("rejects unknown keys loudly instead of forking state", () => {
    const store = new SettingsStore(dir);
    expect(() => store.patch({ NotARealKey: true } as never)).toThrowError(/breached contract/);
  });

  it("writes schemaVersion into the file", () => {
    new SettingsStore(dir).patch({});
    const doc = JSON.parse(readFileSync(join(dir, "tracking_settings.json"), "utf8")) as { schemaVersion: number };
    expect(doc.schemaVersion).toBe(1);
  });
});

describe("ProfilesStore", () => {
  const baseProfile = () => ({
    Name: "Main",
    Host: "1.2.3.4",
    Port: 28082,
    SteamId64: "76561198000000000",
    PlayerToken: "secret-token-abc",
    // Unknown-to-us legacy fields that MUST survive round-trips untouched:
    Devices: [{ Kind: "SmartSwitch", EntityId: 12345, Children: null, Name: "Door" }],
    LogicRules: [{ Id: "r1", SomethingNested: { a: 1 } }],
    deathMarkers: [{ X: 3, Y: 4 }],
    LearnedDaySpeed: 0.24,
    LearnedNightSpeed: 1.2,
    CameraIds: ["CAM-1"],
    IsConnected: false,
    IsFullConnected: false,
  });

  it("upsert + list round-trips core fields and preserves unknown legacy fields byte-faithfully", () => {
    const store = new ProfilesStore(dir, new PassthroughSecretCodec());
    const p = baseProfile();
    store.upsert(p);

    const reloaded = new ProfilesStore(dir, new PassthroughSecretCodec());
    const listed = reloaded.list();
    expect(listed).toHaveLength(1);
    const first = listed[0]!;
    expect(first.Name).toBe("Main");
    expect(first.Port).toBe(28082);

    // Lossless check: every unknown key survives with identical value.
    const raw = JSON.parse(readFileSync(join(dir, "profiles.json"), "utf8")) as { profiles: Record<string, unknown>[] };
    const stored = raw.profiles[0]!;
    expect(stored["Devices"]).toEqual(p.Devices);
    expect(stored["LogicRules"]).toEqual(p.LogicRules);
    expect(stored["deathMarkers"]).toEqual(p.deathMarkers);
    expect(stored["LearnedDaySpeed"]).toBe(0.24);
    expect(stored["CameraIds"]).toEqual(p.CameraIds);
  });

  it("seals PlayerToken at rest (file never contains plaintext) and unseals on read", () => {
    const codec = new PrefixCodec();
    const store = new ProfilesStore(dir, codec);
    store.upsert(baseProfile());

    const raw = readFileSync(join(dir, "profiles.json"), "utf8");
    expect(raw).not.toContain("secret-token-abc");
    expect(raw).toContain(SEALED_PREFIX);

    expect(new ProfilesStore(dir, codec).tokenFor("1.2.3.4:28082|76561198000000000")).toBe("secret-token-abc");
  });

  it("matchKey parity: host:port|steamId, token excluded", () => {
    const store = new ProfilesStore(dir, new PassthroughSecretCodec());
    store.upsert(baseProfile());
    // Same server re-paired with a different token → same matchKey → update, not duplicate.
    store.upsert({ ...baseProfile(), PlayerToken: "rotated-token" });
    expect(store.list()).toHaveLength(1);
    expect(store.tokenFor("1.2.3.4:28082|76561198000000000")).toBe("rotated-token");
  });

  it("removeByMatchKey deletes only the matching profile", () => {
    const store = new ProfilesStore(dir, new PassthroughSecretCodec());
    store.upsert(baseProfile());
    store.upsert({ ...baseProfile(), Name: "Second", Host: "5.6.7.8" });
    expect(store.removeByMatchKey("1.2.3.4:28082|76561198000000000")).toBe(true);
    const listed = store.list();
    expect(listed).toHaveLength(1);
    expect(listed[0]!.Host).toBe("5.6.7.8");
    expect(store.removeByMatchKey("nope:1|2")).toBe(false);
  });

  it("rejects profiles missing required core fields", () => {
    const store = new ProfilesStore(dir, new PassthroughSecretCodec());
    expect(() => store.upsert({ Name: "broken" } as never)).toThrowError(/breached contract/);
  });
});

/** Deterministic non-identity codec to prove sealing actually happens at the file layer. */
class PrefixCodec extends PassthroughSecretCodec {
  override seal(plaintext: string): string {
    return plaintext ? SEALED_PREFIX + Buffer.from(plaintext, "utf8").toString("base64") : plaintext;
  }
  override open(blob: string): string {
    if (!blob.startsWith(SEALED_PREFIX)) return blob;
    return Buffer.from(blob.slice(SEALED_PREFIX.length), "base64").toString("utf8");
  }
}
