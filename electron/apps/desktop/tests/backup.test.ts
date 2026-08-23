/**
 * Backup/restore (v2 crypto) + granular reset tests — real fs, real node:crypto.
 */
import { afterEach, beforeEach, describe, expect, it } from "vitest";
import { existsSync, mkdirSync, mkdtempSync, readFileSync, rmSync, writeFileSync } from "node:fs";
import { tmpdir } from "node:os";
import { join } from "node:path";
import { BACKUP_MAGIC } from "@rpd/shared";
import { unzipSync, strFromU8 } from "fflate";
import type { ResetTarget } from "@rpd/shared";
import { BackupService, decrypt, encrypt } from "../src/main/services/backup-service.js";
import { performReset } from "../src/main/services/reset-service.js";
import { SettingsStore } from "../src/main/stores/settings-store.js";
import { ProfilesStore } from "../src/main/stores/profiles-store.js";
import { PassthroughSecretCodec, SEALED_PREFIX } from "../src/main/stores/secret-codec.js";

let root = "";
let backupsDir = "";

beforeEach(() => {
  root = mkdtempSync(join(tmpdir(), "rpd-backup-"));
  backupsDir = join(root, "backups");
  // Seed stores through their own APIs so every file is a valid store document.
  new ProfilesStore(root, new PassthroughSecretCodec()).upsert({
    Name: "Main",
    Host: "1.2.3.4",
    Port: 28082,
    SteamId64: "s",
    PlayerToken: SEALED_PREFIX + "abc",
  });
  new SettingsStore(root).patch({ LastHost: "x" });
  mkdirSync(join(root, "cache"), { recursive: true });
  writeFileSync(join(root, "cache", "minimap_settings.json"), JSON.stringify({ zoom: 2 }));
});

afterEach(() => {
  rmSync(root, { recursive: true, force: true });
});

const service = () => new BackupService(root, backupsDir);

describe("BackupService", () => {
  it("creates an unencrypted zip containing stores + manifest with sha256 entries", () => {
    const res = service().create(null);
    expect(res.encrypted).toBe(false);
    expect(existsSync(res.path)).toBe(true);

    const zip = unzipSync(new Uint8Array(readFileSync(res.path)));
    const manifest = JSON.parse(strFromU8(zip["manifest.json"]!)) as { format: number; files: Array<{ path: string }> };
    expect(manifest.format).toBe(2);
    expect(manifest.files.map((f) => f.path)).toContain("profiles.json");
    expect(zip["cache/minimap_settings.json"]).toBeDefined();
  });

  it("encrypted backup starts with RPDENC2 and round-trips restore", () => {
    // Mutate after create so restore must actually overwrite the live file.
    const res = service().create("hunter2");
    const head = readFileSync(res.path).subarray(0, BACKUP_MAGIC.length).toString("latin1");
    expect(head).toBe(BACKUP_MAGIC);

    writeFileSync(join(root, "tracking_settings.json"), "CLOBBERED");

    const out = service().restore(res.path, "hunter2");
    expect(out.restored).toContain("profiles.json");
    expect(out.restored).toContain("tracking_settings.json");
    const doc = JSON.parse(readFileSync(join(root, "tracking_settings.json"), "utf8")) as Record<string, unknown>;
    expect(doc["LastHost"]).toBe("x");
  });

  it("wrong password throws (GCM auth tag)", () => {
    const res = service().create("correct-horse");
    expect(() => service().restore(res.path, "wrong")).toThrowError();
  });

  it("tampering with encrypted bytes throws on restore", () => {
    const res = service().create("pw");
    const buf = readFileSync(res.path);
    const idx = buf.length - 5;
    buf[idx]! ^= 0xff;
    writeFileSync(res.path, buf);
    expect(() => service().restore(res.path, "pw")).toThrowError();
  });

  it("restore rejects password for plain zips and missing manifest", () => {
    const plain = service().create(null);
    expect(() => service().restore(plain.path, "pw")).toThrowError(/not encrypted/);
    // A zip without our manifest is rejected.
    const fake = join(backupsDir, "fake.zip");
    writeFileSync(fake, Buffer.from([0x50, 0x4b, 0x05, 0x06, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0]));
    expect(() => service().restore(fake)).toThrowError(/manifest\.json missing/);
  });

  it("encrypt/decrypt primitives round-trip arbitrary bytes", () => {
    const data = new Uint8Array([1, 2, 3, 250, 251]);
    const blob = encrypt(data, "k");
    expect(Buffer.from(blob).toString("latin1").startsWith(BACKUP_MAGIC)).toBe(true);
    expect(decrypt(blob, "k")).toEqual(data);
  });
});

describe("performReset", () => {
  const ctx = () => ({
    userDataDir: root,
    settings: new SettingsStore(root),
    profiles: new ProfilesStore(root, new PassthroughSecretCodec()),
    log: () => undefined,
  });

  it("steam target clears SteamId64 only", () => {
    new SettingsStore(root).patch({ SteamId64: "765", AutoConnectEnabled: true });
    performReset(ctx(), ["steam"]);
    const s = new SettingsStore(root).all;
    expect(s.SteamId64).toBe("");
    expect(s.AutoConnectEnabled).toBe(true); // other keys untouched
  });

  it("pairing + crosshairs delete their files", () => {
    writeFileSync(join(root, "rustplusjs-config.json"), "{}");
    writeFileSync(join(root, "custom_crosshairs.json"), "[]");
    const out = performReset(ctx(), ["pairing", "crosshairs"]);
    expect(out.performed).toEqual(["pairing", "crosshairs"]);
    expect(existsSync(join(root, "rustplusjs-config.json"))).toBe(false);
    expect(existsSync(join(root, "custom_crosshairs.json"))).toBe(false);
  });

  it("profiles clears the list; cache also removes Overlays\\ (legacy side effect)", () => {
    new ProfilesStore(root, new PassthroughSecretCodec()).upsert({
      Name: "A",
      Host: "h",
      Port: 1,
      SteamId64: "s",
      PlayerToken: "",
    });
    mkdirSync(join(root, "Overlays"), { recursive: true });

    const out = performReset(ctx(), ["profiles", "cache"]);
    expect(out.performed).toEqual(["profiles", "cache"]);
    // Fresh instance: reads from disk, no stale cache.
    expect(new ProfilesStore(root, new PassthroughSecretCodec()).list()).toHaveLength(0);
    expect(existsSync(join(root, "cache"))).toBe(false);
    expect(existsSync(join(root, "Overlays"))).toBe(false);
  });

  it("three_d wipes 3DMaps; unknown targets are compile-time impossible", () => {
    mkdirSync(join(root, "3DMaps"), { recursive: true });
    performReset(ctx(), ["three_d"]);
    expect(existsSync(join(root, "3DMaps"))).toBe(false);
    const targets: ResetTarget[] = ["profiles", "steam", "pairing", "crosshairs", "cache", "three_d"];
    expect(targets).toHaveLength(6);
  });
});
