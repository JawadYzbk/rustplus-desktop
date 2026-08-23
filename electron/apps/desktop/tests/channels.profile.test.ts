/**
 * profile/* channel handler tests — bridge between the lossless ProfilesStore and typed
 * renderer DTOs, exercised against a temp-dir store with a plaintext codec (sealing is
 * orthogonal here; covered by profiles-store tests).
 */
import { describe, expect, it } from "vitest";
import { mkdtempSync, rmSync, writeFileSync, mkdirSync } from "node:fs";
import { tmpdir } from "node:os";
import { join } from "node:path";
import { buildProfileHandlers } from "../src/main/channels.logic.js";
import { ProfilesStore } from "../src/main/stores/profiles-store.js";
import type { SecretCodec } from "../src/main/stores/secret-codec.js";

const passthroughCodec: SecretCodec = {
  seal: (v) => v,
  open: (v) => v,
};

function makeStore(dir: string, legacyDoc: unknown): ProfilesStore {
  mkdirSync(dir, { recursive: true });
  // JsonStore expects the NEW wrapped format ({schemaVersion, profiles}) — same as the migrator writes.
  const doc = legacyDoc as { schemaVersion?: number; profiles?: unknown[] };
  const payload = Array.isArray(legacyDoc)
    ? { schemaVersion: 1, profiles: legacyDoc }
    : doc;
  writeFileSync(join(dir, "profiles.json"), JSON.stringify(payload), "utf8");
  return new ProfilesStore(dir, passthroughCodec);
}

const LEGACY_PROFILES = [
  {
    Name: "Main",
    Host: "1.2.3.4",
    Port: 28082,
    SteamId64: "76561198000000001",
    PlayerToken: "-111222333",
    Devices: [
      { EntityId: 10, Kind: "SmartSwitch", Name: "Smart Switch", Alias: "Turret", IsGroup: false, Children: [], IsMissing: false },
      {
        EntityId: 20,
        Kind: null,
        Name: "Base",
        Alias: null,
        IsGroup: true,
        IsMissing: false,
        Children: [
          { EntityId: 21, Kind: "StorageMonitor", Name: "TC", Alias: "Main TC", IsGroup: false, Children: [], IsMissing: false },
          { EntityId: 22, Kind: "SmartSwitch", Name: "Lights", Alias: null, IsGroup: false, Children: [], IsMissing: true },
        ],
      },
    ],
    SomeFutureField: true,
  },
];

describe("profile handlers", () => {
  it("profile/list returns summaries with recursive device counts, no tokens", () => {
    const dir = join(mkdtempSync(join(tmpdir(), "rpd-prof-")), "store");
    const store = makeStore(dir, LEGACY_PROFILES);
    const h = buildProfileHandlers({ profiles: store, activeRef: { key: null } });

    const { profiles } = h["profile/list"]();
    expect(profiles).toHaveLength(1);
    const p = profiles[0]!;
    expect(p.name).toBe("Main");
    expect(p.matchKey).toBe("1.2.3.4:28082|76561198000000001");
    expect(p.deviceCount).toBe(3); // groups don't count
    expect(JSON.stringify(profiles)).not.toContain("PlayerToken");
    rmSync(dir, { recursive: true, force: true });
  });

  it("getDevices parses the PascalCase tree into camelCase DTOs; unknown matchKey → found:false", () => {
    const dir = join(mkdtempSync(join(tmpdir(), "rpd-prof-")), "store");
    const store = makeStore(dir, LEGACY_PROFILES);
    const h = buildProfileHandlers({ profiles: store, activeRef: { key: null } });
    const key = "1.2.3.4:28082|76561198000000001";

    const ok = h["profile/getDevices"]({ matchKey: key });
    expect(ok.found).toBe(true);
    expect(ok.devices).toHaveLength(2);
    expect(ok.devices[0]).toMatchObject({ entityId: 10, alias: "Turret", isGroup: false, children: [] });
    expect(ok.devices[1]!.children![0]).toMatchObject({ entityId: 21, alias: "Main TC" });

    const miss = h["profile/getDevices"]({ matchKey: "nope" });
    expect(miss).toEqual({ devices: [], found: false });
    rmSync(dir, { recursive: true, force: true });
  });

  it("saveDevices writes the whole tree back in legacy PascalCase and survives a reload", () => {
    const dir = join(mkdtempSync(join(tmpdir(), "rpd-prof-")), "store");
    const store = makeStore(dir, LEGACY_PROFILES);
    const h = buildProfileHandlers({ profiles: store, activeRef: { key: null } });
    const key = "1.2.3.4:28082|76561198000000001";

    const saved = h["profile/saveDevices"]({
      matchKey: key,
      devices: [
        {
          entityId: 30,
          kind: "SmartSwitch",
          name: "New Switch",
          alias: null,
          isGroup: false,
          children: [],
          isMissing: true,
          customIconId: null,
          customIconShortName: null,
          inGameAlarmTitle: null,
          oilRigTriggerTarget: null,
        },
      ],
    });
    expect(saved.saved).toBe(true);

    // Reload from disk through a fresh store instance — the write must be durable.
    const store2 = new ProfilesStore(dir, passthroughCodec);
    const raw = store2.devicesFor(key)!;
    expect(raw[0]).toMatchObject({ EntityId: 30, Kind: "SmartSwitch", IsMissing: true });
    expect(h["profile/getDevices"]({ matchKey: key }).devices[0]!.entityId).toBe(30);

    // Unknown matchKey → not saved.
    expect(h["profile/saveDevices"]({ matchKey: "gone", devices: [] }).saved).toBe(false);
    rmSync(dir, { recursive: true, force: true });
  });
});
