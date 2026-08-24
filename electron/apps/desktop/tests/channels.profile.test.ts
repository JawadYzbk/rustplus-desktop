/**
 * profile/* channel handler tests — bridge between the lossless ProfilesStore and typed
 * renderer DTOs, exercised against a temp-dir store with a plaintext codec (sealing is
 * orthogonal here; covered by profiles-store tests).
 */
import { describe, expect, it } from "vitest";
import { mkdtempSync, readFileSync, rmSync, writeFileSync, mkdirSync } from "node:fs";
import { tmpdir } from "node:os";
import { join } from "node:path";
import { buildProfileHandlers } from "../src/main/channels.logic.js";
import { buildDeviceAutomationHandlers } from "../src/main/channels.device-automation.js";
import { buildDeviceDataHandlers } from "../src/main/channels.device-data.js";
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

describe("device automation handlers", () => {
  it("round-trips automation rules in the legacy PascalCase profile fields", () => {
    const dir = join(mkdtempSync(join(tmpdir(), "rpd-automation-")), "store");
    const store = makeStore(dir, LEGACY_PROFILES);
    const h = buildDeviceAutomationHandlers({ profiles: store });
    const key = "1.2.3.4:28082|76561198000000001";
    const rule = {
      id: "r1",
      name: "Night lights",
      isEnabled: true,
      isExpanded: true,
      conditionType: "GameTime" as const,
      playerMatchMode: "AnyOnline" as const,
      specificPlayerSteamId: "",
      locationEntityId: 0,
      distanceMeters: 250,
      startTime: "20:00",
      endTime: "08:00",
      targetEntityId: 10,
      matchedState: true,
      unmatchedState: false,
    };

    expect(h["deviceAutomation/saveRules"]({ matchKey: key, isActive: true, rules: [rule] })).toEqual({ saved: true });
    expect(store.field(key, "DeviceAutomationRules")).toEqual([
      expect.objectContaining({ Id: "r1", ConditionType: "GameTime", TargetEntityId: 10 }),
    ]);
    expect(h["deviceAutomation/getRules"]({ matchKey: key })).toMatchObject({
      found: true,
      isActive: true,
      rules: [expect.objectContaining({ id: "r1", conditionType: "GameTime" })],
    });
    rmSync(dir, { recursive: true, force: true });
  });
});

describe("device data handlers", () => {
  it("exports, previews, applies selected devices, and protects live-device deletion", async () => {
    const dir = join(mkdtempSync(join(tmpdir(), "rpd-device-data-")), "store");
    const store = makeStore(dir, LEGACY_PROFILES);
    const key = "1.2.3.4:28082|76561198000000001";
    const exportPath = join(dir, "devices.json");
    const importPath = join(dir, "import.json");
    writeFileSync(importPath, JSON.stringify({
      LastUpdatedUnix: 1,
      WipeTimeUnix: 0,
      Devices: [{
        EntityId: 99,
        Kind: "SmartSwitch",
        Name: "Imported switch",
        Alias: "Imported",
        IsGroup: false,
        Children: null,
        CustomIconId: null,
        CustomIconShortName: null,
        InGameAlarmTitle: null,
        OilRigTrigger: null,
      }],
    }), "utf8");
    const h = buildDeviceDataHandlers({
      profiles: store,
      showSaveDialog: async () => ({ canceled: false, filePath: exportPath }),
      showOpenDialog: async () => ({ canceled: false, filePaths: [importPath] }),
    });

    expect((await h["profile/exportDevices"]({ matchKey: key })).saved).toBe(true);
    expect(JSON.parse(readFileSync(exportPath, "utf8"))).toMatchObject({ Devices: expect.any(Array) });
    const preview = await h["profile/importPreview"]({ matchKey: key });
    expect(preview.candidates[0]).toMatchObject({ entityId: 99, alias: "Imported", alreadyPresent: false });
    expect(h["profile/applyImport"]({ matchKey: key, devices: [preview.candidates[0]!.originalDto] })).toEqual({ saved: true, imported: 1 });
    expect(store.devicesFor(key)?.some((device) => device.EntityId === 99)).toBe(true);

    expect(h["profile/deleteDevice"]({ matchKey: key, entityId: 22 })).toEqual({ removed: true, reason: "removed" });
    expect(h["profile/deleteDevice"]({ matchKey: key, entityId: 10 })).toEqual({ removed: false, reason: "notMissing" });
    rmSync(dir, { recursive: true, force: true });
  });
});
