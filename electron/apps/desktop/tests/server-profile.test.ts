/**
 * ServerProfile persistence golden tests — profiles.json round-trip contract against the C#
 * file format (PascalCase + deathMarkers exception), load-path validation parity, and the
 * SyncChatCommands mapping lifecycle incl. the NextFreeCommandIndex collision fix.
 */
import { describe, expect, it } from "vitest";
import {
  parseProfilesJson,
  serializeProfilesJson,
  parseServerProfile,
  validateCommand,
  matchKey,
  syncChatCommands,
  flattenAllDevices,
  newEmptyServerProfile,
} from "../src/main/services/devices/server-profile.js";
import { newSmartDevice } from "../src/main/services/devices/device-data.js";

/** A realistic slice of a legacy-written profiles.json. */
const LEGACY_JSON = JSON.stringify([
  {
    Host: "1.2.3.4",
    Port: 28082,
    SteamId64: "76561198000000001",
    PlayerToken: "-123456789",
    BattleMetricsId: null,
    LocalMapFilePath: null,
    LocalMapImagePath: "C:\\maps\\map.png",
    CustomMapUrl: null,
    IsConnected: false,
    IsFullConnected: false,
    UseFacepunchProxy: false,
    LastEventSource: "RustApi",
    Devices: [
      {
        EntityId: 100,
        Kind: "SmartSwitch",
        Name: "Smart Switch",
        Alias: "Turret Power",
        IsGroup: false,
        Children: [],
        IsMissing: false,
      },
      {
        EntityId: 200,
        Kind: null,
        Name: "Base",
        Alias: null,
        IsGroup: true,
        Children: [
          { EntityId: 201, Kind: "SmartSwitch", Name: "Smart Switch", Alias: null, IsGroup: false, Children: [], IsMissing: true },
          { EntityId: 202, Kind: "StorageMonitor", Name: "TC", Alias: "Main TC", IsGroup: false, Children: [], IsMissing: false },
        ],
      },
    ],
    CameraIds: [],
    deathMarkers: [{ x: 1, y: 2 }],
    LearnedDaySpeed: 0.24,
    LearnedNightSpeed: 1.2,
    ChatCommandsEnabled: true,
    CmdPop: "pop",
    CmdList: "commands",
    CmdTime: "time",
    CmdPromote: "!badcmd",
    CmdDeepSea: "123deepsea",
    CmdCargo: "cargo",
    CmdAfk: "afk",
    CmdOilRig: "oilrig",
    CmdHeli: "heli",
    CmdVendor: "vendor",
    CmdUpkeepDetail: "upkeepdetail",
    ChatCommandPrefix: ".",
    ChatCommandDelaySeconds: 9,
    ChatResponseDelaySeconds: 0.5,
    SwitchCommandMappings: [{ Label: "Switch 1", Command: "switch1", EntityId: 100 }],
    UpkeepCommandMappings: [],
    CmdCustomTimer: "timer",
    AlertCustomTimer: true,
    DiscordWebhookChatAlertsUrl: "",
    DiscordWebhookChatAlertsMention: "@everyone @here",
    DiscordWebhookChatAlertsEnabled: false,
    DiscordWebhookChatAlertsTts: false,
    DiscordWebhookChatAlertsExclusive: false,
    TimerAlarmEnabled: true,
    TimerAlarmAudioPath: null,
    TimerCountdownAudioPath: null,
    TimerAlarmSnoozeMinutes: -3,
    TimerAlarmBeepDurationSeconds: 5,
    CustomTimers: [
      {
        Id: "t1",
        Name: "Crate",
        Command: "crate",
        EndTimeUtc: "2024-01-10T18:00:00.000Z",
        EnableCountdownAudio: false,
        EnableAlarmAudio: true,
        CreatedNotified: true,
        Notified60: true,
        Notified30: true,
        Notified10: false,
        Notified3: false,
        CountdownAudioPlayed: false,
        AlarmPlayed: false,
        SnoozedUntilUtc: null,
        AutoDeleteAtUtc: null,
      },
    ],
    RustMapsMapId: null,
    RustMapsFetchTime: null,
    RustMapsWipeTime: "2024-01-10T18:00:00Z",
    WipeTime: "2024-01-05T18:00:00Z",
    LogicRules: [],
    IsLogicEngineActive: false,
    DeviceAutomationRules: [],
    IsDeviceAutomationActive: false,
    SubscribedTeammateSteamIds: [76561198000000002],
    SomeFutureProperty: { nested: true }, // written by a newer app version
  },
]);

describe("profiles.json round-trip", () => {
  it("parses the legacy PascalCase format with defaults applied through load-path validation", () => {
    const profiles = parseProfilesJson(LEGACY_JSON);
    expect(profiles).toHaveLength(1);
    const p = profiles[0]!;
    expect(p.host).toBe("1.2.3.4");
    expect(p.port).toBe(28082);
    expect(matchKey(p)).toBe("1.2.3.4:28082|76561198000000001");

    // Load-path validation parity (setters run during deserialization in C#):
    expect(p.cmdPromote).toBe("badcmd"); // leading '!' stripped
    expect(p.cmdDeepSea).toBe("deepsea"); // digit-leading rejected → default
    expect(p.chatCommandPrefix).toBe("."); // whitelisted prefix kept
    expect(p.chatCommandDelaySeconds).toBe(2); // 9 out of range → default
    expect(p.timerAlarmSnoozeMinutes).toBe(0); // clamped Math.Max(0, v)

    // deathMarkers keeps its special camelCase key:
    expect(p.deathMarkers).toEqual([{ x: 1, y: 2 }]);
    // Devices parsed recursively with groups intact:
    expect(p.devices[1]!.isGroup).toBe(true);
    expect(p.devices[1]!.children![1]!.alias).toBe("Main TC");
    // Timers: ISO strings become epoch ms:
    expect(p.customTimers[0]!.endTimeUtcMs).toBe(Date.UTC(2024, 0, 10, 18, 0, 0));
    // Unknown properties from other app versions survive:
    expect((p.extra as Record<string, unknown>).SomeFutureProperty).toEqual({ nested: true });
  });

  it("serializes back to a legacy-readable shape (PascalCase, deathMarkers exception)", () => {
    const p = parseProfilesJson(LEGACY_JSON)[0]!;
    const raw = JSON.parse(serializeProfilesJson([p]))[0] as Record<string, unknown>;
    expect(raw.Host).toBe("1.2.3.4");
    expect(raw.deathMarkers).toEqual([{ x: 1, y: 2 }]); // NOT DeathMarkers
    expect(raw.CmdPromote).toBe("badcmd");
    expect(raw.ChatCommandPrefix).toBe(".");
    expect(Array.isArray(raw.Devices)).toBe(true);
    expect((raw.Devices as unknown[])[0]).toMatchObject({ EntityId: 100, Alias: "Turret Power" });
    expect(((raw.CustomTimers as unknown[])[0] as Record<string, unknown>).EndTimeUtc).toBe("2024-01-10T18:00:00.000Z");
    expect(raw.SomeFutureProperty).toEqual({ nested: true });
  });

  it("re-parses its own output stably (parse ∘ serialize = idempotent)", () => {
    const once = parseProfilesJson(LEGACY_JSON);
    const twice = parseProfilesJson(serializeProfilesJson(once));
    expect(JSON.stringify(serializeProfilesJson(twice))).toBe(JSON.stringify(serializeProfilesJson(once)));
  });

  it("corrupt or non-array files yield an empty list, never a crash", () => {
    expect(parseProfilesJson("{oops")).toEqual([]);
    expect(parseProfilesJson('{"not":"an array"}')).toEqual([]);
    expect(parseProfilesJson("")).toEqual([]);
  });

  it("a fresh profile carries every C# default", () => {
    const p = newEmptyServerProfile();
    expect(p.port).toBe(28082);
    expect(p.cmdPop).toBe("pop");
    expect(p.cmdList).toBe("commands");
    expect(p.chatCommandPrefix).toBe("!");
    expect(p.chatCommandDelaySeconds).toBe(2);
    expect(p.chatResponseDelaySeconds).toBe(0.5);
    expect(p.alertCustomTimer).toBe(true);
    expect(p.timerAlarmEnabled).toBe(true);
    expect(p.learnedDaySpeed).toBeCloseTo(12 / 50);
    expect(p.learnedNightSpeed).toBeCloseTo(12 / 10);
  });

  it("validateCommand edge cases match the C# helper", () => {
    expect(validateCommand(null, "dflt")).toBe("dflt");
    expect(validateCommand("   ", "dflt")).toBe("dflt");
    expect(validateCommand("!go", "dflt")).toBe("go"); // only LEADING ! stripped
    expect(validateCommand("g!o", "dflt")).toBe("g!o"); // interior ! kept
    expect(validateCommand("9ball", "dflt")).toBe("dflt");
  });
});

describe("SyncChatCommands", () => {
  it("creates switchN mappings for unbound switches and drops mappings for vanished ones", () => {
    const p = newEmptyServerProfile();
    p.devices = [
      newSmartDevice({ entityId: 10, kind: "SmartSwitch" }),
      newSmartDevice({ entityId: 11, kind: "SmartSwitch" }),
    ];
    p.switchCommandMappings = [{ label: "Ghost", command: "switch7", entityId: 999 }];
    syncChatCommands(p);

    expect(p.switchCommandMappings.find((m) => m.entityId === 999)).toBeUndefined();
    expect(p.switchCommandMappings.map((m) => m.command)).toEqual(["switch1", "switch2"]);
    expect(p.switchCommandMappings.map((m) => m.label)).toEqual(["Switch 1", "Switch 2"]);
  });

  it("NextFreeCommandIndex collision fix: after deleting the middle upkeep, reuse is avoided", () => {
    const p = newEmptyServerProfile();
    p.devices = [
      newSmartDevice({ entityId: 50, kind: "StorageMonitor" }), // TC candidate
      newSmartDevice({ entityId: 51, kind: "StorageMonitor" }),
    ];
    // Pre-seed as if upkeep/upkeep2/upkeep3 existed and upkeep2 was deleted with its device.
    p.upkeepCommandMappings = [
      { label: "Upkeep", command: "upkeep", entityId: 40 },
      { label: "Upkeep 3", command: "upkeep3", entityId: 41 },
    ];
    syncChatCommands(p);
    // Count+1 would hand out "upkeep3" again; lowest-free index picks "upkeep2".
    const added = p.upkeepCommandMappings.filter((m) => m.entityId === 50 || m.entityId === 51);
    expect(added.map((m) => m.command)).toContain("upkeep2");
    expect(new Set(p.upkeepCommandMappings.map((m) => m.command)).size).toBe(
      p.upkeepCommandMappings.length,
    );
  });

  it("TC detection treats 'Storage Monitor' alias kind and bare monitors as candidates", () => {
    const p = newEmptyServerProfile();
    p.devices = [
      newSmartDevice({ entityId: 60, kind: "Storage Monitor" }), // legacy spaced kind
      newSmartDevice({ entityId: 61, kind: "StorageMonitor" }),
      newSmartDevice({ entityId: 62, kind: "StorageMonitor", storage: { itemsCount: 12 } } as never), // loot box → not TC
    ];
    syncChatCommands(p);
    const cmds = p.upkeepCommandMappings.map((m) => `${m.command}#${m.entityId}`);
    expect(cmds).toContain("upkeep#60");
    expect(cmds).toContain("upkeep2#61");
    expect(p.upkeepCommandMappings.some((m) => m.entityId === 62)).toBe(false);
  });

  it("group children are flattened for mapping purposes", () => {
    const p = newEmptyServerProfile();
    p.devices = [
      {
        ...newSmartDevice({ entityId: 1, isGroup: true }),
        children: [newSmartDevice({ entityId: 70, kind: "SmartSwitch" })],
      },
    ];
    syncChatCommands(p);
    expect(p.switchCommandMappings.map((m) => m.command)).toEqual(["switch1"]);
    expect(flattenAllDevices(p.devices)).toHaveLength(1); // group itself excluded
  });
});
