/**
 * Device data layer golden tests — DTO round-trips (legacy PascalCase files), MapDtoToDevice
 * IsMissing quirk, trimming limits, import-candidate collection incl. previous-wipe staleness.
 */
import { describe, expect, it } from "vitest";
import {
  parseExportedDeviceDto,
  serializeExportedDeviceDto,
  parseOverlaySaveData,
  mapDeviceToDto,
  mapDtoToDevice,
  countActualDevicesTree,
  countActualDevicesDtos,
  findDeviceById,
  getTrimmedDeviceList,
  collectIndividualDevices,
  isSnapshotFromPreviousWipe,
  newImportItem,
  newSmartDevice,
  type DeviceImportItem,
  type ExportedDeviceDto,
} from "../src/main/services/devices/device-data.js";

const leaf = (entityId: number, over: Partial<ExportedDeviceDto> = {}): ExportedDeviceDto => ({
  entityId,
  kind: "SmartSwitch",
  name: `Switch ${entityId}`,
  alias: null,
  isGroup: false,
  children: null,
  customIconId: null,
  customIconShortName: null,
  inGameAlarmTitle: null,
  oilRigTrigger: null,
  ...over,
});

describe("DTO (de)serialization", () => {
  it("round-trips through PascalCase keys exactly like System.Text.Json defaults", () => {
    const dto = mapDeviceToDto(
      newSmartDevice({
        entityId: 42,
        kind: "StorageMonitor",
        name: "TC",
        alias: "Main TC",
        inGameAlarmTitle: "Base Alarm",
        oilRigTriggerTarget: "SmallOilRig",
      }),
    );
    const json = JSON.stringify(serializeExportedDeviceDto(dto));
    expect(json).toContain('"EntityId":42');
    expect(json).toContain('"InGameAlarmTitle":"Base Alarm"');
    const back = parseExportedDeviceDto(JSON.parse(json));
    expect(back.entityId).toBe(42);
    expect(back.alias).toBe("Main TC");
    expect(back.inGameAlarmTitle).toBe("Base Alarm");
    expect(back.oilRigTrigger).toBe("SmallOilRig");
  });

  it("parses legacy PascalCase overlay files AND camelCase variants", () => {
    const legacy = `{"LastUpdatedUnix":1700000000,"WipeTimeUnix":1699000000,"Devices":[{"EntityId":7,"Kind":"SmartSwitch","Name":"Door","Alias":null,"IsGroup":false,"Children":null,"CustomIconId":null,"CustomIconShortName":null}]}`;
    const data = parseOverlaySaveData(legacy);
    expect(data.devices).toHaveLength(1);
    expect(data.devices[0]!.entityId).toBe(7);

    const modern = `{"lastUpdatedUnix":1,"wipeTimeUnix":2,"devices":[{"entityId":8,"kind":"Switch"}]}`;
    expect(parseOverlaySaveData(modern).devices[0]!.entityId).toBe(8);

    // Garbage entity id → uint.TryParse parity → 0 (callers drop it).
    expect(parseExportedDeviceDto({ EntityId: "abc" }).entityId).toBe(0);
  });

  it("group children serialize only when the group has them; leaves never carry children", () => {
    const group = mapDeviceToDto({
      ...newSmartDevice({ entityId: 1, isGroup: true, name: "Lights" }),
      children: [newSmartDevice({ entityId: 2 }), newSmartDevice({ entityId: 3 })],
    });
    expect(group.children).toHaveLength(2);
    expect(JSON.parse(JSON.stringify(serializeExportedDeviceDto(group))).Children).toHaveLength(2);

    const emptyGroup = mapDeviceToDto(newSmartDevice({ entityId: 9, isGroup: true }));
    expect(emptyGroup.children).toBeNull();
    expect(serializeExportedDeviceDto(emptyGroup).Children).toBeNull();
  });
});

describe("MapDtoToDevice / counting / lookup", () => {
  it("leaves start missing, groups do not (IsMissing = !IsGroup)", () => {
    const tree = mapDtoToDevice({
      ...leaf(1),
      isGroup: true,
      children: [leaf(2), { ...leaf(3), isGroup: true, children: [leaf(4)] }],
    });
    expect(tree.isMissing).toBe(false);
    expect(tree.children![0]!.isMissing).toBe(true);
    expect(tree.children![1]!.children![0]!.isMissing).toBe(true);
  });

  it("counts actual devices recursively for both trees", () => {
    const tree = [
      newSmartDevice({ entityId: 1 }),
      {
        ...newSmartDevice({ entityId: 5, isGroup: true }),
        children: [newSmartDevice({ entityId: 6 }), newSmartDevice({ entityId: 7 })],
      },
      { ...newSmartDevice({ entityId: 9, isGroup: true }), children: [] },
    ];
    expect(countActualDevicesTree(tree)).toBe(3);
    expect(countActualDevicesDtos(tree.map(mapDeviceToDto))).toBe(3);
  });

  it("findDeviceById searches top level and nested groups", () => {
    const nested = newSmartDevice({ entityId: 10 });
    const devices = [
      newSmartDevice({ entityId: 1 }),
      { ...newSmartDevice({ entityId: 5, isGroup: true }), children: [nested] },
    ];
    expect(findDeviceById(devices, 10)).toBe(nested);
    expect(findDeviceById(devices, 99)).toBeNull();
    expect(findDeviceById(null, 1)).toBeNull();
  });
});

describe("GetTrimmedDeviceList (freemium ceiling)", () => {
  it("trims leaves first, clones groups to the remaining quota, drops empty groups", () => {
    const dtos = [
      leaf(1),
      { ...leaf(0), isGroup: true, children: [leaf(2), leaf(3)] },
      leaf(4),
    ];
    // Max 3: leaf 1 (1 slot) + group fits BOTH remaining children (2 slots) → leaf 4 dropped.
    const trimmed = getTrimmedDeviceList(dtos, 3);
    expect(countActualDevicesDtos(trimmed)).toBe(3);
    expect(trimmed[0]!.entityId).toBe(1);
    expect(trimmed[1]!.isGroup).toBe(true);
    expect(trimmed[1]!.children).toHaveLength(2);

    // Max 2 forces a PARTIAL group clone: only child 2 survives.
    const tighter = getTrimmedDeviceList(dtos, 2);
    expect(tighter[1]!.children).toHaveLength(1);
    expect(tighter[1]!.children![0]!.entityId).toBe(2);
  });

  it("nested groups count toward the limit and empty sub-groups are skipped", () => {
    const dtos = [
      {
        ...leaf(100),
        isGroup: true,
        children: [{ ...leaf(101), isGroup: true, children: [leaf(102), leaf(103)] }, leaf(104)],
      },
    ];
    const trimmed = getTrimmedDeviceList([dtos[0]!], 2);
    expect(countActualDevicesDtos(trimmed)).toBe(2); // 102, 103 — inner group filled the quota
    // A limit of 0 yields nothing at all.
    expect(getTrimmedDeviceList(dtos, 0)).toHaveLength(0);
  });
});

describe("import candidate collection", () => {
  const ctx = (existing: number[] = []) => ({
    existingDevices: existing.map((id) => newSmartDevice({ entityId: id })),
    serverName: "My Server",
  });

  it("groups are transparent; leaves land with owner metadata", () => {
    const items: DeviceImportItem[] = [];
    const dto = { ...leaf(0), isGroup: true, children: [leaf(11), leaf(12)] };
    collectIndividualDevices(items, dto, { steamId: "123", name: "Bob" }, ctx(), false);
    expect(items.map((i) => i.entityId)).toEqual([11, 12]);
    expect(items.every((i) => i.ownerName === "Bob" && i.serverName === "My Server")).toBe(true);
  });

  it("already-present entries are marked local and unselected; stale wipe entries visible but unselected", () => {
    const items: DeviceImportItem[] = [];
    collectIndividualDevices(items, leaf(21), { steamId: "1", name: "Me" }, ctx([21]), false);
    collectIndividualDevices(items, leaf(22), { steamId: "1", name: "Me" }, ctx(), true); // previous wipe
    collectIndividualDevices(items, leaf(23), { steamId: "1", name: "Me" }, ctx(), false);

    const byId = new Map(items.map((i) => [i.entityId, i]));
    expect(byId.get(21)!.alreadyPresent).toBe(true);
    expect(byId.get(21)!.isSelected).toBe(false);
    expect(byId.get(21)!.existsState).toBe("local");
    expect(byId.get(21)!.extraInfo()).toBe("Already present");
    expect(byId.get(22)!.fromPreviousWipe).toBe(true);
    expect(byId.get(22)!.isSelected).toBe(false); // dead IDs not auto-selected
    expect(byId.get(23)!.isSelected).toBe(true);
  });

  it("display label is Alias > Name > #ID with the id always shown", () => {
    expect(newImportItem({ ownerSteamId: "1", ownerName: "", entityId: 55, originalDto: null }).displayName()).toBe("#55");
    expect(
      newImportItem({ ownerSteamId: "1", ownerName: "", entityId: 55, originalDto: null, name: "Door Controller" }).displayName(),
    ).toBe("Door Controller (#55)");
    expect(
      newImportItem({ ownerSteamId: "1", ownerName: "", entityId: 55, originalDto: null, name: "Door", alias: "Garage" }).displayName(),
    ).toBe("Garage (#55)");
  });

  it("probe states map to legacy ExtraInfo strings", () => {
    const item = newImportItem({ ownerSteamId: "1", ownerName: "A", entityId: 1, originalDto: null });
    item.existsState = "ok";
    expect(item.extraInfo()).toBe("Reachable in current map");
    item.existsState = "missing";
    expect(item.extraInfo()).toBe("Missing in current map");
    item.existsState = "?";
    expect(item.extraInfo()).toBe("");
  });
});

describe("IsSnapshotFromPreviousWipe", () => {
  it("exact WipeTimeUnix comparison when present; LastUpdated fallback otherwise", () => {
    const wipeMs = Date.UTC(2024, 0, 10, 18, 0, 0); // 1704912000
    const exactMatch = { lastUpdatedUnix: 1, wipeTimeUnix: Math.floor(wipeMs / 1000), devices: [] };
    const exactStale = { lastUpdatedUnix: 1, wipeTimeUnix: Math.floor(wipeMs / 1000) - 3600, devices: [] };
    expect(isSnapshotFromPreviousWipe(exactMatch, wipeMs)).toBe(false);
    expect(isSnapshotFromPreviousWipe(exactStale, wipeMs)).toBe(true);

    // Old snapshots without WipeTimeUnix: written before the wipe began → stale.
    const oldSnapshot = { lastUpdatedUnix: Math.floor(wipeMs / 1000) - 60, wipeTimeUnix: 0, devices: [] };
    const freshSnapshot = { lastUpdatedUnix: Math.floor(wipeMs / 1000) + 60, wipeTimeUnix: 0, devices: [] };
    expect(isSnapshotFromPreviousWipe(oldSnapshot, wipeMs)).toBe(true);
    expect(isSnapshotFromPreviousWipe(freshSnapshot, wipeMs)).toBe(false);

    // No known wipe time → cannot judge.
    expect(isSnapshotFromPreviousWipe(exactStale, null)).toBe(false);
    expect(isSnapshotFromPreviousWipe(null, wipeMs)).toBe(false);
  });
});
