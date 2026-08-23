/**
 * Device data layer — port of Services/Data/DeviceDataModule.cs (pure parts), Models/DeviceImportItem.cs
 * and the import-candidate collection from MainWindow.Devices.cs L2039-2123.
 *
 * JSON compatibility contract: legacy local overlay files are written by System.Text.Json with
 * default (PascalCase) property names. We PARSE both PascalCase (existing user files) and
 * camelCase, and SERIALIZE with PascalCase keys so files stay byte-compatible.
 */
import { randomUUID } from "node:crypto";

// --------------------------------------------------------------------- models

export interface ExportedDeviceDto {
  entityId: number;
  kind: string | null;
  name: string | null;
  alias: string | null;
  isGroup: boolean;
  children: ExportedDeviceDto[] | null;
  customIconId: number | null;
  customIconShortName: string | null;
  /** In-game alarm text. The cloud worker matches pushes against this. */
  inGameAlarmTitle: string | null;
  /** "SmallOilRig"/"LargeOilRig" when a rule uses this alarm as a rig trigger. */
  oilRigTrigger: string | null;
}

export interface OverlaySaveData {
  lastUpdatedUnix: number;
  /**
   * The server wipe these entity IDs belong to, Unix seconds. Zero for snapshots written
   * before this was recorded. Rust hands out a fresh net ID on every wipe while the server
   * key (ip-port) survives it — a stale snapshot lists devices that no longer exist.
   */
  wipeTimeUnix: number;
  devices: ExportedDeviceDto[];
}

export interface SmartDeviceNode {
  entityId: number;
  kind: string | null;
  name: string | null;
  alias: string | null;
  isGroup: boolean;
  children: SmartDeviceNode[];
  /** MapDtoToDevice quirk: groups arrive NOT missing, leaves always start missing. */
  isMissing: boolean;
  customIconId: number | null;
  customIconShortName: string | null;
  inGameAlarmTitle: string | null;
  oilRigTriggerTarget: string | null;
}

export function newSmartDevice(init: Partial<SmartDeviceNode> & Pick<SmartDeviceNode, "entityId">): SmartDeviceNode {
  return {
    kind: null,
    name: null,
    alias: null,
    isGroup: false,
    children: [],
    isMissing: false,
    customIconId: null,
    customIconShortName: null,
    inGameAlarmTitle: null,
    oilRigTriggerTarget: null,
    ...init,
  };
}

// ------------------------------------------------------------- JSON (de)serialization

type RawRecord = Record<string, unknown>;
const str = (v: unknown): string | null => (typeof v === "string" ? v : null);
const numOrNull = (v: unknown): number | null =>
  typeof v === "number" && Number.isFinite(v) ? v : null;

function pick(raw: RawRecord, ...keys: string[]): unknown {
  for (const k of keys) if (k in raw) return raw[k];
  return undefined;
}

/** Parses a device DTO from either PascalCase (legacy files) or camelCase. */
export function parseExportedDeviceDto(raw: unknown): ExportedDeviceDto {
  const r = (raw ?? {}) as RawRecord;
  const childrenRaw = pick(r, "Children", "children");
  const children =
    Array.isArray(childrenRaw) ? childrenRaw.map((c) => parseExportedDeviceDto(c)) : null;
  return {
    // uint.TryParse-falsy→0 parity: garbage IDs become 0 and are dropped by callers.
    entityId: typeof pick(r, "EntityId", "entityId") === "number" ? (pick(r, "EntityId", "entityId") as number) : 0,
    kind: str(pick(r, "Kind", "kind")),
    name: str(pick(r, "Name", "name")),
    alias: str(pick(r, "Alias", "alias")),
    isGroup: pick(r, "IsGroup", "isGroup") === true,
    children,
    customIconId: numOrNull(pick(r, "CustomIconId", "customIconId")),
    customIconShortName: str(pick(r, "CustomIconShortName", "customIconShortName")),
    inGameAlarmTitle: str(pick(r, "InGameAlarmTitle", "inGameAlarmTitle")),
    oilRigTrigger: str(pick(r, "OilRigTrigger", "oilRigTrigger")),
  };
}

/** Serializes with PascalCase keys exactly like System.Text.Json defaults did. */
export function serializeExportedDeviceDto(dto: ExportedDeviceDto): RawRecord {
  const out: RawRecord = {
    EntityId: dto.entityId,
    Kind: dto.kind,
    Name: dto.name,
    Alias: dto.alias,
    IsGroup: dto.isGroup,
    Children: null,
    CustomIconId: dto.customIconId,
    CustomIconShortName: dto.customIconShortName,
    InGameAlarmTitle: dto.inGameAlarmTitle,
    OilRigTrigger: dto.oilRigTrigger,
  };
  if (dto.children && dto.children.length > 0) {
    out.Children = dto.children.map(serializeExportedDeviceDto);
  }
  return out;
}

export function parseOverlaySaveData(json: string): OverlaySaveData {
  const r = JSON.parse(json) as RawRecord;
  return {
    lastUpdatedUnix: typeof pick(r, "LastUpdatedUnix", "lastUpdatedUnix") === "number" ? (pick(r, "LastUpdatedUnix", "lastUpdatedUnix") as number) : 0,
    wipeTimeUnix: typeof pick(r, "WipeTimeUnix", "wipeTimeUnix") === "number" ? (pick(r, "WipeTimeUnix", "wipeTimeUnix") as number) : 0,
    devices: Array.isArray(r.Devices)
      ? (r.Devices as unknown[]).map(parseExportedDeviceDto)
      : Array.isArray(r.devices)
        ? (r.devices as unknown[]).map(parseExportedDeviceDto)
        : [],
  };
}

// ------------------------------------------------------------------ mappings

/** DeviceDataModule.MapDeviceToDto — groups carry their children, leaves never do. */
export function mapDeviceToDto(d: SmartDeviceNode): ExportedDeviceDto {
  const dto: ExportedDeviceDto = {
    entityId: d.entityId,
    kind: d.kind,
    name: d.name,
    alias: d.alias,
    isGroup: d.isGroup,
    children: null,
    customIconId: d.customIconId,
    customIconShortName: d.customIconShortName,
    inGameAlarmTitle: d.inGameAlarmTitle,
    oilRigTrigger: d.oilRigTriggerTarget,
  };
  if (d.isGroup && d.children.length > 0) {
    dto.children = d.children.map(mapDeviceToDto);
  }
  return dto;
}

/** DeviceDataModule.MapDtoToDevice — IsMissing = !IsGroup (leaves start missing until probed). */
export function mapDtoToDevice(dto: ExportedDeviceDto): SmartDeviceNode {
  const dev = newSmartDevice({
    entityId: dto.entityId,
    kind: dto.kind,
    name: dto.name,
    alias: dto.alias,
    isGroup: dto.isGroup,
    isMissing: !dto.isGroup,
    customIconId: dto.customIconId,
    customIconShortName: dto.customIconShortName,
  });
  if (dto.children && dto.children.length > 0) {
    dev.children = dto.children.map(mapDtoToDevice);
  }
  return dev;
}

export function countActualDevicesTree(devices: readonly SmartDeviceNode[] | null): number {
  if (!devices) return 0;
  let count = 0;
  for (const d of devices) {
    count += d.isGroup ? countActualDevicesTree(d.children) : 1;
  }
  return count;
}

export function countActualDevicesDtos(dtos: readonly ExportedDeviceDto[] | null): number {
  if (!dtos) return 0;
  let count = 0;
  for (const d of dtos) {
    count += d.isGroup ? countActualDevicesDtos(d.children) : 1;
  }
  return count;
}

/** Recursive FindDeviceById parity over top-level devices and group children. */
export function findDeviceById(
  devices: readonly SmartDeviceNode[] | null | undefined,
  entityId: number,
): SmartDeviceNode | null {
  if (!devices) return null;
  for (const d of devices) {
    if (d.entityId === entityId) return d;
    if (d.children.length > 0) {
      const hit = findDeviceById(d.children, entityId);
      if (hit) return hit;
    }
  }
  return null;
}

/**
 * DeviceDataModule.GetTrimmedDeviceList — freemium ceiling. Groups are cloned with only as
 * many children as fit; a group that adds nothing is dropped entirely.
 */
export function getTrimmedDeviceList(dtos: readonly ExportedDeviceDto[], maxDevices: number): ExportedDeviceDto[] {
  const result: ExportedDeviceDto[] = [];
  let currentCount = 0;
  for (const dto of dtos) {
    if (currentCount >= maxDevices) break;
    if (dto.isGroup) {
      const [trimmed, added] = cloneGroupWithLimit(dto, maxDevices - currentCount);
      if (added > 0) {
        result.push(trimmed);
        currentCount += added;
      }
    } else {
      result.push(dto);
      currentCount++;
    }
  }
  return result;
}

function cloneGroupWithLimit(group: ExportedDeviceDto, remainingLimit: number): [ExportedDeviceDto, number] {
  let added = 0;
  const newGroup: ExportedDeviceDto = {
    entityId: group.entityId,
    kind: group.kind,
    name: group.name,
    alias: group.alias,
    isGroup: true,
    children: [],
    customIconId: null,
    customIconShortName: null,
    inGameAlarmTitle: null,
    oilRigTrigger: null,
  };
  for (const child of group.children ?? []) {
    if (added >= remainingLimit) break;
    if (child.isGroup) {
      const [sub, subAdded] = cloneGroupWithLimit(child, remainingLimit - added);
      if (subAdded > 0) {
        newGroup.children!.push(sub);
        added += subAdded;
      }
    } else {
      newGroup.children!.push(child);
      added++;
    }
  }
  return [newGroup, added];
}

// ------------------------------------------------------------- import picker

export interface TeamMemberRef {
  steamId: string; // u64 as string
  name: string | null;
}

export interface DeviceImportItem {
  ownerSteamId: string;
  ownerName: string;
  entityId: number;
  kind: string | null;
  name: string | null;
  alias: string | null;
  alreadyPresent: boolean;
  /** Snapshot predates the current wipe → entity IDs are dead. Visible but not auto-selected. */
  fromPreviousWipe: boolean;
  serverName: string;
  originalDto: ExportedDeviceDto | null;
  isSelected: boolean;
  existsState: "?" | "ok" | "missing" | "err" | "local";
  readonly id: string;

  /** Alias > Name > #ID display label: "Garage (#123456)" or bare "#123456". */
  displayName(): string;
  tagline(): string;
  extraInfo(): string;
  isSelectable(): boolean;
}

export function newImportItem(init: Pick<DeviceImportItem, "ownerSteamId" | "ownerName" | "entityId" | "originalDto"> & Partial<DeviceImportItem>): DeviceImportItem {
  const state: { existsState: DeviceImportItem["existsState"] } = { existsState: init.existsState ?? (init.alreadyPresent ? "local" : "?") };
  const item: DeviceImportItem = {
    id: randomUUID(),
    kind: null,
    name: null,
    alias: null,
    alreadyPresent: false,
    fromPreviousWipe: false,
    serverName: "",
    isSelected: false,
    ...init,
    get existsState() {
      return state.existsState;
    },
    set existsState(v) {
      state.existsState = v;
    },
    displayName() {
      const label = item.alias && item.alias.trim().length > 0 ? item.alias : item.name && item.name.trim().length > 0 ? item.name : null;
      return label ? `${label} (#${item.entityId})` : `#${item.entityId}`;
    },
    tagline() {
      return item.alreadyPresent ? `${item.ownerName} (already in your list)` : item.ownerName;
    },
    extraInfo() {
      switch (state.existsState) {
        case "ok":
          return "Reachable in current map";
        case "missing":
          return "Missing in current map";
        case "err":
          return "Status unknown";
        case "local":
          return "Already present";
        default:
          return "";
      }
    },
    isSelectable() {
      return !item.alreadyPresent;
    },
  };
  return item;
}

export interface ImportCollectorContext {
  existingDevices: readonly SmartDeviceNode[] | null;
  serverName: string;
}

/** AddDeviceToImportItems parity (L2039-2061). */
export function addDeviceToImportItems(
  items: DeviceImportItem[],
  d: ExportedDeviceDto,
  tm: TeamMemberRef,
  ctx: ImportCollectorContext,
  fromPreviousWipe = false,
): void {
  const already = findDeviceById(ctx.existingDevices, d.entityId) !== null;
  items.push(
    newImportItem({
      ownerSteamId: tm.steamId,
      ownerName: tm.name ?? tm.steamId,
      entityId: d.entityId,
      kind: d.kind,
      name: d.name,
      alias: d.alias,
      alreadyPresent: already,
      fromPreviousWipe,
      // Stale entries stay visible but nobody imports dead devices by accident.
      isSelected: !already && !fromPreviousWipe,
      existsState: already ? "local" : "?",
      serverName: ctx.serverName,
      originalDto: d, // full DTO including children
    }),
  );
}

/** CollectIndividualDevices parity — groups are transparent, only leaves land in the list. */
export function collectIndividualDevices(
  items: DeviceImportItem[],
  d: ExportedDeviceDto,
  tm: TeamMemberRef,
  ctx: ImportCollectorContext,
  fromPreviousWipe = false,
): void {
  if (d.isGroup) {
    for (const child of d.children ?? []) {
      collectIndividualDevices(items, child, tm, ctx, fromPreviousWipe);
    }
  } else {
    addDeviceToImportItems(items, d, tm, ctx, fromPreviousWipe);
  }
}

/**
 * IsSnapshotFromPreviousWipe parity (L2070-2080): WipeTimeUnix is the exact answer when present;
 * otherwise the write timestamp is the next best thing — a snapshot last written before the
 * wipe began cannot describe anything that exists now.
 */
export function isSnapshotFromPreviousWipe(data: OverlaySaveData | null, wipeTimeUtcMs: number | null): boolean {
  if (data === null || wipeTimeUtcMs === null) return false;
  const wipeUnix = Math.floor(wipeTimeUtcMs / 1000);
  if (data.wipeTimeUnix > 0) return data.wipeTimeUnix !== wipeUnix;
  return data.lastUpdatedUnix > 0 && data.lastUpdatedUnix < wipeUnix;
}
