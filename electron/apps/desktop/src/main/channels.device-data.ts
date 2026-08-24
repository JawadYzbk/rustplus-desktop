import { readFileSync, writeFileSync } from "node:fs";
import { profileApplyImport, profileDeleteDevice, profileExportDevices, profileImportPreview } from "@rpd/shared";
import type { z } from "zod";
import {
  collectIndividualDevices,
  isSnapshotFromPreviousWipe,
  mapDeviceToDto,
  mapDtoToDevice,
  newImportItem,
  parseOverlaySaveData,
  type DeviceImportItem,
  type ExportedDeviceDto,
  type SmartDeviceNode,
} from "./services/devices/device-data.js";
import { parseDevices, serializeDevices } from "./services/devices/server-profile.js";

type ExportRequest = z.infer<typeof profileExportDevices["request"]>;
type PreviewRequest = z.infer<typeof profileImportPreview["request"]>;
type ApplyRequest = z.infer<typeof profileApplyImport["request"]>;
type DeleteRequest = z.infer<typeof profileDeleteDevice["request"]>;

export interface DeviceDataBridgeDeps {
  profiles: {
    list(): Array<{ Name: string; Host: string; Port: number; SteamId64: string }>;
    matchKey(profile: { Host: string; Port: number; SteamId64: string }): string;
    devicesFor(key: string): Record<string, unknown>[] | null;
    saveDevices(key: string, devices: Record<string, unknown>[]): boolean;
    field(key: string, name: string): unknown;
  };
  showSaveDialog(): Promise<{ canceled: boolean; filePath?: string }>;
  showOpenDialog(): Promise<{ canceled: boolean; filePaths: string[] }>;
}

export function buildDeviceDataHandlers(deps: DeviceDataBridgeDeps): {
  "profile/exportDevices": (request: ExportRequest) => Promise<{ saved: boolean; canceled: boolean; path: string | null; bytes: number }>;
  "profile/importPreview": (request: PreviewRequest) => Promise<z.infer<typeof profileImportPreview["response"]>>;
  "profile/applyImport": (request: ApplyRequest) => { saved: boolean; imported: number };
  "profile/deleteDevice": (request: DeleteRequest) => { removed: boolean; reason: "removed" | "notFound" | "notMissing" };
} {
  const knownProfile = (key: string): boolean =>
    deps.profiles.list().some((profile) => deps.profiles.matchKey(profile) === key);

  return {
    "profile/exportDevices": async (request) => {
      if (!knownProfile(request.matchKey)) return { saved: false, canceled: false, path: null, bytes: 0 };
      const selected = await deps.showSaveDialog();
      if (selected.canceled || !selected.filePath) return { saved: false, canceled: true, path: null, bytes: 0 };
      const raw = deps.profiles.devicesFor(request.matchKey);
      if (raw === null) return { saved: false, canceled: false, path: selected.filePath, bytes: 0 };
      const payload = {
        LastUpdatedUnix: Math.floor(Date.now() / 1000),
        WipeTimeUnix: profileWipeUnix(deps.profiles.field(request.matchKey, "WipeTime")),
        Devices: parseDevices(raw).map(mapDeviceToDto),
      };
      const json = `${JSON.stringify(payload, null, 2)}\n`;
      writeFileSync(selected.filePath, json, "utf8");
      return { saved: true, canceled: false, path: selected.filePath, bytes: Buffer.byteLength(json) };
    },

    "profile/importPreview": async (request) => {
      if (!knownProfile(request.matchKey)) return { canceled: false, path: null, candidates: [] };
      const selected = await deps.showOpenDialog();
      if (selected.canceled || !selected.filePaths[0]) return { canceled: true, path: null, candidates: [] };
      const path = selected.filePaths[0];
      const snapshot = parseOverlaySaveData(readFileSync(path, "utf8"));
      const existing = parseDevices(deps.profiles.devicesFor(request.matchKey) ?? []);
      const profile = deps.profiles.list().find((item) => deps.profiles.matchKey(item) === request.matchKey);
      const items: DeviceImportItem[] = [];
      const wipeTime = profileWipeUnix(deps.profiles.field(request.matchKey, "WipeTime"));
      const fromPreviousWipe = isSnapshotFromPreviousWipe(snapshot, wipeTime > 0 ? wipeTime * 1000 : null);
      const owner = { steamId: "local", name: "Imported snapshot" };
      for (const device of snapshot.devices) {
        collectIndividualDevices(items, device, owner, { existingDevices: existing, serverName: profile?.Name ?? "" }, fromPreviousWipe);
      }
      return {
        canceled: false,
        path,
        candidates: items.map(candidateDto),
      };
    },

    "profile/applyImport": (request) => {
      if (!knownProfile(request.matchKey)) return { saved: false, imported: 0 };
      const current = parseDevices(deps.profiles.devicesFor(request.matchKey) ?? []);
      const existingIds = new Set<number>();
      collectIds(current, existingIds);
      let imported = 0;
      for (const dto of request.devices) {
        if (existingIds.has(dto.entityId)) continue;
        current.push(mapDtoToDevice(dto as unknown as ExportedDeviceDto));
        collectIds([current[current.length - 1]!], existingIds);
        imported++;
      }
      return { saved: deps.profiles.saveDevices(request.matchKey, serializeDevices(current)), imported };
    },

    "profile/deleteDevice": (request) => {
      if (!knownProfile(request.matchKey)) return { removed: false, reason: "notFound" };
      const devices = parseDevices(deps.profiles.devicesFor(request.matchKey) ?? []);
      const hit = findDevice(devices, request.entityId);
      if (!hit) return { removed: false, reason: "notFound" };
      if (hit.isGroup || !hit.isMissing) return { removed: false, reason: "notMissing" };
      const removed = removeDevice(devices, request.entityId);
      return { removed: deps.profiles.saveDevices(request.matchKey, serializeDevices(devices)) && removed, reason: removed ? "removed" : "notFound" };
    },
  };
}

function candidateDto(item: DeviceImportItem): z.infer<typeof profileImportPreview["response"]>["candidates"][number] {
  return {
    id: item.id,
    ownerSteamId: item.ownerSteamId,
    ownerName: item.ownerName,
    entityId: item.entityId,
    kind: item.kind,
    name: item.name,
    alias: item.alias,
    alreadyPresent: item.alreadyPresent,
    fromPreviousWipe: item.fromPreviousWipe,
    serverName: item.serverName,
    existsState: item.existsState,
    originalDto: item.originalDto as z.infer<typeof profileApplyImport["request"]>["devices"][number],
  };
}

function profileWipeUnix(value: unknown): number {
  if (typeof value === "number" && Number.isFinite(value)) return Math.floor(value / 1000);
  if (typeof value === "string") {
    const time = Date.parse(value);
    if (Number.isFinite(time)) return Math.floor(time / 1000);
  }
  return 0;
}

function collectIds(devices: readonly SmartDeviceNode[], ids: Set<number>): void {
  for (const device of devices) {
    ids.add(device.entityId);
    collectIds(device.children, ids);
  }
}

function findDevice(devices: readonly SmartDeviceNode[], entityId: number): SmartDeviceNode | null {
  for (const device of devices) {
    if (device.entityId === entityId) return device;
    const nested = findDevice(device.children, entityId);
    if (nested) return nested;
  }
  return null;
}

function removeDevice(devices: SmartDeviceNode[], entityId: number): boolean {
  for (let index = 0; index < devices.length; index++) {
    const device = devices[index]!;
    if (device.entityId === entityId) {
      devices.splice(index, 1);
      return true;
    }
    if (removeDevice(device.children, entityId)) return true;
  }
  return false;
}
