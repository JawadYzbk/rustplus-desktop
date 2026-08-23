/**
 * DeviceEventHub — port of the observable behavior of RustPlusClientReal's device/storage surface
 * (HookEventsIfNeeded / HandleStorageMonitorEvent / BuildAndStoreSnapshotFromStorageNode):
 *  - smart-switch state from entityChanged broadcasts → DeviceStateEvent(id, on, "SmartSwitch");
 *  - storage snapshots from entityInfo responses; tool-cupboard detection via hasProtection with a
 *    STICKY flag (a single event without the field must not demote a known TC back to a box);
 *  - after every TC snapshot a follow-up entityInfo pull is scheduled at 1500 ms ("Rust+ events are
 *    often one step behind"); storage-ish broadcasts trigger a 150 ms-delayed fresh pull.
 *
 * The 2.5.0 raw contract replaces the legacy reflection gymnastics: AppEntityPayload carries
 * value/items/capacity/hasProtection/protectionExpiry directly.
 */
import { EventEmitter } from "node:events";
import { rq } from "./protocol.js";

export interface StorageItem {
  itemId: number;
  amount: number;
  itemIsBlueprint: boolean;
}

export interface StorageSnapshot {
  upkeepSeconds: number | null;
  isToolCupboard: boolean;
  items: StorageItem[];
  capacity: number | null;
  snapshotUtc: number;
}

export type DeviceHubEvents =
  | { kind: "deviceState"; entityId: number; on: boolean; deviceType: "SmartSwitch" }
  | { kind: "storageSnapshot"; entityId: number; snapshot: StorageSnapshot };

/** Raw AppMessage broadcast → typed union, or null when not an entity change. */
export function extractEntityChanged(
  message: unknown,
): { entityId: number; payload: EntityPayload } | null {
  const m = (message ?? {}) as {
    broadcast?: { entityChanged?: { entityId?: unknown; payload?: unknown } };
  };
  const ec = m.broadcast?.entityChanged;
  if (!ec || typeof ec.entityId !== "number") return null;
  return { entityId: ec.entityId, payload: normalizePayload(ec.payload) };
}

export interface EntityPayload {
  value?: unknown;
  items?: Array<{ itemId?: unknown; quantity?: unknown; itemIsBlueprint?: unknown }>;
  capacity?: unknown;
  hasProtection?: unknown;
  protectionExpiry?: unknown;
}

function normalizePayload(p: unknown): EntityPayload {
  return (p ?? {}) as EntityPayload;
}

const STORAGE_TRIGGER_DELAY_MS = 150;
const TC_FOLLOWUP_PULL_MS = 1500;

export interface DeviceHubDeps {
  now?: () => number;
  /** Timers are injectable for tests (default setTimeout/clearTimeout). */
  setTimer?: (fn: () => void, ms: number) => unknown;
  clearTimer?: (handle: unknown) => void;
  /** Manager-level contract send used for follow-up pulls (getEntityInfo). */
  send: (data: Record<string, unknown>, timeoutMs?: number) => Promise<Record<string, unknown>>;
}

export class DeviceEventHub extends EventEmitter {
  private readonly cache = new Map<number, StorageSnapshot>();
  private readonly timers = new Map<number, unknown>();
  private readonly now: () => number;
  private readonly setTimer: (fn: () => void, ms: number) => unknown;
  private readonly clearTimer: (handle: unknown) => void;

  constructor(private readonly deps: DeviceHubDeps) {
    super();
    this.now = deps.now ?? (() => Date.now());
    this.setTimer = deps.setTimer ?? ((fn, ms) => setTimeout(fn, ms));
    this.clearTimer = deps.clearTimer ?? ((h) => clearTimeout(h as ReturnType<typeof setTimeout>));
  }

  /** Drop all cached snapshots + pending pulls (per-connection state dies with the connection). */
  reset(): void {
    for (const h of this.timers.values()) this.clearTimer(h);
    this.timers.clear();
    this.cache.clear();
  }

  tryGetCachedStorage(entityId: number): StorageSnapshot | undefined {
    return this.cache.get(entityId);
  }

  /** Entry point for AppBroadcast.entityChanged messages from the transport. */
  async handleEntityChanged(entityId: number, payload: EntityPayload): Promise<void> {
    const isStorageish =
      Array.isArray(payload.items) ||
      typeof payload.hasProtection === "boolean" ||
      typeof payload.protectionExpiry === "number";
    if (!isStorageish) {
      // Pure switch toggle → immediate device-state parity.
      this.emitDeviceState(entityId, payload);
      return;
    }
    // Storage-ish change: Rust+ fires one step ahead of server state — wait 150 ms then pull fresh.
    await this.delay(STORAGE_TRIGGER_DELAY_MS);
    await this.pullEntityInfo(entityId).catch(() => undefined);
  }

  /** Process an entityInfo RESPONSE payload (also used by scheduled follow-up pulls). */
  handleEntityInfoResponse(entityId: number, payload: EntityPayload): void {
    const isStorageish =
      Array.isArray(payload.items) ||
      typeof payload.hasProtection === "boolean" ||
      typeof payload.protectionExpiry === "number" ||
      typeof payload.capacity === "number";
    if (isStorageish) {
      this.buildAndStoreSnapshot(entityId, payload);
      return;
    }
    if (payload.value !== undefined) this.emitDeviceState(entityId, payload);
  }

  /** ScheduleEntityInfoPull parity: latest timer wins per entity. */
  schedulePull(entityId: number, delayMs = TC_FOLLOWUP_PULL_MS): void {
    const prev = this.timers.get(entityId);
    if (prev !== undefined) this.clearTimer(prev);
    const handle = this.setTimer(() => {
      this.timers.delete(entityId);
      void this.pullEntityInfo(entityId).catch(() => undefined);
    }, delayMs);
    this.timers.set(entityId, handle);
  }

  private async pullEntityInfo(entityId: number): Promise<void> {
    const res = await this.deps.send(rq.getEntityInfo(entityId));
    const info = res["entityInfo"] as { payload?: unknown } | undefined;
    if (info?.payload) {
      this.handleEntityInfoResponse(entityId, normalizePayload(info.payload));
    }
  }

  private emitDeviceState(entityId: number, payload: EntityPayload): void {
    if (typeof payload.value !== "boolean") return;
    this.emit("event", {
      kind: "deviceState",
      entityId,
      on: payload.value,
      deviceType: "SmartSwitch",
    } satisfies DeviceHubEvents);
  }

  /** BuildAndStoreSnapshotFromStorageNode parity, incl. the sticky tool-cupboard flag. */
  private buildAndStoreSnapshot(entityId: number, node: EntityPayload): void {
    const upkeepSeconds =
      typeof node.protectionExpiry === "number" ? node.protectionExpiry : null;

    let isToolCupboard = node.hasProtection === true;
    // Sticky TC: without this, one event lacking the optional field demotes a known TC to a box —
    // losing the upkeep mapping (legacy comment preserved).
    if (!isToolCupboard) {
      const old = this.cache.get(entityId);
      if (old?.isToolCupboard) isToolCupboard = true;
    }

    const items: StorageItem[] = (Array.isArray(node.items) ? node.items : []).map((it) => ({
      itemId: typeof it.itemId === "number" ? it.itemId : 0,
      amount: typeof it.quantity === "number" ? it.quantity : 0,
      itemIsBlueprint: it.itemIsBlueprint === true,
    }));

    const snap: StorageSnapshot = {
      upkeepSeconds: isToolCupboard ? upkeepSeconds : null, // boxes carry no upkeep
      isToolCupboard,
      items,
      capacity: typeof node.capacity === "number" ? node.capacity : null,
      snapshotUtc: this.now(),
    };

    this.cache.set(entityId, snap);
    this.emit("event", { kind: "storageSnapshot", entityId, snapshot: snap } satisfies DeviceHubEvents);

    if (isToolCupboard) this.schedulePull(entityId, TC_FOLLOWUP_PULL_MS);
  }

  private delay(ms: number): Promise<void> {
    return new Promise((resolve) => {
      this.setTimer(resolve, ms);
    });
  }
}
