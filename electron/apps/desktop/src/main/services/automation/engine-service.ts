/**
 * LogicEngineService — the main-process host binding the headless LogicEngine to live state:
 *  - rules/flags/timers read+written through the lossless ProfilesStore (rules persisted as
 *    legacy PascalCase via parseLogicRule/serializeLogicRule);
 *  - device on/off cache seeded from stored trees and updated from hub device-state events;
 *  - refresh = sequential getEntityInfo pulls fed back through the hub (parity with
 *    RefreshAllDevicesStatusAsync's per-device handling);
 *  - toggles ride the rate-limited ConnectionManager.setEntityValue;
 *  - MonumentWatcher rig timers are a documented seam: the watcher lands with stage 6 maps,
 *    so TriggerExternal currently reports success and logs.
 */
import { randomUUID } from "node:crypto";
import {
  LogicEngine,
  type CustomTimerInput,
  type EngineDevice,
  type LogicEngineHost,
} from "./logic-engine.js";
import { parseDevices, serializeDevices } from "../devices/server-profile.js";
import { findDeviceById, type SmartDeviceNode } from "../devices/device-data.js";
import {
  parseLogicRule,
  serializeLogicRule,
  type LogicRule,
} from "./logic-rule.js";
import { rq } from "../rustplus/protocol.js";

export interface EngineProfilesAdapter {
  /** matchKey of the profile the engine operates on, or null when none is selected. */
  activeKey(): string | null;
  field(key: string, name: string): unknown;
  setField(key: string, name: string, value: unknown): boolean;
  devicesFor(key: string): Record<string, unknown>[] | null;
  saveDevices(key: string, devices: Record<string, unknown>[]): boolean;
}

export interface EngineConnAdapter {
  isConnected(): boolean;
  setEntityValue(entityId: number, value: boolean): Promise<void>;
  getEntityInfo(entityId: number): Promise<Record<string, unknown>>;
  sendTeamMessage(message: string): Promise<void>;
}

export interface EngineHubAdapter {
  /** Feed a getEntityInfo response through the hub so snapshots/sticky-TC stay consistent. */
  handleEntityInfoResponse(entityId: number, entityPayload: unknown): void;
  onDeviceState(listener: (e: { entityId: number; value: boolean }) => void): () => void;
}

/** Builds the hub adapter from the real DeviceEventHub (EventEmitter "event" → kind discriminator). */
export function hubAdapter(
  hub: {
    handleEntityInfoResponse(entityId: number, entityPayload: unknown): void;
    on(event: string, listener: (...args: unknown[]) => void): unknown;
    off(event: string, listener: (...args: unknown[]) => void): unknown;
  },
): EngineHubAdapter {
  return {
    handleEntityInfoResponse: (entityId, entityPayload) =>
      hub.handleEntityInfoResponse(entityId, entityPayload),
    onDeviceState: (listener) => {
      const wrapped = (...args: unknown[]): void => {
        const p = args[0] as { kind?: string; entityId?: number; on?: boolean } | undefined;
        if (p?.kind === "deviceState" && typeof p.entityId === "number") {
          listener({ entityId: p.entityId, value: p.on === true });
        }
      };
      hub.on("event", wrapped);
      return () => hub.off("event", wrapped);
    },
  };
}

interface RuntimeStatus {
  isRunning: boolean;
  currentRuleName: string | null;
  currentStepNumber: number;
  currentStepType: string | null;
  pendingRules: string[];
}

export class LogicEngineService {
  private readonly engine: LogicEngine;
  private readonly deviceStates = new Map<number, boolean>();
  private unsubHub: (() => void) | null = null;

  constructor(
    private readonly profiles: EngineProfilesAdapter,
    private readonly conn: EngineConnAdapter,
    private readonly hub: EngineHubAdapter,
    private readonly log: (message: string) => void = () => undefined,
  ) {
    const host: LogicEngineHost = {
      isEngineActive: () => this.withKey((key) => this.profiles.field(key, "IsLogicEngineActive") === true) ?? false,
      rules: () => this.loadRules(),
      findDevice: (entityId) => this.findDevice(entityId),
      findGroupSwitches: (groupName) => this.findGroupSwitches(groupName),
      isConnected: () => this.conn.isConnected(),
      toggleSmartSwitch: async (entityId, on) => {
        await this.conn.setEntityValue(entityId, on);
        this.deviceStates.set(entityId, on);
      },
      refreshAllDevices: () => this.refreshAllDevices(),
      triggerRigTimer: (rigName, seconds, showCrate) => {
        // Stage-6 seam: MonumentWatcher.TriggerExternal parity arrives with the map stage.
        this.log(
          `[LogicEngine] rig timer requested: ${rigName} ${seconds}s showCrate=${showCrate} (MonumentWatcher not yet ported — accepted, no marker).`,
        );
        return true;
      },
      customTimerCount: () => this.customTimers().length,
      removeCustomTimerByName: (name) => {
        const key = this.profiles.activeKey();
        if (!key) return;
        const timers = this.customTimers().filter((t) => t.name.toLowerCase() !== name.toLowerCase());
        this.saveCustomTimers(key, timers);
      },
      addCustomTimer: (t) => {
        const key = this.profiles.activeKey();
        if (!key) return;
        const timers = [...this.customTimers(), { ...t, id: randomUUID() }];
        this.saveCustomTimers(key, timers);
      },
      alertCustomTimers: () => this.withKey((key) => this.profiles.field(key, "AlertCustomTimer") !== false) ?? true,
      sendTeamChat: (message) => {
        void this.conn.sendTeamMessage(message).catch((err: unknown) =>
          this.log(`[LogicEngine] team chat send failed: ${String(err instanceof Error ? err.message : err)}`),
        );
      },
      chatAlert: (message) => {
        this.log(`[chat-alert] ${message}`);
        // chatAlert is sendTeamChat with a local log line (MainWindow parity).
        void this.conn.sendTeamMessage(message).catch((err: unknown) =>
          this.log(`[LogicEngine] chat alert send failed: ${String(err instanceof Error ? err.message : err)}`),
        );
      },
      log: (m) => this.log(m),
      isBusy: () => false, // manual-intervention mutexes arrive with the device-control UI
    };
    this.engine = new LogicEngine(host);

    this.unsubHub = hub.onDeviceState((e) => {
      this.deviceStates.set(e.entityId, e.value);
      this.engine.onDeviceEvent(e.entityId, e.value);
    });
  }

  dispose(): void {
    this.unsubHub?.();
    this.unsubHub = null;
  }

  /** Manual device-event entry (tests / future pipelines that bypass the hub subscription). */
  onDeviceEvent(entityId: number, isOn: boolean): void {
    this.engine.onDeviceEvent(entityId, isOn);
  }

  /** RefreshAllDevicesStatusAsync parity: sequential getEntityInfo per flat node, hub-consistent. */
  private async refreshAllDevices(): Promise<void> {
    for (const id of this.flatEntityIds()) {
      try {
        const res = await this.conn.getEntityInfo(id);
        // Feed the raw entity payload back through the hub so snapshots stay consistent.
        const entityPayload = (res as { entityInfo?: unknown })["entityInfo"];
        if (entityPayload !== undefined) this.hub.handleEntityInfoResponse(id, entityPayload);
      } catch (err) {
        // Unreachable device → mark missing rather than failing the whole rule.
        this.log(`[LogicEngine] probe #${id} failed: ${String(err instanceof Error ? err.message : err)}`);
        const node = this.findNode(id);
        if (node) node.isMissing = true;
      }
    }
  }

  status(): RuntimeStatus & { activeKey: string | null } {
    return {
      activeKey: this.profiles.activeKey(),
      isRunning: this.engine.isRunning,
      currentRuleName: this.engine.currentRuleName,
      currentStepNumber: this.engine.currentStepNumber,
      currentStepType: this.engine.currentStepType,
      pendingRules: [...this.engine.pending],
    };
  }

  requestStop(): void {
    this.engine.requestStop();
  }

  async runRule(ruleId: string): Promise<boolean> {
    const rule = this.loadRules().find((r) => r.id === ruleId);
    if (!rule) return false;
    void this.engine.enqueue(rule);
    return true;
  }

  /** Manual chat-command entry point (renderer console + future chat pipeline share this). */
  onChatCommand(text: string): void {
    this.engine.onChatCommand(text);
  }

  // ------------------------------------------------------------------ internals

  private withKey<T>(fn: (key: string) => T): T | null {
    const key = this.profiles.activeKey();
    return key ? fn(key) : null;
  }

  loadRules(): LogicRule[] {
    return this.withKey((key) => {
      const raw = this.profiles.field(key, "LogicRules");
      return Array.isArray(raw) ? raw.map((r) => parseLogicRule((r ?? {}) as Record<string, unknown>)) : [];
    }) ?? [];
  }

  saveRules(rules: LogicRule[]): boolean {
    const key = this.profiles.activeKey();
    if (!key) return false;
    return this.profiles.setField(key, "LogicRules", rules.map(serializeLogicRule));
  }

  /** Rule records for an ARBITRARY profile (renderer may target a non-active profile). */
  rulesFor(matchKey: string): unknown[] {
    const raw = this.profiles.field(matchKey, "LogicRules");
    return Array.isArray(raw) ? raw : [];
  }

  isEngineActiveFor(matchKey: string): boolean {
    return this.profiles.field(matchKey, "IsLogicEngineActive") === true;
  }

  /** Full rule (incl. steps) for the step editor; null when unknown.
   * Matching happens AFTER parsing because stored records are PascalCase (legacy format). */
  ruleFor(matchKey: string, ruleId: string): LogicRule | null {
    return (
      this.rulesFor(matchKey)
        .map((r) => parseLogicRule((r ?? {}) as Record<string, unknown>))
        .find((r) => r.id === ruleId) ?? null
    );
  }

  /** Wholesale rule save (header + steps); appends when the id is new. */
  saveFullRuleFor(matchKey: string, rule: LogicRule): boolean {
    const existing = this.rulesFor(matchKey);
    const parsed = existing.map((r) => parseLogicRule((r ?? {}) as Record<string, unknown>));
    const idx = parsed.findIndex((r) => r.id === rule.id);
    if (idx >= 0) parsed[idx] = rule;
    else parsed.push(rule);
    return this.profiles.setField(
      matchKey,
      "LogicRules",
      parsed.map(serializeLogicRule),
    );
  }

  /** Header-level save: steps of unchanged rule ids are preserved verbatim; new ids get none. */
  saveRulesFor(
    matchKey: string,
    headers: Array<Record<string, unknown>>,
    isEngineActive: boolean,
  ): boolean {
    const existing = this.rulesFor(matchKey).map((r) => parseLogicRule((r ?? {}) as Record<string, unknown>));
    const byId = new Map(existing.map((r) => [r.id, r]));
    const merged: unknown[] = [];
    for (const h of headers) {
      const prev = byId.get(String(h["id"]));
      const rule: LogicRule = {
        id: String(h["id"]),
        name: String(h["name"] ?? "New Rule"),
        isEnabled: h["isEnabled"] === true,
        isLoopEnabled: h["isLoopEnabled"] === true,
        loopCount: Math.max(0, Math.trunc(Number(h["loopCount"] ?? 1)) || 0),
        triggerType: (h["triggerType"] as LogicRule["triggerType"]) ?? "SmartAlarm",
        triggerEntityId: Math.max(0, Math.trunc(Number(h["triggerEntityId"] ?? 0)) || 0),
        triggerCommand: String(h["triggerCommand"] ?? ""),
        triggerRuleId: String(h["triggerRuleId"] ?? ""),
        triggerState: h["triggerState"] !== false,
        conditionOperator: (h["conditionOperator"] as LogicRule["conditionOperator"]) ?? "NONE",
        conditionDeviceEntityId: Math.max(0, Math.trunc(Number(h["conditionDeviceEntityId"] ?? 0)) || 0),
        conditionDeviceState: h["conditionDeviceState"] !== false,
        // Steps survive a header-only edit untouched:
        steps: prev?.steps ?? [],
      };
      merged.push(serializeLogicRule(rule));
    }
    let ok = this.profiles.setField(matchKey, "LogicRules", merged);
    ok = this.profiles.setField(matchKey, "IsLogicEngineActive", isEngineActive === true) && ok;
    return ok;
  }

  /** Recursive find with live state — shared with the chat-command dispatcher. */
  findDeviceFor(entityId: number): EngineDevice | null {
    return this.findDevice(entityId);
  }

  /** Timer list for chat consumers ({name, command, endTimeUtcMs}). */
  timersForChat(): Array<{ name: string; command: string; endTimeUtcMs: number }> {
    return this.customTimers().map((t) => ({
      name: t.name,
      command: t.command,
      endTimeUtcMs: t.endUtc,
    }));
  }

  /** Adds a timer from the chat pipeline (id assigned here). */
  addTimerFromChat(t: {
    name: string;
    command: string;
    endTimeUtcMs: number;
    createdNotified: boolean;
    notified60: boolean;
    notified30: boolean;
    notified10: boolean;
    notified3: boolean;
  }): void {
    const key = this.profiles.activeKey();
    if (!key) return;
    this.saveCustomTimers(key, [
      ...this.customTimers(),
      {
        id: randomUUID(),
        name: t.name,
        command: t.command,
        endUtc: t.endTimeUtcMs,
        createdNotified: t.createdNotified,
        notified60: t.notified60,
        notified30: t.notified30,
        notified10: t.notified10,
        notified3: t.notified3,
      },
    ]);
  }

  private findNode(entityId: number): SmartDeviceNode | null {
    return this.withKey((key) => {
      const raw = this.profiles.devicesFor(key);
      const tree = parseDevices(raw ?? []);
      return findDeviceById(tree, entityId);
    }) ?? null;
  }

  private findDevice(entityId: number): EngineDevice | null {
    const node = this.findNode(entityId);
    if (!node) return null;
    return {
      entityId: node.entityId,
      alias: node.alias ?? undefined,
      isGroup: node.isGroup,
      // Live state wins over the stale stored flag; absent → treated as off.
      isOn: this.deviceStates.get(node.entityId) ?? (node as { isOn?: boolean }).isOn ?? false,
      isMissing: node.isMissing,
    };
  }

  private findGroupSwitches(groupName: string): EngineDevice[] | null {
    return this.withKey((key) => {
      const tree = parseDevices(this.profiles.devicesFor(key) ?? []);
      const group = tree.find((d) => d.isGroup && d.alias === groupName);
      if (!group) return null;
      const switches: SmartDeviceNode[] = [];
      const walk = (nodes: readonly SmartDeviceNode[]): void => {
        for (const n of nodes) {
          if (!n.isGroup && n.kind === "SmartSwitch") switches.push(n);
          walk(n.children);
        }
      };
      walk(group.children);
      return switches.map((sw): EngineDevice => ({
        entityId: sw.entityId,
        alias: sw.alias ?? undefined,
        isGroup: false,
        isOn: this.deviceStates.get(sw.entityId) ?? (sw as { isOn?: boolean }).isOn ?? false,
        isMissing: sw.isMissing,
      }));
    });
  }

  private flatEntityIds(): number[] {
    return (
      this.withKey((key) => {
        const tree = parseDevices(this.profiles.devicesFor(key) ?? []);
        const ids: number[] = [];
        const walk = (nodes: readonly SmartDeviceNode[]): void => {
          for (const n of nodes) {
            if (!n.isGroup) ids.push(n.entityId);
            walk(n.children);
          }
        };
        walk(tree);
        return ids;
      }) ?? []
    );
  }

  private customTimers(): Array<CustomTimerInput & { id: string }> {
    return this.withKey((key) => {
      const raw = this.profiles.field(key, "CustomTimers");
      if (!Array.isArray(raw)) return [];
      return raw.map((t) => {
        const r = (t ?? {}) as Record<string, unknown>;
        const num = (v: unknown, dflt: number): number => (typeof v === "number" ? v : dflt);
        return {
          id: typeof r.Id === "string" ? r.Id : randomUUID(),
          name: typeof r.Name === "string" ? r.Name : "Timer",
          command: typeof r.Command === "string" ? r.Command : "timer",
          endUtc: num(r.EndTimeUtcMs, Date.now()),
          createdNotified: r.CreatedNotified === true,
          notified60: r.Notified60 === true,
          notified30: r.Notified30 === true,
          notified10: r.Notified10 === true,
          notified3: r.Notified3 === true,
        };
      });
    }) ?? [];
  }

  private saveCustomTimers(key: string, timers: Array<CustomTimerInput & { id: string }>): void {
    this.profiles.setField(
      key,
      "CustomTimers",
      timers.map((t) => ({
        Id: t.id,
        Name: t.name,
        Command: t.command,
        EndTimeUtcMs: t.endUtc,
        CreatedNotified: t.createdNotified,
        Notified60: t.notified60,
        Notified30: t.notified30,
        Notified10: t.notified10,
        Notified3: t.notified3,
      })),
    );
  }
}
