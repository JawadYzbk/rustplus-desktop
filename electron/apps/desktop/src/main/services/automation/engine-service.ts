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
  type EngineDevice,
  type LogicEngineHost,
} from "./logic-engine.js";
import { parseDevices, serializeDevices, parseTimer, serializeTimer, type CustomTimer } from "../devices/server-profile.js";
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
  /** CheckCustomTimers first-tick purge flag (C# _timerStartupCleanupDone). */
  private timerStartupCleanupDone = false;

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
        this.saveCustomTimers(key, [
          ...this.customTimers(),
          {
            id: randomUUID(),
            name: t.name,
            command: t.command,
            endTimeUtcMs: t.endUtc,
            enableCountdownAudio: true,
            enableAlarmAudio: false,
            createdNotified: t.createdNotified,
            notified60: t.notified60,
            notified30: t.notified30,
            notified10: t.notified10,
            notified3: t.notified3,
            countdownAudioPlayed: false,
            alarmPlayed: false,
            snoozedUntilUtcMs: null,
            autoDeleteAtUtcMs: null,
          },
        ]);
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

  /** Live switch state for automation consumers (null when never seen). */
  deviceIsOn(entityId: number): boolean | null {
    return this.deviceStates.has(entityId) ? (this.deviceStates.get(entityId) ?? null) : null;
  }

  /** Optimistic local state after a successful main-process toggle. */
  setDeviceState(entityId: number, isOn: boolean): void {
    this.deviceStates.set(entityId, isOn);
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

  /** _logicEngineRunningAction parity — automation skips toggles while a rule acts. */
  isActionRunning(): boolean {
    return this.engine.isRunning;
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

  /** Full stored node for Device Automation anchors and target validation. */
  findNodeFor(entityId: number): SmartDeviceNode | null {
    return this.findNode(entityId);
  }

  /** Canonical CustomTimer records for the ACTIVE profile (parseTimer handles dual-casing,
   * ISO dates and the C# defaults — including EnableCountdownAudio defaulting to true). */
  private customTimers(): CustomTimer[] {
    return this.withKey((key) => {
      const raw = this.profiles.field(key, "CustomTimers");
      if (!Array.isArray(raw)) return [];
      return raw.map((t) => parseTimer((t ?? {}) as Record<string, unknown>));
    }) ?? [];
  }

  private saveCustomTimers(key: string, timers: CustomTimer[]): void {
    this.profiles.setField(
      key,
      "CustomTimers",
      timers.map(serializeTimer),
    );
  }

  /**
   * CheckCustomTimers tick (MainWindow.Map.Timers.cs L107-268) — call ~1×/s while connected.
   * Startup purge drops already-expired timers once; removal at ≤ -60 s; countdown flag at
   * ≤60 s; expiry alert "{name}: 00:00" gated on AlertCustomTimer within the -60..0 window;
   * milestone alerts at 60/30/10 min with the ≥59/29/9 guards; audio cues are the stage-12
   * seam (flags still flip exactly like the C#).
   */
  tickTimers(): Array<{ id: string; name: string; command: string; endTimeUtcMs: number }> {
    const key = this.profiles.activeKey();
    if (!key) return [];
    const now = Date.now();
    let timers = this.customTimers();

    // First tick: silently purge anything expired from a previous session (L118-127).
    if (!this.timerStartupCleanupDone) {
      this.timerStartupCleanupDone = true;
      const before = timers.length;
      timers = timers.filter((t) => t.endTimeUtcMs - now > 0);
      if (timers.length !== before) this.saveCustomTimers(key, timers);
      if (timers.length === 0) return [];
    }

    const alertCustom = this.profiles.field(key, "AlertCustomTimer") === true;
    let changed = false;
    const remaining: Array<{ id: string; name: string; command: string; endTimeUtcMs: number }> = [];

    timers = timers.filter((t) => {
      const remMs = t.endTimeUtcMs - now;
      if (remMs <= -60_000) {
        changed = true;
        return false; // L136-139: gone a minute past expiry → remove
      }
      const remSecs = Math.floor(remMs / 1000);
      const remMins = Math.floor(remSecs / 60);

      // Countdown boundary (L142-152): flag flips once regardless of the audio setting.
      if (remSecs <= 60 && remSecs > 0 && !t.countdownAudioPlayed) {
        t.countdownAudioPlayed = true;
        changed = true;
        // PlayTimerAudio(true) lands with the audio stage.
      } else if (remSecs <= 0 && !t.alarmPlayed) {
        // Expiry alert inside the -60..0 window only (L154-157).
        if (alertCustom && remSecs >= -60) {
          this.hostSendTeamChat(`${t.name}: 00:00`);
        }
        t.alarmPlayed = true;
        changed = true;
        // PlayTimerAudio(false) lands with the audio stage.
      }

      // Milestone chat alerts (L183-219): flag flips at ≤X minutes EVEN when the message is
      // suppressed (late load below the guard) — message only fires within a minute of the
      // boundary (>= X-1).
      if (alertCustom) {
        const milestone = (
          flag: boolean,
          limit: number,
          guard: number,
          text: string,
        ): boolean => {
          if (flag || remMins > limit) return flag;
          if (remMins >= guard) this.hostSendTeamChat(text);
          changed = true;
          return true;
        };
        t.notified60 = milestone(t.notified60, 60, 59, `${t.name}: 60:00`);
        t.notified30 = milestone(t.notified30, 30, 29, `${t.name}: 30:00`);
        t.notified10 = milestone(t.notified10, 10, 9, `${t.name}: 10:00`);
        t.notified3 = milestone(t.notified3, 3, 2, `${t.name}: 03:00`);
      }

      remaining.push({ id: t.id, name: t.name, command: t.command, endTimeUtcMs: t.endTimeUtcMs });
      return true;
    });

    if (changed) this.saveCustomTimers(key, timers);
    return remaining;
  }

  /** Team-chat send used by timer alerts (bypasses the chat-alert master block in legacy). */
  private hostSendTeamChat(text: string): void {
    this.conn.sendTeamMessage(text).catch(() => undefined);
  }

  /** Timer list for chat consumers ({name, command, endTimeUtcMs}). */
  timersForChat(): Array<{ name: string; command: string; endTimeUtcMs: number }> {
    return this.customTimers().map((t) => ({
      name: t.name,
      command: t.command,
      endTimeUtcMs: t.endTimeUtcMs,
    }));
  }

  /** Full timer records for ANY profile (timers panel). */
  timersFor(matchKey: string): CustomTimer[] {
    const raw = this.profiles.field(matchKey, "CustomTimers");
    if (!Array.isArray(raw)) return [];
    return raw.map((t) => parseTimer((t ?? {}) as Record<string, unknown>));
  }

  removeTimerFor(matchKey: string, id: string): boolean {
    const timers = this.timersFor(matchKey);
    const next = timers.filter((t) => t.id !== id);
    if (next.length === timers.length) return false;
    this.saveCustomTimers(matchKey, next);
    return true;
  }

  /** BtnAddTimer_Click parity — returns the exact legacy validation outcome. */
  tryAddTimerFor(
    matchKey: string,
    name: string,
    hours: number,
    minutes: number,
    seconds: number,
  ): { ok: true; id: string } | { ok: false; reason: "limit" | "letter" | "duration" } {
    // Five-timer limit (L36-41).
    if (this.timersFor(matchKey).length >= 5) return { ok: false, reason: "limit" };
    const trimmed = name.trim();
    if (trimmed.length === 0 || !/[a-zA-Z]/.test(trimmed[0]!)) return { ok: false, reason: "letter" };
    const h = Math.max(0, Math.trunc(hours) || 0);
    const m = Math.max(0, Math.trunc(minutes) || 0);
    const s = Math.max(0, Math.trunc(seconds) || 0);
    if (h === 0 && m === 0 && s === 0) return { ok: false, reason: "duration" };
    // Command = whitespace-stripped lowercase name preview (TxtTimerName_TextChanged).
    let cmd = trimmed.replace(/\s+/g, "").toLowerCase();
    if (cmd.length === 0) cmd = trimmed.toLowerCase();
    const totalSecs = h * 3600 + m * 60 + s;
    const totalMins = totalSecs / 60;
    const record: CustomTimer = {
      id: randomUUID(),
      name: trimmed,
      command: cmd,
      endTimeUtcMs: Date.now() + totalSecs * 1000,
      enableCountdownAudio: true,
      enableAlarmAudio: false,
      createdNotified: false,
      notified60: totalMins <= 60,
      notified30: totalMins <= 30,
      notified10: totalMins <= 10,
      notified3: totalMins <= 3,
      countdownAudioPlayed: false,
      alarmPlayed: false,
      snoozedUntilUtcMs: null,
      autoDeleteAtUtcMs: null,
    };
    this.saveCustomTimers(matchKey, [...this.timersFor(matchKey), record]);
    return { ok: true, id: record.id };
  }

  /** Adds a timer from the chat pipeline (id assigned here; C# defaults for the audio flags). */
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
        endTimeUtcMs: t.endTimeUtcMs,
        enableCountdownAudio: true, // CustomTimer.cs default
        enableAlarmAudio: false, // CustomTimer.cs default
        createdNotified: t.createdNotified,
        notified60: t.notified60,
        notified30: t.notified30,
        notified10: t.notified10,
        notified3: t.notified3,
        countdownAudioPlayed: false,
        alarmPlayed: false,
        snoozedUntilUtcMs: null,
        autoDeleteAtUtcMs: null,
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
}
