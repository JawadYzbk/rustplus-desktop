/**
 * EvaluateDeviceAutomationAsync — line-faithful port of Views/MainWindow/Devices/
 * MainWindow.DeviceAutomation.cs (L65-129) driven from the team poll tick
 * (MainWindow.Team.Core.cs TeamTimer_Tick L322-341 calls it after every team load).
 *
 * Semantics preserved:
 *  - re-entrancy gate: a still-running evaluation makes the next tick return immediately
 *    (C# SemaphoreSlim(1,1) + WaitAsync(0));
 *  - only enabled rules; PlayerProximity skips anchor lookup for *Offline modes (no anchor
 *    needed); GameTime with an unparseable server time skips the rule;
 *  - targets must exist, not be missing, and be Smart Switches;
 *  - conflicting decisions on one target log "Conflicting rules ... no action taken";
 *  - already-at-state targets are skipped; global busy (toggle/logic engine) blocks actions;
 *  - toggle timeout 8 s, then an 800 ms gap after every attempt (success or failure);
 *  - LastAppliedState is tracked in-memory only ([JsonIgnore]).
 */
import {
  isProximityMatch,
  tryGetTimeMatch,
  type PlayerSnapshot,
} from "./device-automation-evaluator.js";
import {
  findDeviceById,
  type SmartDeviceNode,
} from "../devices/device-data.js";
import type { DeviceAutomationRule } from "../devices/server-profile.js";

export const AUTOMATION_TOGGLE_TIMEOUT_MS = 8_000;
export const AUTOMATION_TOGGLE_GAP_MS = 800;

/** Extracts team members from a raw getTeamInfo payload (proto camelCase keys). */
export function extractTeamMembers(teamInfo: unknown): PlayerSnapshot[] {
  const members = (teamInfo as { members?: unknown })?.members;
  if (!Array.isArray(members)) return [];
  return members.flatMap((m): PlayerSnapshot[] => {
    const r = (m ?? {}) as Record<string, unknown>;
    const rawSteamId = r.steamId ?? r.steamId64;
    if (rawSteamId === undefined) return [];
    // IsDead lives on the UI model fed by map markers — marker ingestion is the stage-6 seam.
    return [
      {
        steamId: String(rawSteamId),
        isOnline: r.online === true && r.isDead !== true,
        x: typeof r.x === "number" ? r.x : null,
        y: typeof r.y === "number" ? r.y : null,
      },
    ];
  });
}

export interface AutomationHost {
  isActive(): boolean;
  rules(): DeviceAutomationRule[];
  /** Live tree lookup incl. paired-location fields. */
  findDevice(entityId: number): SmartDeviceNode | null;
  /** Current live switch state (null when unknown). */
  getIsOn(entityId: number): boolean | null;
  players(): PlayerSnapshot[];
  /** _vm.ServerTime ("HH:mm") or null when unknown. */
  serverTime(): string | null;
  /** _globalToggleBusy || _logicEngineRunningAction parity. */
  busy(): boolean;
  toggle(entityId: number, on: boolean): Promise<void>;
  log(message: string): void;
}

interface Decision {
  rule: DeviceAutomationRule;
  target: SmartDeviceNode;
  state: boolean;
}

export class DeviceAutomationService {
  private running = false;
  private readonly lastApplied = new Map<string, boolean>(); // rule.id → state

  constructor(private readonly host: AutomationHost) {}

  async evaluate(): Promise<void> {
    if (!this.host.isActive() || this.host.rules().length === 0) return;
    if (this.running) return; // WaitAsync(0) skip parity
    this.running = true;
    try {
      await this.evaluateInner();
    } finally {
      this.running = false;
    }
  }

  private async evaluateInner(): Promise<void> {
    const players = this.host.players();
    const decisions: Decision[] = [];

    for (const rule of this.host.rules()) {
      if (!rule.isEnabled) continue;

      let matched: boolean;
      if (rule.conditionType === "PlayerProximity") {
        let x = 0;
        let y = 0;
        if (!rule.playerMatchMode.endsWith("Offline")) {
          const location = this.host.findDevice(rule.locationEntityId);
          if (!location || location.pairedX === null || location.pairedY === null) continue;
          x = location.pairedX;
          y = location.pairedY;
        }
        matched = isProximityMatch(
          {
            playerMatchMode: rule.playerMatchMode,
            specificPlayerSteamId: rule.specificPlayerSteamId,
            distanceMeters: rule.distanceMeters,
          },
          x,
          y,
          players,
        );
      } else if (rule.conditionType === "GameTime") {
        const now = this.host.serverTime();
        if (now === null) continue;
        const tm = tryGetTimeMatch({ startTime: rule.startTime, endTime: rule.endTime }, now);
        if (!tm.parsed) continue;
        matched = tm.matched;
      } else {
        continue;
      }

      const target = this.host.findDevice(rule.targetEntityId);
      if (
        !target ||
        target.isMissing ||
        target.kind?.replace(/ /g, "").toLowerCase() !== "smartswitch"
      ) {
        continue;
      }
      decisions.push({ rule, target, state: matched ? rule.matchedState : rule.unmatchedState });
    }

    // Group by target entity id — conflicts take no action.
    const byTarget = new Map<number, Decision[]>();
    for (const d of decisions) {
      const list = byTarget.get(d.target.entityId);
      if (list) list.push(d);
      else byTarget.set(d.target.entityId, [d]);
    }

    for (const [entityId, group] of byTarget) {
      const states = [...new Set(group.map((d) => d.state))];
      if (states.length !== 1) {
        this.host.log(`[DeviceAutomation] Conflicting rules for device #${entityId}; no action taken.`);
        continue;
      }
      const decision = group[0]!;
      const currentIsOn = this.host.getIsOn(decision.target.entityId);
      if (currentIsOn === decision.state) continue;
      if (this.host.busy()) continue;

      const displayName =
        decision.target.alias && decision.target.alias.length > 0
          ? decision.target.alias
          : String(decision.target.entityId);
      try {
        await Promise.race([
          this.host.toggle(decision.target.entityId, decision.state),
          new Promise<never>((_, reject) =>
            setTimeout(() => reject(new Error("toggle timed out")), AUTOMATION_TOGGLE_TIMEOUT_MS),
          ),
        ]);
        this.lastApplied.set(decision.rule.id, decision.state);
        this.host.log(
          `[DeviceAutomation] '${decision.rule.name}' set ${displayName} ${decision.state ? "ON" : "OFF"}.`,
        );
        // Task.Delay(800) sits inside the C# try — the gap follows successful toggles only.
        await new Promise((r) => setTimeout(r, AUTOMATION_TOGGLE_GAP_MS));
      } catch (err: unknown) {
        this.host.log(
          `[DeviceAutomation] '${decision.rule.name}' failed: ${err instanceof Error ? err.message : String(err)}`,
        );
      }
    }
  }

  lastAppliedState(ruleId: string): boolean | null {
    return this.lastApplied.get(ruleId) ?? null;
  }
}
