/**
 * OilRigTriggerRegistry — port of Services/OilRigTriggerRegistry.cs. Which Smart Alarms are wired
 * to an oil rig, across every saved profile; must answer BEFORE any connection exists (push
 * backlogs land seconds after launch). The rules are the truth — deliberately not a persisted copy.
 *
 * C# statics become an instance with injected localized labels ({small, large}).
 */
import { isOilRigTimer, type LogicRule } from "./logic-rule.js";

export interface RigProfileInput {
  devices?: RigDevice[] | null;
  logicRules?: LogicRule[] | null;
}

export interface RigDevice {
  entityId: number;
  inGameAlarmTitle?: string | null;
  children?: RigDevice[] | null;
}

/** UiBadgeSmallOil / UiBadgeLargeOil equivalents — injected so language stays a bootstrap concern. */
export interface OilRigLabels {
  small: string;
  large: string;
}

/** Texts every unrenamed alarm sends — matching one would swallow real raid alerts (must never). */
const DEFAULT_ALARM_TEXTS = new Set(
  [
    "alarm",
    "smart alarm",
    "alarm activated!",
    "your base is under attack!",
    "your base is under attack",
    "base attacked",
    "triggered",
  ].map((t) => t),
);

function isDistinctive(text: string | null | undefined): boolean {
  if (text === undefined || text === null || text.trim().length === 0) return false;
  const trimmed = text.trim();
  // Two characters cannot identify anything, and a stray keystroke in the field must not start
  // matching real raid alerts.
  if (trimmed.length < 3) return false;
  return !DEFAULT_ALARM_TEXTS.has(trimmed.toLowerCase());
}

/** Steps including those nested in conditional branches (EnumerateSteps parity). */
export function enumerateSteps(rule: LogicRule): LogicRule["steps"] {
  const out: LogicRule["steps"] = [];
  for (const step of rule.steps) {
    out.push(step);
    out.push(...step.conditionalSteps);
  }
  return out;
}

function findDevice(devices: RigDevice[] | null | undefined, entityId: number): RigDevice | null {
  if (!devices) return null;
  for (const device of devices) {
    if (device.entityId === entityId) return device;
    const child = findDevice(device.children, entityId);
    if (child !== null) return child;
  }
  return null;
}

const labelFor = (target: string, labels: OilRigLabels): string =>
  target === "SmallOilRig" ? labels.small : labels.large;

export class OilRigTriggerRegistry {
  private byId = new Map<number, string>();
  private byName = new Map<string, string>();

  constructor(private readonly labels: OilRigLabels) {}

  /** Re-reads every profile's rules. Cheap; call it freely. */
  rebuild(profiles: Iterable<RigProfileInput | null | undefined>): void {
    const byId = new Map<number, string>();
    const byName = new Map<string, string>();

    for (const profile of profiles) {
      if (!profile?.logicRules) continue;
      for (const rule of profile.logicRules) {
        if (!rule.isEnabled) continue;
        if (rule.triggerType !== "SmartAlarm" || rule.triggerEntityId === 0) continue;

        const step = enumerateSteps(rule).find((s) => s.stepType === "StartTimer" && isOilRigTimer(s));
        if (!step) continue;

        const label = labelFor(step.timerTarget, this.labels);
        byId.set(rule.triggerEntityId, label);

        // The alarm's own text first — it follows renames in-game. The typed hint stays as a second
        // key because someone who filled it in expects it to work, but it goes stale on rename.
        // Matching only the hint is what let a renamed alarm ring as a raid again after restart.
        const device = findDevice(profile.devices, rule.triggerEntityId);
        if (isDistinctive(device?.inGameAlarmTitle)) {
          byName.set(device!.inGameAlarmTitle!.trim(), label);
        }
        if (isDistinctive(step.alarmTextHint)) {
          byName.set(step.alarmTextHint.trim(), label);
        }
      }
    }

    this.byId = byId;
    this.byName = byName;
  }

  /**
   * The rig label if this alarm is a rig trigger, else null.
   * Entity ID first (unambiguous); name fallback exists because FCM pushes frequently carry no ID.
   */
  lookup(entityId: number | null | undefined, ...names: Array<string | null | undefined>): string | null {
    if (entityId !== null && entityId !== undefined && this.byId.has(entityId)) {
      return this.byId.get(entityId)!;
    }
    for (const name of names) {
      if (name === null || name === undefined || name.trim().length === 0) continue;
      const hit = this.byName.get(name.trim());
      if (hit !== undefined) return hit;
    }
    return null;
  }

  /** All trigger devices on one profile keyed to the UNTRANSLATED target ("SmallOilRig"/"LargeOilRig") —
   *  what gets synced; a translated badge means nothing to the worker and changes with language. */
  targetsForProfile(profile: RigProfileInput | null | undefined): Map<number, string> {
    const map = new Map<number, string>();
    if (!profile?.logicRules) return map;
    for (const rule of profile.logicRules) {
      if (!rule.isEnabled) continue;
      if (rule.triggerType !== "SmartAlarm" || rule.triggerEntityId === 0) continue;
      const step = enumerateSteps(rule).find((s) => s.stepType === "StartTimer" && isOilRigTimer(s));
      if (!step) continue;
      map.set(rule.triggerEntityId, step.timerTarget);
    }
    return map;
  }

  /** ForProfile parity — same devices, but with translated badge labels (for the device list UI). */
  badgesForProfile(profile: RigProfileInput | null | undefined): Map<number, string> {
    const out = new Map<number, string>();
    for (const [id, target] of this.targetsForProfile(profile)) {
      out.set(id, labelFor(target, this.labels));
    }
    return out;
  }

  /**
   * Records the text an alarm actually sends, once identified by entity ID. Proven correct by
   * definition, so it overwrites whatever was typed — a typo repairs itself on first identified fire.
   * Title only, never the message (parity comment preserved in learnAlarmText below).
   * Returns true when something changed and the profiles want saving.
   */
  learnAlarmText(
    profiles: Iterable<RigProfileInput | null | undefined>,
    entityId: number,
    title: string | null | undefined,
  ): boolean {
    const learned = isDistinctive(title) ? title!.trim() : null;
    if (learned === null) return false;

    let changed = false;
    for (const profile of profiles) {
      if (!profile?.logicRules) continue;
      for (const rule of profile.logicRules) {
        if (rule.triggerType !== "SmartAlarm" || rule.triggerEntityId !== entityId) continue;
        for (const step of enumerateSteps(rule)) {
          if (step.stepType !== "StartTimer" || !isOilRigTimer(step)) continue;
          if (step.alarmTextHint === learned) continue;
          step.alarmTextHint = learned;
          changed = true;
        }
      }
    }
    return changed;
  }
}
