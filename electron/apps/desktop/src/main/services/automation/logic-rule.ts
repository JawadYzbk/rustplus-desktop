/**
 * LogicRule / LogicStep models — port of Models/LogicRule.cs (observable state only; WPF
 * INotifyPropertyChanged plumbing is irrelevant headless). Defaults match the C# initializers
 * exactly, including the clamps (TimerMinutes ≥ 1 — zero would fire its own expiry immediately;
 * LoopCount ≥ 0).
 */

export type LogicStepType = "Wait" | "Toggle" | "CheckAvailability" | "StartTimer";
export type LogicTriggerType = "SmartAlarm" | "SmartSwitch" | "ChatCommand" | "RuleTriggered" | "RuleCompleted";
export type TimerTarget = "Custom" | "SmallOilRig" | "LargeOilRig";
/** CheckAvailability operators (single target: IS_*) and conditional-branch operators (ALL_/ANY_*). */
export type AvailabilityOperator =
  | "IS_OFFLINE"
  | "IS_ONLINE"
  | "ALL_OFFLINE"
  | "ANY_OFFLINE"
  | "ALL_ONLINE"
  | "ANY_ONLINE";

export interface LogicStep {
  stepType: LogicStepType;
  // ---- StartTimer ----
  timerMinutes: number; // clamped ≥ 1 on write
  /** Custom, SmallOilRig or LargeOilRig. */
  timerTarget: TimerTarget;
  /** Only used when timerTarget is Custom. */
  timerName: string;
  showCrateOnMap: boolean;
  /** The text this alarm sends, as set in-game; learned from pushes identified by entity ID. */
  alarmTextHint: string;
  waitSeconds: number;
  targetEntityId: number;
  targetGroupName: string;
  /** null = invert, true = ON, false = OFF. */
  toggleState: boolean | null;
  conditionOperator: AvailabilityOperator;
  conditionDeviceIdsCsv: string;
  conditionalSteps: LogicStep[];
}

export interface LogicRule {
  id: string;
  name: string;
  isEnabled: boolean;
  isLoopEnabled: boolean;
  loopCount: number; // clamped ≥ 0 on write
  triggerType: LogicTriggerType;
  triggerEntityId: number;
  triggerCommand: string;
  triggerRuleId: string;
  triggerState: boolean;
  /** NONE, AND, OR — gate evaluated at fire time. */
  conditionOperator: "NONE" | "AND" | "OR";
  conditionDeviceEntityId: number;
  conditionDeviceState: boolean;
  steps: LogicStep[];
}

export function newLogicStep(over: Partial<LogicStep> = {}): LogicStep {
  return {
    stepType: "Wait",
    timerMinutes: 15,
    timerTarget: "Custom",
    timerName: "",
    showCrateOnMap: true,
    alarmTextHint: "",
    waitSeconds: 10,
    targetEntityId: 0,
    targetGroupName: "",
    toggleState: null,
    conditionOperator: "ALL_OFFLINE",
    conditionDeviceIdsCsv: "",
    conditionalSteps: [],
    ...over,
  };
}

export function newLogicRule(over: Partial<LogicRule> = {}): LogicRule {
  return {
    id: crypto.randomUUID(),
    name: "New Rule",
    isEnabled: false,
    isLoopEnabled: false,
    loopCount: 1,
    triggerType: "SmartAlarm",
    triggerEntityId: 0,
    triggerCommand: "rulecommand",
    triggerRuleId: "",
    triggerState: true,
    conditionOperator: "NONE",
    conditionDeviceEntityId: 0,
    conditionDeviceState: true,
    steps: [],
    ...over,
  };
}

/** IsOilRigTimer parity. */
export function isOilRigTimer(step: Pick<LogicStep, "timerTarget">): boolean {
  return step.timerTarget === "SmallOilRig" || step.timerTarget === "LargeOilRig";
}

/** OilRigName parity — the name MonumentWatcher keys its events on, null for a plain timer. */
export function oilRigName(step: Pick<LogicStep, "timerTarget">): string | null {
  switch (step.timerTarget) {
    case "SmallOilRig":
      return "Small Oil Rig";
    case "LargeOilRig":
      return "Large Oil Rig";
    default:
      return null;
  }
}

/** Clamp helpers preserving the C# setter guards. */
export const clampTimerMinutes = (v: number): number => Math.max(1, Math.trunc(v));
export const clampLoopCount = (v: number): number => Math.max(0, Math.trunc(v));

// ---------------------------------------------------------------- JSON (de)serialization

type Raw = Record<string, unknown>;
const pick = (r: Raw, ...keys: string[]): unknown => {
  for (const k of keys) if (k in r) return r[k];
  return undefined;
};
const strOf = (v: unknown, dflt = ""): string => (typeof v === "string" ? v : dflt);
const boolOf = (v: unknown, dflt = false): boolean => (typeof v === "boolean" ? v : dflt);
const intOf = (v: unknown, dflt: number): number =>
  typeof v === "number" && Number.isFinite(v) ? Math.trunc(v) : dflt;

export function parseStep(raw: Raw): LogicStep {  const stepType = strOf(pick(raw, "StepType", "stepType"), "Wait");
  const conditionalRaw = pick(raw, "ConditionalSteps", "conditionalSteps");
  const toggleStateRaw = pick(raw, "ToggleState", "toggleState");
  const timerTargetRaw = strOf(pick(raw, "TimerTarget", "timerTarget"), "Custom");
  const conditionOpRaw = strOf(pick(raw, "ConditionOperator", "conditionOperator"), "ALL_OFFLINE");
  return {
    // C# deserializes enums by NAME; unknown names throw — we fall back to the defaults instead
    // of crashing the whole profile load, matching the legacy app's tolerance for hand-edited files.
    stepType: (["Wait", "Toggle", "CheckAvailability", "StartTimer"] as const).includes(stepType as LogicStepType)
      ? (stepType as LogicStepType)
      : "Wait",
    timerMinutes: clampTimerMinutes(intOf(pick(raw, "TimerMinutes", "timerMinutes"), 15)),
    timerTarget: (["Custom", "SmallOilRig", "LargeOilRig"] as const).includes(timerTargetRaw as TimerTarget)
      ? (timerTargetRaw as TimerTarget)
      : "Custom",
    timerName: strOf(pick(raw, "TimerName", "timerName")),
    showCrateOnMap: boolOf(pick(raw, "ShowCrateOnMap", "showCrateOnMap"), true),
    alarmTextHint: strOf(pick(raw, "AlarmTextHint", "alarmTextHint")),
    waitSeconds: Math.max(0, intOf(pick(raw, "WaitSeconds", "waitSeconds"), 10)),
    targetEntityId: intOf(pick(raw, "TargetEntityId", "targetEntityId"), 0),
    targetGroupName: strOf(pick(raw, "TargetGroupName", "targetGroupName")),
    // JSON has no null-tolerant bool? here: absent → null (invert semantics preserved).
    toggleState: typeof toggleStateRaw === "boolean" ? toggleStateRaw : null,
    conditionOperator: (
      ["IS_OFFLINE", "IS_ONLINE", "ALL_OFFLINE", "ANY_OFFLINE", "ALL_ONLINE", "ANY_ONLINE"] as const
    ).includes(conditionOpRaw as AvailabilityOperator)
      ? (conditionOpRaw as AvailabilityOperator)
      : "ALL_OFFLINE",
    conditionDeviceIdsCsv: strOf(pick(raw, "ConditionDeviceIdsCsv", "conditionDeviceIdsCsv")),
    conditionalSteps: Array.isArray(conditionalRaw) ? conditionalRaw.map((s) => parseStep((s ?? {}) as Raw)) : [],
  };
}

export function serializeStep(s: LogicStep): Raw {
  return {
    StepType: s.stepType,
    TimerMinutes: s.timerMinutes,
    TimerTarget: s.timerTarget,
    TimerName: s.timerName,
    ShowCrateOnMap: s.showCrateOnMap,
    AlarmTextHint: s.alarmTextHint,
    WaitSeconds: s.waitSeconds,
    TargetEntityId: s.targetEntityId,
    TargetGroupName: s.targetGroupName,
    ToggleState: s.toggleState,
    ConditionOperator: s.conditionOperator,
    ConditionDeviceIdsCsv: s.conditionDeviceIdsCsv,
    ConditionalSteps: s.conditionalSteps.map(serializeStep),
  };
}

/** Parses a stored LogicRule (PascalCase from the C# app; camelCase tolerated). */
export function parseLogicRule(raw: Raw): LogicRule {
  const stepsRaw = pick(raw, "Steps", "steps");
  const triggerTypeRaw = strOf(pick(raw, "TriggerType", "triggerType"), "SmartAlarm");
  const condOpRaw = strOf(pick(raw, "ConditionOperator", "conditionOperator"), "NONE");
  return {
    id: strOf(pick(raw, "Id", "id")) || crypto.randomUUID(),
    name: strOf(pick(raw, "Name", "name"), "New Rule"),
    isEnabled: boolOf(pick(raw, "IsEnabled", "isEnabled")),
    isLoopEnabled: boolOf(pick(raw, "IsLoopEnabled", "isLoopEnabled")),
    loopCount: clampLoopCount(intOf(pick(raw, "LoopCount", "loopCount"), 1)),
    triggerType: (["SmartAlarm", "SmartSwitch", "ChatCommand", "RuleTriggered", "RuleCompleted"] as const).includes(
      triggerTypeRaw as LogicTriggerType,
    )
      ? (triggerTypeRaw as LogicTriggerType)
      : "SmartAlarm",
    triggerEntityId: intOf(pick(raw, "TriggerEntityId", "triggerEntityId"), 0),
    triggerCommand: strOf(pick(raw, "TriggerCommand", "triggerCommand"), "rulecommand"),
    triggerRuleId: strOf(pick(raw, "TriggerRuleId", "triggerRuleId")),
    triggerState: boolOf(pick(raw, "TriggerState", "triggerState"), true),
    conditionOperator: (["NONE", "AND", "OR"] as const).includes(condOpRaw as LogicRule["conditionOperator"])
      ? (condOpRaw as LogicRule["conditionOperator"])
      : "NONE",
    conditionDeviceEntityId: intOf(pick(raw, "ConditionDeviceEntityId", "conditionDeviceEntityId"), 0),
    conditionDeviceState: boolOf(pick(raw, "ConditionDeviceState", "conditionDeviceState"), true),
    steps: Array.isArray(stepsRaw) ? stepsRaw.map((s) => parseStep((s ?? {}) as Raw)) : [],
  };
}

export function serializeLogicRule(r: LogicRule): Raw {
  return {
    Id: r.id,
    Name: r.name,
    IsEnabled: r.isEnabled,
    IsLoopEnabled: r.isLoopEnabled,
    LoopCount: r.loopCount,
    TriggerType: r.triggerType,
    TriggerEntityId: r.triggerEntityId,
    TriggerCommand: r.triggerCommand,
    TriggerRuleId: r.triggerRuleId,
    TriggerState: r.triggerState,
    ConditionOperator: r.conditionOperator,
    ConditionDeviceEntityId: r.conditionDeviceEntityId,
    ConditionDeviceState: r.conditionDeviceState,
    Steps: r.steps.map(serializeStep),
  };
}
