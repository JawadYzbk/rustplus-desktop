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
