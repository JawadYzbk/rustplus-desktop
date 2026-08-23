/**
 * LogicEngine — headless port of MainWindow.LogicEngine.cs execution semantics:
 *  - strictly sequential rule runs (semaphore parity), runtime surface mirrors
 *    LogicEngineRuntimeService (pending queue, current rule/step, requestStop);
 *  - triggers: device events (SmartAlarm/SmartSwitch + state), chat commands (prefix-stripped,
 *    case-insensitive), RuleCompleted / RuleTriggered chaining incl. the self-trigger guard;
 *  - trigger conditions: NONE→true, AND→device state match, OR→true (the event just happened),
 *    missing condition device→false;
 *  - steps: Wait, Toggle (single or group; force-or-invert; per-call 5 s timeout; 800 ms gap),
 *    StartTimer (rig timers handed to MonumentWatcher; custom timers capped at five, replaced
 *    by name, milestone-suppressed like the manual dialog), CheckAvailability gate
 *    (single-target operator quirks preserved: ANY_/ALL_OFFLINE all mean "offline");
 *  - loops only re-enqueue when a Wait step > 0 exists; LoopCount 0 = infinite.
 */
import type { Clock } from "../rustplus/timing.js";
import { realClock } from "../rustplus/timing.js";
import type { LogicRule, LogicStep } from "./logic-rule.js";
import { oilRigName } from "./logic-rule.js";

export interface EngineDevice {
  entityId: number;
  alias?: string;
  isGroup?: boolean;
  isOn?: boolean | null;
  isMissing?: boolean;
}

export interface CustomTimerInput {
  name: string;
  command: string;
  endUtc: number;
  createdNotified: boolean;
  notified60: boolean;
  notified30: boolean;
  notified10: boolean;
  notified3: boolean;
}

export interface LogicEngineHost {
  /** profile.IsLogicEngineActive */
  isEngineActive(): boolean;
  rules(): LogicRule[];
  /** Recursive FindDeviceById parity. */
  findDevice(entityId: number): EngineDevice | null;
  /** Group lookup by alias (d.IsGroup && d.Alias == name) with recursive switch flattening done by the host. */
  findGroupSwitches(groupName: string): EngineDevice[] | null;
  isConnected(): boolean;
  /** _rust.ToggleSmartSwitchAsync parity — caller wraps with the 5 s per-call timeout. */
  toggleSmartSwitch(entityId: number, on: boolean): Promise<void>;
  refreshAllDevices(): Promise<void>;
  /** MonumentWatcher.TriggerExternal parity; false = timer already running (left alone). */
  triggerRigTimer(rigName: string, seconds: number, showCrate: boolean): boolean;
  customTimerCount(): number;
  removeCustomTimerByName(name: string): void;
  addCustomTimer(timer: CustomTimerInput): void;
  alertCustomTimers(): boolean; // profile.AlertCustomTimer
  sendTeamChat(message: string): void;
  chatAlert(message: string): void;
  log(message: string): void;
  /** Strict mutex flags (_globalToggleBusy / _refreshAllBusy == 1). */
  isBusy(): boolean;
  /** profile.ChatCommandPrefix ?? "!" */
  chatCommandPrefix?(): string;
  clock?: Clock;
}

export const TOGGLE_CALL_TIMEOUT_MS = 5_000;
export const TOGGLE_GAP_MS = 800;
export const BUSY_POLL_MS = 500;
export const CUSTOM_TIMER_LIMIT = 5;

/** OperationCanceledException analogue. */
export class RuleCancelled extends Error {
  constructor() {
    super("rule cancelled");
  }
}

interface Cancellation {
  cancelled: boolean;
}

type PushEvent =
  | { kind: "enqueued"; rule: string }
  | { kind: "started"; rule: string }
  | { kind: "step"; rule: string; number: number; stepType: LogicStep["stepType"] }
  | { kind: "completed"; rule: string }
  | { kind: "stopped"; rule: string }
  | { kind: "failed"; rule: string; error: string };

export class LogicEngine {
  private readonly clock: Clock;
  private queue: Promise<void> = Promise.resolve(); // semaphore(1,1) parity
  private currentCancellation: Cancellation | null = null;
  private readonly pendingRules: string[] = [];

  // Runtime service surface:
  isRunning = false;
  currentRuleName: string | null = null;
  currentStepNumber = 0;
  currentStepType: LogicStep["stepType"] | null = null;

  constructor(
    private readonly host: LogicEngineHost,
    private readonly events: (e: PushEvent) => void = () => undefined,
  ) {
    this.clock = host.clock ?? realClock;
  }

  get pending(): readonly string[] {
    return this.pendingRules;
  }

  /** IsLogicEngineActiveAndWaiting parity: profile selected && IsLogicEngineActive. */
  private get activeAndWaiting(): boolean {
    return this.host.isEngineActive();
  }

  requestStop(): void {
    if (this.currentCancellation) this.currentCancellation.cancelled = true;
    this.host.log("[LogicEngine] Stop requested. Current rule will abort after the current operation.");
  }

  // ------------------------------------------------------------------ triggers

  onDeviceEvent(entityId: number, isOn: boolean): void {
    if (!this.activeAndWaiting) return;
    for (const rule of this.host.rules()) {
      if (!rule.isEnabled) continue;
      if (rule.triggerType !== "SmartAlarm" && rule.triggerType !== "SmartSwitch") continue;
      if (rule.triggerEntityId === entityId && rule.triggerState === isOn && this.evaluateTriggerCondition(rule)) {
        void this.enqueue(rule);
      }
    }
  }

  onChatCommand(cmdText: string): void {
    if (!this.activeAndWaiting) return;

    const prefix = this.host.chatCommandPrefix?.() ?? "!";
    let cmd = cmdText.trim().toLowerCase();
    // Prefix stripping happens for BOTH the incoming text and the rule command.
    if (prefix.length > 0 && cmd.startsWith(prefix)) cmd = cmd.slice(prefix.length).trim();

    for (const rule of this.host.rules()) {
      if (!rule.isEnabled || rule.triggerType !== "ChatCommand") continue;
      let ruleCmd = (rule.triggerCommand ?? "").trim().toLowerCase();
      if (prefix.length > 0 && ruleCmd.startsWith(prefix)) ruleCmd = ruleCmd.slice(prefix.length).trim();
      if (ruleCmd === cmd && this.evaluateTriggerCondition(rule)) void this.enqueue(rule);
    }
  }

  private onRuleCompleted(completedRuleId: string): void {
    if (!this.activeAndWaiting) return;
    for (const rule of this.host.rules()) {
      if (rule.isEnabled && rule.triggerType === "RuleCompleted" && rule.triggerRuleId === completedRuleId) {
        if (this.evaluateTriggerCondition(rule)) void this.enqueue(rule);
      }
    }
  }

  private onRuleTriggered(triggeredRule: LogicRule): void {
    if (!this.activeAndWaiting) return;
    for (const rule of this.host.rules()) {
      if (!(rule.isEnabled && rule.triggerType === "RuleTriggered" && rule.triggerRuleId === triggeredRule.id)) {
        continue;
      }
      if (rule.id === triggeredRule.id) {
        this.host.log(`[LogicEngine] Rule '${rule.name}' cannot trigger itself on start; use Loop after completion.`);
        continue;
      }
      if (this.evaluateTriggerCondition(rule)) void this.enqueue(rule);
    }
  }

  /** EvaluateTriggerCondition parity (incl. OR always true once the event occurred). */
  evaluateTriggerCondition(rule: LogicRule): boolean {
    if (rule.conditionOperator === "NONE" || !rule.conditionOperator) return true;
    const condDev = this.host.findDevice(rule.conditionDeviceEntityId);
    if (!condDev) return false; // condition device deleted/missing → cannot be met
    const condState = condDev.isOn ?? false;
    const targetState = rule.conditionDeviceState;
    if (rule.conditionOperator === "AND") return condState === targetState;
    if (rule.conditionOperator === "OR") return true; // the main trigger event already happened
    return true;
  }

  // ------------------------------------------------------------------ execution

  enqueue(rule: LogicRule, remainingLoopCount?: number): Promise<void> {
    this.events({ kind: "enqueued", rule: rule.name });
    this.host.log(`[LogicEngine] Rule '${rule.name}' triggered. Enqueuing...`);
    const run = this.queue.catch(() => undefined).then(() => this.runRuleLifecycle(rule, remainingLoopCount));
    this.queue = run;
    return run;
  }

  private async sleepCancellable(ms: number, cancellation: Cancellation | null): Promise<void> {
    await this.clock.sleep(ms);
    if (cancellation?.cancelled) throw new RuleCancelled();
  }

  private async runRuleLifecycle(rule: LogicRule, remainingLoopCount?: number): Promise<void> {
    this.pendingRules.push(rule.name);
    try {
      // Chat Master gate (checked AFTER dequeuing, before execution starts).
      const blockedChat = false; // host hook reserved; ChatCommand master-block not wired headless yet
      void blockedChat;

      const cancellation: Cancellation = { cancelled: false };
      this.currentCancellation = cancellation;
      this.isRunning = true;
      this.currentRuleName = rule.name;
      this.currentStepNumber = 0;
      this.currentStepType = null;
      this.events({ kind: "started", rule: rule.name });
      this.onRuleTriggered(rule);

      try {
        await this.runSteps(rule, cancellation);
        if (rule.isLoopEnabled) {
          const hasWait = rule.steps.some((s) => s.stepType === "Wait" && s.waitSeconds > 0);
          if (hasWait) {
            const remaining = remainingLoopCount ?? (rule.loopCount === 0 ? -1 : rule.loopCount);
            if (remaining !== 0) {
              void this.enqueue(rule, remaining < 0 ? -1 : remaining - 1);
            }
          } else {
            this.host.log(
              `[LogicEngine] Rule '${rule.name}' loop skipped: add a Wait step greater than 0 seconds.`,
            );
          }
        }
        this.events({ kind: "completed", rule: rule.name });
        this.onRuleCompleted(rule.id);
      } finally {
        this.isRunning = false;
        this.currentRuleName = null;
        this.currentStepNumber = 0;
        this.currentStepType = null;
        this.currentCancellation = null;
      }
    } catch (err) {
      if (err instanceof RuleCancelled) {
        this.events({ kind: "stopped", rule: rule.name });
        this.host.log(`[LogicEngine] Rule '${rule.name}' execution was stopped.`);
      } else {
        const message = err instanceof Error ? err.message : String(err);
        this.events({ kind: "failed", rule: rule.name, error: message });
        this.host.log(`[LogicEngine] Rule '${rule.name}' error: ${message}`);
        this.handleRuleFailure(rule, message);
      }
    } finally {
      const idx = this.pendingRules.indexOf(rule.name);
      if (idx >= 0) this.pendingRules.splice(idx, 1);
    }
  }

  private async waitWhileBusy(cancellation: Cancellation): Promise<void> {
    while (this.host.isBusy()) {
      await this.sleepCancellable(BUSY_POLL_MS, cancellation);
    }
  }

  private async runSteps(rule: LogicRule, cancellation: Cancellation): Promise<void> {
    this.host.log(`[LogicEngine] Starting execution of rule '${rule.name}'...`);
    let stepNum = 0;
    for (const step of rule.steps) {
      if (cancellation.cancelled) throw new RuleCancelled();
      stepNum++;
      this.currentStepNumber = stepNum;
      this.currentStepType = step.stepType;
      this.events({ kind: "step", rule: rule.name, number: stepNum, stepType: step.stepType });
      this.host.log(`[LogicEngine] Running step ${stepNum} (${step.stepType}) for rule '${rule.name}'...`);

      await this.waitWhileBusy(cancellation);

      if (step.stepType === "Wait") {
        await this.sleepCancellable(step.waitSeconds * 1000, cancellation);
      } else if (step.stepType === "Toggle") {
        await this.executeToggle(step, cancellation);
      } else if (step.stepType === "StartTimer") {
        await this.executeStartTimer(step);
      } else if (step.stepType === "CheckAvailability") {
        const conditionMet = await this.executeCheckAvailability(step, cancellation);
        if (!conditionMet) {
          this.host.log(`[LogicEngine] Gating condition failed for rule '${rule.name}'. Aborting rule execution.`);
          break;
        }
      }
    }
    this.host.log(`[LogicEngine] Completed execution of rule '${rule.name}'.`);
  }

  private async executeToggle(step: LogicStep, cancellation: Cancellation): Promise<void> {
    if (step.targetGroupName.length > 0) {
      const switches = this.host.findGroupSwitches(step.targetGroupName);
      if (!switches || switches.length === 0) {
        throw new Error(`Group '${step.targetGroupName}' not found or has no devices.`);
      }
      if (!this.host.isConnected()) throw new Error("Companion app not connected.");

      // Force the given state; otherwise invert based on the FIRST switch.
      const targetOn = step.toggleState ?? !(switches[0]!.isOn ?? false);
      for (const sw of switches) {
        if (cancellation.cancelled) throw new RuleCancelled();
        if (sw.isOn === targetOn) continue;
        await this.toggleWithTimeout(sw.entityId, targetOn, cancellation);
        sw.isOn = targetOn;
        await this.sleepCancellable(TOGGLE_GAP_MS, cancellation); // wait between toggle calls
      }
    } else {
      const dev = this.host.findDevice(step.targetEntityId);
      if (!dev) throw new Error(`Target switch #${step.targetEntityId} not found.`);
      if (dev.isMissing) throw new Error(`Target switch #${step.targetEntityId} is offline/missing.`);
      if (!this.host.isConnected()) throw new Error("Companion app not connected.");

      const targetOn = step.toggleState ?? !(dev.isOn ?? false);
      await this.toggleWithTimeout(dev.entityId, targetOn, cancellation);
      dev.isOn = targetOn;
      await this.sleepCancellable(TOGGLE_GAP_MS, cancellation);
    }
  }

  private async toggleWithTimeout(entityId: number, on: boolean, cancellation: Cancellation): Promise<void> {
    // Work registered FIRST so instant outcomes win microtask ties (established tie rule).
    const work = this.host.toggleSmartSwitch(entityId, on);
    const guard = this.clock.sleep(TOGGLE_CALL_TIMEOUT_MS).then(() => {
      if (!cancellation.cancelled) throw new Error(`toggle ${entityId} timed out`);
      throw new RuleCancelled();
    });
    await Promise.race([work, guard]);
  }

  private async executeStartTimer(step: LogicStep): Promise<void> {
    const minutes = Math.max(1, step.timerMinutes);
    const rig = oilRigName(step);

    if (rig !== null) {
      const started = this.host.triggerRigTimer(rig, minutes * 60, step.showCrateOnMap);
      this.host.log(
        started
          ? `[LogicEngine] Started ${minutes} min hack timer for ${rig}` +
              (step.showCrateOnMap ? " (crate shown on map)." : " (no map marker).")
          : `[LogicEngine] ${rig} already has a running timer — left it alone.`,
      );
      return;
    }

    // Plain timer. Same ceiling as the manual dialog, since they share the list.
    const name = step.timerName.trim().length === 0 ? "Timer" : step.timerName.trim();

    if (this.host.customTimerCount() >= CUSTOM_TIMER_LIMIT) {
      this.host.log(`[LogicEngine] Cannot start timer '${name}': the five-timer limit is reached.`);
      return;
    }
    // Replace rather than stack: a rule firing twice means the same countdown restarted.
    this.host.removeCustomTimerByName(name);

    let cmd = [...name.toLowerCase()].filter((ch) => /[a-z0-9]/.test(ch)).join("");
    if (cmd.length === 0) cmd = "timer";

    this.host.addCustomTimer({
      name,
      command: cmd,
      endUtc: Date.now() + minutes * 60_000,
      createdNotified: false,
      // Suppress milestones the timer starts below, exactly as the manual path does.
      notified60: minutes <= 60,
      notified30: minutes <= 30,
      notified10: minutes <= 10,
      notified3: minutes <= 3,
    });

    this.host.log(`[LogicEngine] Started ${minutes} min timer '${name}'.`);

    if (this.host.alertCustomTimers()) {
      this.host.sendTeamChat(`⏱️ Timer created: !${cmd} (${minutes} min)`);
    }
  }

  private async executeCheckAvailability(step: LogicStep, cancellation: Cancellation): Promise<boolean> {
    // Wait until manual refresh finishes if busy, then run a single refresh.
    while (this.host.isBusy()) {
      await this.sleepCancellable(BUSY_POLL_MS, cancellation);
    }

    this.host.log("[LogicEngine] Refreshing device availability states...");
    await this.host.refreshAllDevices();
    if (cancellation.cancelled) throw new RuleCancelled();

    let conditionMet = false;
    if (step.targetEntityId !== 0) {
      const dev = this.host.findDevice(step.targetEntityId);
      const isOffline = dev === null || dev.isMissing === true;
      const op = step.conditionOperator;
      if (op === "IS_OFFLINE" || op === "ALL_OFFLINE" || op === "ANY_OFFLINE") conditionMet = isOffline;
      else if (op === "IS_ONLINE" || op === "ALL_ONLINE" || op === "ANY_ONLINE") conditionMet = !isOffline;
    } else {
      const ids = step.conditionDeviceIdsCsv
        .split(",")
        .map((s) => s.trim())
        .filter((s) => s.length > 0)
        .map((s) => (/^\d+$/.test(s) ? Number(s) : 0))
        .filter((id) => id !== 0);
      if (ids.length === 0) return false; // no devices listed → condition fails outright

      let offlineCount = 0;
      let onlineCount = 0;
      for (const id of ids) {
        const dev = this.host.findDevice(id);
        if (!dev || dev.isMissing) offlineCount++;
        else onlineCount++;
      }
      switch (step.conditionOperator) {
        case "ALL_OFFLINE":
          conditionMet = offlineCount === ids.length;
          break;
        case "ANY_OFFLINE":
          conditionMet = offlineCount > 0;
          break;
        case "ALL_ONLINE":
          conditionMet = onlineCount === ids.length;
          break;
        case "ANY_ONLINE":
          conditionMet = onlineCount > 0;
          break;
      }
    }

    if (conditionMet) {
      this.host.log(
        `[LogicEngine] Availability condition '${step.conditionOperator}' met. Executing conditional steps...`,
      );
      for (const condStep of step.conditionalSteps) {
        if (cancellation.cancelled) throw new RuleCancelled();
        await this.waitWhileBusy(cancellation);
        if (condStep.stepType === "Wait") {
          await this.sleepCancellable(condStep.waitSeconds * 1000, cancellation);
        } else if (condStep.stepType === "Toggle") {
          await this.executeToggle(condStep, cancellation);
        } else if (condStep.stepType === "StartTimer") {
          await this.executeStartTimer(condStep);
        }
      }
    }
    return conditionMet;
  }

  private handleRuleFailure(rule: LogicRule, errorMsg: string): void {
    // Chat alert (Discord premium hook lands with the cloud stage).
    this.host.chatAlert(`⚠️ Rule '${rule.name}' failed: ${errorMsg}`);
  }
}
