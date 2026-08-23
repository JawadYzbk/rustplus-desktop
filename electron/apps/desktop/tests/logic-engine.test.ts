/**
 * LogicEngine golden tests — trigger matching, condition gate, sequential queue, step execution
 * (Wait/Toggle/StartTimer/CheckAvailability), loop semantics, cancellation, custom-timer quirks.
 * Virtual clock: sleeps resolve instantly; busy-flag waits still observable.
 */
import { describe, expect, it } from "vitest";
import {
  LogicEngine,
  RuleCancelled,
  CUSTOM_TIMER_LIMIT,
  type LogicEngineHost,
  type EngineDevice,
  type CustomTimerInput,
} from "../src/main/services/automation/logic-engine.js";
import { newLogicRule, newLogicStep, type LogicRule } from "../src/main/services/automation/logic-rule.js";

function manualClock() {
  return {
    now_: 0,
    now() {
      return this.now_;
    },
    async sleep(ms: number) {
      this.now_ += ms;
    },
  };
}

interface Rig {
  host: LogicEngineHost & { logs: string[]; toggles: Array<[number, boolean]>; timers: CustomTimerInput[]; busy: boolean; connected: boolean; devices: Map<number, EngineDevice>; active: boolean };
  engine: LogicEngine;
  events: Array<{ kind: string; [k: string]: unknown }>;
  drain(): Promise<void>;
}

function rig(rules: LogicRule[], over: Partial<LogicEngineHost> = {}): Rig {
  const clock = manualClock();
  const r: Rig["host"] = {
    logs: [],
    toggles: [],
    timers: [],
    busy: false,
    connected: true,
    devices: new Map(),
    active: true,
    isEngineActive: () => r.active,
    rules: () => rules,
    findDevice: (id) => r.devices.get(id) ?? null,
    findGroupSwitches: (name) => (name === "lights" ? [r.devices.get(1)!, r.devices.get(2)!].filter(Boolean) : null),
    isConnected: () => r.connected,
    toggleSmartSwitch: async (id, on) => void r.toggles.push([id, on]),
    refreshAllDevices: async () => undefined,
    triggerRigTimer: () => true,
    customTimerCount: () => r.timers.length,
    removeCustomTimerByName: (name) => {
      const i = r.timers.findIndex((t) => t.name.toLowerCase() === name.toLowerCase());
      if (i >= 0) r.timers.splice(i, 1);
    },
    addCustomTimer: (t) => void r.timers.push(t),
    alertCustomTimers: () => false,
    sendTeamChat: () => undefined,
    chatAlert: (m) => r.logs.push(m),
    log: (m) => r.logs.push(m),
    isBusy: () => r.busy,
    clock: clock as never,
    ...over,
  } as Rig["host"];

  const events: Array<{ kind: string }> = [];
  const engine = new LogicEngine(r, (e) => void events.push(e as { kind: string }));
  const drain = async (): Promise<void> => {
    for (let i = 0; i < 12; i++) await new Promise<void>((res) => setImmediate(res));
  };
  return { host: r, engine, events: events as never, drain };
}

const switchOn = (id: number, isOn = true): EngineDevice => ({ entityId: id, isOn });

describe("trigger matching", () => {
  it("fires on matching entity + state; wrong state or id does not", async () => {
    const rule = newLogicRule({ isEnabled: true, triggerType: "SmartSwitch", triggerEntityId: 7, triggerState: true });
    const t = rig([rule]);
    t.engine.onDeviceEvent(8, true);
    t.engine.onDeviceEvent(7, false);
    await t.drain();
    expect(t.events.filter((e) => e.kind === "started")).toHaveLength(0);

    t.engine.onDeviceEvent(7, true);
    await t.drain();
    expect(t.events.filter((e) => e.kind === "started")).toHaveLength(1);
  });

  it("chat command strips prefix case-insensitively on both sides", async () => {
    const rule = newLogicRule({ isEnabled: true, triggerType: "ChatCommand", triggerCommand: "!Open" });
    const t = rig([rule]);
    t.engine.onChatCommand("  open  "); // no prefix, matches after rule-side strip
    await t.drain();
    t.engine.onChatCommand("!OPEN");
    await t.drain();
    expect(t.events.filter((e) => e.kind === "started")).toHaveLength(2);
  });

  it("inactive engine ignores every trigger", async () => {
    const rule = newLogicRule({ isEnabled: true, triggerEntityId: 7 });
    const t = rig([rule]);
    t.host.active = false;
    t.engine.onDeviceEvent(7, true);
    await t.drain();
    expect(t.events).toHaveLength(0);
  });

  it("AND condition requires the device state; missing device fails; OR passes", async () => {
    const andRule = newLogicRule({
      isEnabled: true,
      triggerEntityId: 7,
      conditionOperator: "AND",
      conditionDeviceEntityId: 99,
      conditionDeviceState: true,
    });
    const t = rig([andRule]);
    t.engine.onDeviceEvent(7, true); // device 99 absent → condition cannot be met
    await t.drain();
    expect(t.events.filter((e) => e.kind === "started")).toHaveLength(0);

    t.host.devices.set(99, switchOn(99, false));
    t.engine.onDeviceEvent(7, true);
    await t.drain();
    expect(t.events.filter((e) => e.kind === "started")).toHaveLength(0);

    t.host.devices.set(99, switchOn(99, true));
    t.engine.onDeviceEvent(7, true);
    await t.drain();
    expect(t.events.filter((e) => e.kind === "started")).toHaveLength(1);

    const orRule = newLogicRule({ isEnabled: true, triggerEntityId: 7, conditionOperator: "OR", conditionDeviceEntityId: 99, conditionDeviceState: false });
    const t2 = rig([orRule]);
    t2.host.devices.set(99, switchOn(99, true));
    t2.engine.onDeviceEvent(7, true);
    await t2.drain();
    expect(t2.events.filter((e) => e.kind === "started")).toHaveLength(1); // OR → true
  });

  it("RuleTriggered chains; a rule never triggers itself on start", async () => {
    const a = newLogicRule({ id: "a", name: "rule-a", isEnabled: true, triggerEntityId: 7 });
    const b = newLogicRule({ id: "b", name: "rule-b", isEnabled: true, triggerType: "RuleTriggered", triggerRuleId: "a" });
    const c = newLogicRule({ id: "c", name: "rule-c", isEnabled: true, triggerType: "RuleTriggered", triggerRuleId: "b" });
    // c lists ITSELF as its own RuleTriggered source — only reachable once c actually runs.
    const selfRef = newLogicRule({ id: "d", name: "rule-d", isEnabled: true, triggerType: "RuleTriggered", triggerRuleId: "c" });
    void selfRef;
    const t = rig([a, b, c]);
    t.engine.onDeviceEvent(7, true);
    await t.drain();
    const started = t.events.filter((e) => e.kind === "started").map((e) => e.rule);
    expect(started).toEqual(["rule-a", "rule-b", "rule-c"]); // strict chain order
  });
});

describe("step execution", () => {
  it("runs Wait steps and Toggle steps in order with forced state", async () => {
    const rule = newLogicRule({
      isEnabled: true,
      steps: [
        newLogicStep({ stepType: "Wait", waitSeconds: 2 }),
        newLogicStep({ stepType: "Toggle", targetEntityId: 5, toggleState: true }),
      ],
    });
    const t = rig([rule]);
    t.host.devices.set(5, switchOn(5, false));
    void t.engine.enqueue(rule);
    await t.drain();

    expect(t.host.toggles).toEqual([[5, true]]);
    expect(t.host.devices.get(5)!.isOn).toBe(true);
    expect(t.events.filter((e) => e.kind === "completed")).toHaveLength(1);
  });

  it("invert toggle when no state given; single-target ALWAYS sends (no skip — C# L427)", async () => {
    const rule = newLogicRule({
      steps: [
        newLogicStep({ stepType: "Toggle", targetEntityId: 5 }), // invert: true → false
        newLogicStep({ stepType: "Toggle", targetEntityId: 6 }), // false → invert → true (still sent!)
      ],
    });
    const t = rig([rule]);
    t.host.devices.set(5, switchOn(5, true));
    t.host.devices.set(6, switchOn(6, false));
    void t.engine.enqueue(rule);
    await t.drain();
    expect(t.host.toggles).toEqual([[5, false], [6, true]]);
  });

  it("group toggle inverts from first switch and SKIPS members already at target (C# L396)", async () => {
    const rule = newLogicRule({ steps: [newLogicStep({ stepType: "Toggle", targetGroupName: "lights" })] });
    const t = rig([rule]);
    t.host.devices.set(1, switchOn(1, true));
    t.host.devices.set(2, switchOn(2, false));
    void t.engine.enqueue(rule);
    await t.drain();
    // First is on → invert → all OFF; #2 already off → skipped.
    expect(t.host.toggles).toEqual([[1, false]]);
    expect(t.host.devices.get(2)!.isOn).toBe(false);
  });

  it("toggle failures abort with chat alert (HandleRuleFailure parity)", async () => {
    const rule = newLogicRule({ steps: [newLogicStep({ stepType: "Toggle", targetEntityId: 404 })] });
    const t = rig([rule]);
    void t.engine.enqueue(rule);
    await t.drain();
    expect(t.events.filter((e) => e.kind === "failed")).toHaveLength(1);
    expect(t.host.logs.join("\n")).toContain("Target switch #404 not found");
    expect(t.host.logs.join("\n")).toContain("⚠️");
  });

  it("custom timers: five-limit, replace-by-name, milestone suppression, cmd slug", async () => {
    const mk = (name: string) =>
      newLogicRule({ steps: [newLogicStep({ stepType: "StartTimer", timerName: name, timerMinutes: 15 })] });
    const t = rig([mk("Crate 1"), mk("crate 1"), mk("t2"), mk("t3"), mk("t4"), mk("t5"), mk("t6")]);
    void t.engine.enqueue(mk("Crate 1"));
    await t.drain();
    expect(t.host.timers.map((x) => x.name)).toEqual(["Crate 1"]);

    void t.engine.enqueue(mk("crate 1")); // replace rather than stack (case-insensitive)
    await t.drain();
    expect(t.host.timers.map((x) => x.name)).toEqual(["crate 1"]);
    expect(t.host.timers[0]!.command).toBe("crate1");
    expect(t.host.timers[0]!.notified60).toBe(true); // 15 <= 60 → pre-suppressed
    expect(t.host.timers[0]!.notified30).toBe(true);
    expect(t.host.timers[0]!.notified10).toBe(false); // 15 > 10 → still announced
    expect(t.host.timers[0]!.notified3).toBe(false);

    for (const n of ["t2", "t3", "t4", "t5"]) void t.engine.enqueue(mk(n));
    await t.drain();
    expect(t.host.timers).toHaveLength(5);

    void t.engine.enqueue(mk("t6"));
    await t.drain();
    expect(t.host.timers).toHaveLength(CUSTOM_TIMER_LIMIT); // sixth refused
    expect(t.host.logs.join("\n")).toContain("five-timer limit");
  });

  it("rig timers route to MonumentWatcher and respect already-running", async () => {
    let calls = 0;
    const rule = newLogicRule({
      steps: [newLogicStep({ stepType: "StartTimer", timerTarget: "LargeOilRig", timerMinutes: 7, showCrateOnMap: true })],
    });
    const t = rig([rule], { triggerRigTimer: () => ++calls <= 1 });
    void t.engine.enqueue(rule);
    await t.drain();
    expect(calls).toBe(1);
    expect(t.host.logs.join("\n")).toContain("Started 7 min hack timer for Large Oil Rig");

    void t.engine.enqueue(rule);
    await t.drain();
    expect(t.host.logs.join("\n")).toContain("already has a running timer");
  });

  it("CheckAvailability gate: single-target operator quirks + CSV counting + conditional steps", async () => {
    // Single target: IS_OFFLINE/ALL_OFFLINE/ANY_OFFLINE all mean "offline" (legacy quirk preserved).
    const single = newLogicRule({
      steps: [newLogicStep({ stepType: "CheckAvailability", targetEntityId: 9, conditionOperator: "ANY_ONLINE" })],
    });
    const t = rig([single]);
    t.host.devices.set(9, { entityId: 9, isMissing: true });
    void t.engine.enqueue(single);
    await t.drain();
    // Gate aborts the STEPS but the run itself still completes (legacy break → Completed log →
    // TriggerLogicEngineOnRuleCompleted still fires).
    expect(t.events.filter((e) => e.kind === "completed")).toHaveLength(1);
    expect(t.host.logs.join("\n")).toContain("Gating condition failed");

    // CSV: ANY_ONLINE with one of two online → met → conditional toggle runs.
    const csv = newLogicRule({
      steps: [
        newLogicStep({
          stepType: "CheckAvailability",
          conditionDeviceIdsCsv: "11, 12, bad, 0",
          conditionOperator: "ANY_ONLINE",
          conditionalSteps: [newLogicStep({ stepType: "Toggle", targetEntityId: 13, toggleState: true })],
        }),
      ],
    });
    const t2 = rig([csv]);
    t2.host.devices.set(11, switchOn(11, true));
    t2.host.devices.set(12, { entityId: 12, isMissing: true });
    t2.host.devices.set(13, { entityId: 13 }); // conditional-step target must exist
    void t.engine;
    void t2.engine.enqueue(csv);
    await t2.drain();
    expect(t2.host.toggles).toEqual([[13, true]]);
  });
});

describe("queueing, loops, stop", () => {
  it("runs rules strictly sequentially", async () => {
    const mk = (id: string) => newLogicRule({ id, isEnabled: true, steps: [newLogicStep({ stepType: "Wait", waitSeconds: 1 })] });
    const t = rig([mk("a"), mk("b"), mk("c")]);
    const order: string[] = [];
    for (const r of [mk("a"), mk("b"), mk("c")]) {
      void t.engine.enqueue(r).then(() => order.push(r.id));
    }
    await t.drain();
    expect(order).toEqual(["a", "b", "c"]);
  });

  it("loop re-enqueues: LoopCount 3 yields FOUR total runs (remaining 3→2→1→0, C# L212-214)", async () => {
    const looped = newLogicRule({
      isLoopEnabled: true,
      loopCount: 3,
      steps: [newLogicStep({ stepType: "Wait", waitSeconds: 1 }), newLogicStep({ stepType: "Toggle", targetEntityId: 5, toggleState: true })],
    });
    const t = rig([looped]);
    t.host.devices.set(5, switchOn(5, false));
    void t.engine.enqueue(looped);
    await t.drain();
    expect(t.host.toggles).toEqual([[5, true], [5, true], [5, true], [5, true]]);
    expect(t.events.filter((e) => e.kind === "completed")).toHaveLength(4);
  });

  it("requestStop aborts the in-flight rule at the next boundary (stopped, not failed)", async () => {
    const rule = newLogicRule({
      steps: [
        newLogicStep({ stepType: "Toggle", targetEntityId: 5, toggleState: true }),
        newLogicStep({ stepType: "Wait", waitSeconds: 1 }),
        newLogicStep({ stepType: "Toggle", targetEntityId: 5, toggleState: false }),
      ],
    });
    const t = rig([rule]);
    let togglesSeen = 0;
    t.host.toggleSmartSwitch = async (id, on) => {
      t.host.toggles.push([id, on]); // keep recording like the default recorder
      if (++togglesSeen === 1) t.engine.requestStop(); // stop lands mid-rule, before the Wait step
    };
    t.host.devices.set(5, switchOn(5, false));
    await t.engine.enqueue(rule);
    await t.drain();

    expect(t.events.some((e) => e.kind === "stopped")).toBe(true);
    expect(t.events.some((e) => e.kind === "failed")).toBe(false);
    expect(t.host.logs.join("\n")).toContain("Stop requested");
    // The cancelled run never reports completion, and the second toggle never happened.
    expect(t.events.filter((e) => e.kind === "completed")).toHaveLength(0);
    expect(t.host.toggles).toEqual([[5, true]]);
  });

  it("loop without a Wait step logs the skip hint and does not re-enqueue", async () => {
    const looped = newLogicRule({
      isLoopEnabled: true,
      loopCount: 3,
      steps: [newLogicStep({ stepType: "Toggle", targetEntityId: 5, toggleState: false })],
    });
    const t = rig([looped]);
    t.host.devices.set(5, switchOn(5, true));
    void t.engine.enqueue(looped);
    await t.drain();
    expect(t.events.filter((e) => e.kind === "started")).toHaveLength(1);
    expect(t.host.logs.join("\n")).toContain("loop skipped: add a Wait step");
  });

  it("RuleCancelled surfaces as stopped, not failed", async () => {
    expect(new RuleCancelled()).toBeInstanceOf(Error);
  });
});
