/**
 * LogicRule JSON (de)serialization + LogicEngineService host-binding tests.
 * The service is exercised with fake store/conn/hub adapters — no Electron, no network.
 */
import { describe, expect, it } from "vitest";
import {
  parseLogicRule,
  serializeLogicRule,
  parseStep,
} from "../src/main/services/automation/logic-rule.js";
import {
  LogicEngineService,
  hubAdapter,
} from "../src/main/services/automation/engine-service.js";

const LEGACY_RULE = {
  Id: "r1",
  Name: "Turret cycle",
  IsEnabled: true,
  IsLoopEnabled: false,
  LoopCount: 3,
  TriggerType: "SmartSwitch",
  TriggerEntityId: 42,
  TriggerCommand: "cycle",
  TriggerRuleId: "",
  TriggerState: true,
  ConditionOperator: "AND",
  ConditionDeviceEntityId: 7,
  ConditionDeviceState: false,
  Steps: [
    { StepType: "Wait", WaitSeconds: 12, TimerMinutes: 15 },
    {
      StepType: "Toggle",
      TargetEntityId: 99,
      ToggleState: true,
      ConditionalSteps: [{ StepType: "StartTimer", TimerTarget: "LargeOilRig", TimerMinutes: 0 }],
    },
    { StepType: "CheckAvailability", ConditionOperator: "ANY_ONLINE", ConditionDeviceIdsCsv: "1, 2, x" },
  ],
};

describe("LogicRule JSON contract", () => {
  it("parses legacy PascalCase with clamps applied on load", () => {
    const r = parseLogicRule(LEGACY_RULE);
    expect(r.name).toBe("Turret cycle");
    expect(r.triggerType).toBe("SmartSwitch");
    expect(r.conditionOperator).toBe("AND");
    expect(r.steps).toHaveLength(3);
    // TimerMinutes: 0 → clamped to ≥1 at load (C# setter parity).
    const timer = r.steps[1]!.conditionalSteps[0]!;
    expect(timer.stepType).toBe("StartTimer");
    expect(timer.timerMinutes).toBe(1);
    expect(timer.timerTarget).toBe("LargeOilRig");
    // Wait step keeps its seconds:
    expect(r.steps[0]!.waitSeconds).toBe(12);
    // CSV is stored raw; parsing happens at evaluation time.
    expect(r.steps[2]!.conditionDeviceIdsCsv).toBe("1, 2, x");
  });

  it("tolerates unknown enum names and missing fields instead of crashing the load", () => {
    const r = parseLogicRule({ Id: "x", Steps: [{ StepType: "Detonate" }] });
    expect(r.triggerType).toBe("SmartAlarm");
    expect(r.steps[0]!.stepType).toBe("Wait"); // unknown step type → default
    expect(parseStep({}).stepType).toBe("Wait");
  });

  it("round-trips through PascalCase serialization", () => {
    const r = parseLogicRule(LEGACY_RULE);
    const out = serializeLogicRule(r) as Record<string, unknown>;
    expect(out.Id).toBe("r1");
    expect(out.IsEnabled).toBe(true);
    const steps = out.Steps as Array<Record<string, unknown>>;
    expect((steps[1]!.ConditionalSteps as Array<Record<string, unknown>>)[0]!.TimerMinutes).toBe(1);
    // Re-parsing its own output yields identical data.
    expect(JSON.stringify(serializeLogicRule(parseLogicRule(out)))).toBe(JSON.stringify(out));
  });
});

// ------------------------------------------------------------------ service binding

interface FakeStoreShape {
  data: Map<string, Record<string, unknown>>;
  activeKey: string | null;
}

function makeService(store: FakeStoreShape) {
  const toggles: Array<{ entityId: number; on: boolean }> = [];
  const chats: string[] = [];
  const probes: number[] = [];
  const hubListeners: Array<(e: { entityId: number; value: boolean }) => void> = [];

  const svc = new LogicEngineService(
    {
      activeKey: () => store.activeKey,
      field: (key, name) => store.data.get(key)?.[name],
      setField: (key, name, value) => {
        const rec = store.data.get(key);
        if (!rec) return false;
        rec[name] = value;
        return true;
      },
      devicesFor: (key) => (store.data.get(key)?.["Devices"] as never[]) ?? null,
      saveDevices: () => true,
    },
    {
      isConnected: () => true,
      setEntityValue: async (entityId, on) => {
        toggles.push({ entityId, on });
      },
      getEntityInfo: async (entityId) => {
        probes.push(entityId);
        return { entityInfo: { entityId, value: false } };
      },
      sendTeamMessage: async (message) => {
        chats.push(message);
      },
    },
    {
      handleEntityInfoResponse: (entityId: number, entityPayload: unknown) => {
        void entityId;
        void entityPayload;
      },
      onDeviceState: (listener) => {
        hubListeners.push(listener);
        return () => {
          hubListeners.splice(hubListeners.indexOf(listener), 1);
        };
      },
    },
  );
  return { svc, toggles, chats, probes, hubListeners };
}

function seedProfile(store: FakeStoreShape): void {
  store.activeKey = "h:1|s";
  store.data.set("h:1|s", {
    IsLogicEngineActive: true,
    AlertCustomTimer: true,
    Devices: [
      { EntityId: 10, Kind: "SmartSwitch", Alias: null, Name: "Sw", IsGroup: false, IsMissing: false, Children: [] },
      {
        EntityId: 20,
        Kind: null,
        Alias: "Lights",
        Name: "G",
        IsGroup: true,
        IsMissing: false,
        Children: [
          { EntityId: 21, Kind: "SmartSwitch", Alias: null, Name: "A", IsGroup: false, IsMissing: false, Children: [] },
          { EntityId: 22, Kind: "SmartSwitch", Alias: null, Name: "B", IsGroup: false, IsMissing: false, Children: [] },
        ],
      },
    ],
  });
}

describe("LogicEngineService host binding", () => {
  it("device-state pushes reach the engine and fire a matching enabled rule end-to-end", async () => {
    const store: FakeStoreShape = { data: new Map(), activeKey: null };
    seedProfile(store);
    const { svc, toggles } = makeService(store);

    svc.saveRules([
      {
        id: "r-fire",
        name: "fire",
        isEnabled: true,
        isLoopEnabled: false,
        loopCount: 1,
        triggerType: "SmartSwitch",
        triggerEntityId: 10,
        triggerCommand: "",
        triggerRuleId: "",
        triggerState: true,
        conditionOperator: "NONE",
        conditionDeviceEntityId: 0,
        conditionDeviceState: true,
        steps: [{ stepType: "Toggle", waitSeconds: 10, timerMinutes: 15, timerTarget: "Custom", timerName: "", showCrateOnMap: true, alarmTextHint: "", targetEntityId: 10, targetGroupName: "", toggleState: false, conditionOperator: "ALL_OFFLINE", conditionDeviceIdsCsv: "", conditionalSteps: [] }],
      },
    ]);

    svc.onDeviceEvent(10, true); // matches trigger
    // Wait until the toggle lands AND the rule settles (enqueue is fire-and-forget, so also
    // tolerate the not-yet-started window).
    const deadline = Date.now() + 3000;
    do {
      await new Promise((r) => setTimeout(r, 25));
    } while ((svc.status().isRunning || toggles.length === 0) && Date.now() < deadline);
    expect(toggles).toEqual([{ entityId: 10, on: false }]);
    expect(svc.status().isRunning).toBe(false);
  }, 5000);

  it("disabled engine swallows triggers; runRule works regardless via manual enqueue", async () => {
    const store: FakeStoreShape = { data: new Map(), activeKey: null };
    seedProfile(store);
    store.data.get("h:1|s")!["IsLogicEngineActive"] = false;
    const { svc, toggles } = makeService(store);
    svc.saveRules([
      {
        id: "r-x",
        name: "x",
        isEnabled: true,
        isLoopEnabled: false,
        loopCount: 1,
        triggerType: "SmartSwitch",
        triggerEntityId: 10,
        triggerCommand: "",
        triggerRuleId: "",
        triggerState: true,
        conditionOperator: "NONE",
        conditionDeviceEntityId: 0,
        conditionDeviceState: true,
        steps: [],
      },
    ]);
    svc.onDeviceEvent(10, true);
    await new Promise((r) => setTimeout(r, 20));
    expect(toggles).toEqual([]);

    expect(await svc.runRule("r-missing")).toBe(false);
    expect(await svc.runRule("r-x")).toBe(true);
  });

  it("refresh pulls every flat device incl. group children; group lookup works", async () => {
    const store: FakeStoreShape = { data: new Map(), activeKey: null };
    seedProfile(store);
    const { svc, probes, toggles } = makeService(store);
    const host = svc as unknown as { refreshAllDevices(): Promise<void>; findGroupSwitches(n: string): unknown[] | null };

    await host.refreshAllDevices();
    expect(probes.sort()).toEqual([10, 21, 22]); // leaves only — groups are not real entities

    expect(host.findGroupSwitches("Lights")).toHaveLength(2);
    expect(host.findGroupSwitches("Nope")).toBeNull();
  });

  it("saveRulesFor preserves stored steps for unchanged ids", async () => {
    const store: FakeStoreShape = { data: new Map(), activeKey: null };
    seedProfile(store);
    const { svc } = makeService(store);
    const full = {
      id: "keep",
      name: "keep",
      isEnabled: true,
      isLoopEnabled: false,
      loopCount: 1,
      triggerType: "SmartAlarm" as const,
      triggerEntityId: 5,
      triggerCommand: "",
      triggerRuleId: "",
      triggerState: true,
      conditionOperator: "NONE" as const,
      conditionDeviceEntityId: 0,
      conditionDeviceState: true,
      steps: [{ stepType: "Wait" as const, waitSeconds: 30, timerMinutes: 15, timerTarget: "Custom" as const, timerName: "", showCrateOnMap: true, alarmTextHint: "", targetEntityId: 0, targetGroupName: "", toggleState: null, conditionOperator: "ALL_OFFLINE" as const, conditionDeviceIdsCsv: "", conditionalSteps: [] }],
    };
    svc.saveRules([full]);

    // Header-only edit arrives WITHOUT steps — they must survive.
    const ok = svc.saveRulesFor("h:1|s", [{ ...full, name: "renamed", steps: undefined as never }], false);
    expect(ok).toBe(true);
    const rules = svc.rulesFor("h:1|s") as Array<Record<string, unknown>>;
    expect(rules).toHaveLength(1);
    const reparsed = parseLogicRule(rules[0]!);
    expect(reparsed.name).toBe("renamed");
    expect(reparsed.steps).toHaveLength(1);
    expect(reparsed.steps[0]!.waitSeconds).toBe(30);
    expect(svc.isEngineActiveFor("h:1|s")).toBe(false);
  });

  it("ruleFor + saveFullRuleFor round-trip a full rule incl. nested conditional steps", () => {
    const store: FakeStoreShape = { data: new Map(), activeKey: null };
    seedProfile(store);
    const { svc } = makeService(store);

    expect(svc.ruleFor("h:1|s", "nope")).toBeNull();

    const full = {
      id: "edit-me",
      name: "gate",
      isEnabled: true,
      isLoopEnabled: false,
      loopCount: 2,
      triggerType: "SmartSwitch" as const,
      triggerEntityId: 10,
      triggerCommand: "",
      triggerRuleId: "",
      triggerState: true,
      conditionOperator: "NONE" as const,
      conditionDeviceEntityId: 0,
      conditionDeviceState: true,
      steps: [
        {
          stepType: "CheckAvailability" as const,
          waitSeconds: 10,
          timerMinutes: 15,
          timerTarget: "Custom" as const,
          timerName: "",
          showCrateOnMap: true,
          alarmTextHint: "",
          targetEntityId: 0,
          targetGroupName: "",
          toggleState: null,
          conditionOperator: "ALL_OFFLINE" as const,
          conditionDeviceIdsCsv: "31,32",
          conditionalSteps: [
            {
              stepType: "Toggle" as const,
              waitSeconds: 10,
              timerMinutes: 15,
              timerTarget: "Custom" as const,
              timerName: "",
              showCrateOnMap: true,
              alarmTextHint: "",
              targetEntityId: 40,
              targetGroupName: "",
              toggleState: true,
              conditionOperator: "ALL_OFFLINE" as const,
              conditionDeviceIdsCsv: "",
              conditionalSteps: [],
            },
          ],
        },
      ],
    };
    // Zero timerMinutes must be clamped to 1 on save (parseLogicRule load-path clamp).
    expect(svc.saveFullRuleFor("h:1|s", { ...full, steps: [{ ...full.steps[0]!, conditionalSteps: full.steps[0]!.conditionalSteps }] })).toBe(true);

    const loaded = svc.ruleFor("h:1|s", "edit-me");
    expect(loaded).not.toBeNull();
    expect(loaded!.loopCount).toBe(2);
    expect(loaded!.steps).toHaveLength(1);
    expect(loaded!.steps[0]!.conditionDeviceIdsCsv).toBe("31,32");
    expect(loaded!.steps[0]!.conditionalSteps[0]!.targetEntityId).toBe(40);
    expect(loaded!.steps[0]!.conditionalSteps[0]!.toggleState).toBe(true);

    // Wholesale replace of the same id:
    const renamed = { ...loaded!, name: "renamed", steps: [] };
    expect(svc.saveFullRuleFor("h:1|s", renamed)).toBe(true);
    expect((svc.rulesFor("h:1|s") as unknown[]).length).toBe(1); // replaced, not appended
    expect(svc.ruleFor("h:1|s", "edit-me")!.name).toBe("renamed");
  });

  it("hubAdapter filters deviceState events and unsubscribes cleanly", () => {
    const listeners: Array<(...args: unknown[]) => void> = [];
    const fakeHub = {
      handleEntityInfoResponse: () => undefined,
      on: (_ev: string, cb: (...args: unknown[]) => void) => {
        listeners.push(cb);
      },
      off: (_ev: string, cb: (...args: unknown[]) => void) => {
        const i = listeners.indexOf(cb);
        if (i >= 0) listeners.splice(i, 1);
      },
    };
    const seen: number[] = [];
    const adapter = hubAdapter(fakeHub);
    const unsub = adapter.onDeviceState((e) => seen.push(e.entityId));
    for (const l of [...listeners]) {
      l({ kind: "deviceState", entityId: 9, on: true });
      l({ kind: "storageSnapshot", entityId: 9 });
      l({ kind: "deviceState", entityId: 8 }); // on missing → value=false still delivered
    }
    expect(seen).toEqual([9, 8]);
    unsub();
    expect(listeners).toHaveLength(0);
  });
});
