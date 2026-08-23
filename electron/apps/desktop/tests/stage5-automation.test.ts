/**
 * Stage-5 automation cores golden tests — DeviceAutomationEvaluator (incl. the C# DEBUG Verify()
 * asserts), LogicRule model clamps, OilRigTriggerRegistry behaviors from its source comments,
 * AlertTemplateService override persistence + format fallback chain.
 */
import { describe, expect, it } from "vitest";
import { mkdtempSync, readFileSync, writeFileSync, rmSync } from "node:fs";
import { tmpdir } from "node:os";
import { join } from "node:path";
import {
  isProximityMatch,
  isTimeMatch,
  tryGetTimeMatch,
  clampDistanceMeters,
  type PlayerSnapshot,
} from "../src/main/services/automation/device-automation-evaluator.js";
import {
  newLogicRule,
  newLogicStep,
  clampTimerMinutes,
  clampLoopCount,
  oilRigName,
} from "../src/main/services/automation/logic-rule.js";
import {
  OilRigTriggerRegistry,
  type RigProfileInput,
} from "../src/main/services/automation/oil-rig-trigger-registry.js";
import {
  AlertTemplateService,
  dotNetFormat,
} from "../src/main/services/automation/alert-template-service.js";

const p = (steamId: string, isOnline: boolean, x: number | null = null, y: number | null = null): PlayerSnapshot => ({
  steamId,
  isOnline,
  x,
  y,
});

describe("DeviceAutomationEvaluator", () => {
  // The C# static ctor's Verify() asserts, transcribed:
  it("Verify(): proximity + overnight window basics", () => {
    const rule = {
      playerMatchMode: "AnyOnline",
      specificPlayerSteamId: "",
      distanceMeters: 250,
      startTime: "",
      endTime: "",
    };
    expect(isProximityMatch(rule, 1000, 1000, [p("1", true, 1100, 1000)])).toBe(true);
    expect(isProximityMatch(rule, 1000, 1000, [p("1", false, 1000, 1000)])).toBe(false);

    const allOffline = { ...rule, playerMatchMode: "AllOffline" };
    expect(isProximityMatch(allOffline, 0, 0, [p("1", false)])).toBe(true);

    const timed = { ...allOffline, startTime: "20:00", endTime: "08:00" };
    expect(isTimeMatch(timed, "23:30")).toBe(true);
    expect(isTimeMatch(timed, "07:59")).toBe(true);
    expect(isTimeMatch(timed, "12:00")).toBe(false);
    expect(tryGetTimeMatch(timed, "–").parsed).toBe(false);
  });

  it("Specific modes filter by SteamId; SpecificOffline requires exactly one match", () => {
    const rule = {
      playerMatchMode: "SpecificOffline",
      specificPlayerSteamId: "76561198000000001",
      distanceMeters: 250,
      startTime: "",
      endTime: "",
    };
    expect(isProximityMatch(rule, 0, 0, [p("76561198000000001", false)])).toBe(true);
    expect(isProximityMatch(rule, 0, 0, [p("76561198000000002", false)])).toBe(false);
    // Two matches for the same id → selected.Count == 1 fails
    expect(isProximityMatch(rule, 0, 0, [p("76561198000000001", false), p("76561198000000001", false)])).toBe(false);
  });

  it("AllOnline needs every selected player online AND near", () => {
    const rule = {
      playerMatchMode: "AllOnline",
      specificPlayerSteamId: "",
      distanceMeters: 250,
      startTime: "",
      endTime: "",
    };
    expect(isProximityMatch(rule, 0, 0, [])).toBe(false); // empty selection → false
    expect(
      isProximityMatch(rule, 0, 0, [p("a", true, 200, 0), p("b", true, 0, 240)]),
    ).toBe(true);
    expect(
      isProximityMatch(rule, 0, 0, [p("a", true, 200, 0), p("b", true, 300, 0)]),
    ).toBe(false); // b out of range
    expect(
      isProximityMatch(rule, 0, 0, [p("a", true, 200, 0), p("b", false, 0, 0)]),
    ).toBe(false); // offline never near
  });

  it("time window: equal start/end means always; wrap uses OR", () => {
    const base = { startTime: "", endTime: "" };
    expect(isTimeMatch({ ...base, startTime: "09:00", endTime: "09:00" }, "03:00")).toBe(true);
    expect(tryGetTimeMatch({ ...base, startTime: "09:00", endTime: "17:00" }, "17:00")).toEqual({
      parsed: true,
      matched: false,
    }); // end exclusive
    expect(isTimeMatch({ ...base, startTime: "00:00", endTime: "24:00" }, "23:59")).toBe(true);
    // 25:00 parses via TimeSpan → wraps to 60 min mod 1440
    expect(isTimeMatch(base, "25:00")).toBe(false); // empty start unparseable → parsed=false path
  });
});

describe("LogicRule model", () => {
  it("defaults match the C# initializers", () => {
    const r = newLogicRule({ id: "x" });
    expect(r.name).toBe("New Rule");
    expect(r.triggerType).toBe("SmartAlarm");
    expect(r.triggerCommand).toBe("rulecommand");
    expect(r.loopCount).toBe(1);
    const s = newLogicStep();
    expect(s.stepType).toBe("Wait");
    expect(s.timerMinutes).toBe(15);
    expect(s.timerTarget).toBe("Custom");
    expect(s.showCrateOnMap).toBe(true);
    expect(s.conditionOperator).toBe("ALL_OFFLINE");
  });

  it("clamps and derived names", () => {
    expect(clampTimerMinutes(0)).toBe(1); // zero would fire its own expiry immediately
    expect(clampLoopCount(-3)).toBe(0);
    expect(oilRigName(newLogicStep({ timerTarget: "SmallOilRig" }))).toBe("Small Oil Rig");
    expect(oilRigName(newLogicStep({ timerTarget: "LargeOilRig" }))).toBe("Large Oil Rig");
    expect(oilRigName(newLogicStep({ timerTarget: "Custom" }))).toBeNull();
  });

  it("distance clamp parity Math.Max(1, value)", () => {
    expect(clampDistanceMeters(250)).toBe(250);
    expect(clampDistanceMeters(0.5)).toBe(1);
  });
});

describe("OilRigTriggerRegistry", () => {
  const LABELS = { small: "Small Oil Rig", large: "Large Oil Rig" };

  function rigWithStep(over: Partial<Parameters<typeof newLogicStep>[0]> = {}, ruleOver: Partial<ReturnType<typeof newLogicRule>> = {}): { profile: RigProfileInput; step: ReturnType<typeof newLogicStep>; rule: ReturnType<typeof newLogicRule> } {
    const step = newLogicStep({ stepType: "StartTimer", timerTarget: "SmallOilRig", ...over });
    const rule = newLogicRule({
      isEnabled: true,
      triggerType: "SmartAlarm",
      triggerEntityId: 12345,
      steps: [step],
      ...ruleOver,
    });
    return { profile: { devices: [], logicRules: [rule] }, step, rule };
  }

  it("registers enabled SmartAlarm rules with rig timers, by ID and distinctive name", () => {
    const reg = new OilRigTriggerRegistry(LABELS);
    const { profile } = rigWithStep({ alarmTextHint: "Crate is coming" }, {});
    reg.rebuild([profile]);

    expect(reg.lookup(12345)).toBe(LABELS.small); // by entity ID
    expect(reg.lookup(null, "Crate is coming")).toBe(LABELS.small); // by hint text
    expect(reg.lookup(null, "something else")).toBeNull();
  });

  it("ignores disabled rules, non-alarm triggers, zero ids and custom timers", () => {
    const reg = new OilRigTriggerRegistry(LABELS);
    reg.rebuild([
      { logicRules: [newLogicRule({ isEnabled: false })] },
      { logicRules: [newLogicRule({ isEnabled: true, triggerType: "SmartSwitch", triggerEntityId: 5 })] },
      { logicRules: [newLogicRule({ isEnabled: true, triggerEntityId: 0 })] },
      rigWithStep({ timerTarget: "Custom", timerName: "my timer" }).profile,
    ]);
    expect(reg.lookup(5, null)).toBeNull();
    expect(reg.lookup(null)).toBeNull();
  });

  it("default alarm texts never take part in text matching (real raids must ring)", () => {
    const reg = new OilRigTriggerRegistry(LABELS);
    const { profile } = rigWithStep({ alarmTextHint: "Your base is under attack!" });
    reg.rebuild([profile]);
    expect(reg.lookup(null, "Your base is under attack!")).toBeNull(); // default text refused
    expect(reg.lookup(12345)).not.toBeNull(); // but entity-ID matching still works

    const shortHint = rigWithStep({ alarmTextHint: "ab" }); // <3 chars refused too
    reg.rebuild([shortHint.profile]);
    expect(reg.lookup(null, "ab")).toBeNull();
  });

  it("device title wins over stale rule hint after in-game rename", () => {
    const reg = new OilRigTriggerRegistry(LABELS);
    const { profile } = rigWithStep({ alarmTextHint: "old name" });
    profile.devices = [{ entityId: 12345, inGameAlarmTitle: "new name" }];
    reg.rebuild([profile]);

    expect(reg.lookup(null, "new name")).toBe(LABELS.small);
    expect(reg.lookup(null, "old name")).toBe(LABELS.small); // hint kept as second key (parity)
  });

  it("learnAlarmText overwrites typed hints with the proven title", () => {
    const reg = new OilRigTriggerRegistry(LABELS);
    const { profile } = rigWithStep({ alarmTextHint: "typo here" });
    const profiles = [profile];

    expect(reg.learnAlarmText(profiles, 12345, "correct name")).toBe(true);
    expect(profile.logicRules![0]!.steps[0]!.alarmTextHint).toBe("correct name");

    // Idempotent when unchanged.
    expect(reg.learnAlarmText(profiles, 12345, "correct name")).toBe(false);

    // Default texts and short strings are refused as learned values.
    expect(reg.learnAlarmText(profiles, 12345, "Alarm")).toBe(false);
    expect(reg.learnAlarmText(profiles, 12345, "xy")).toBe(false);
    expect(reg.learnAlarmText(profiles, 99999, "unrelated")).toBe(false); // no matching rule
  });

  it("targetsForProfile returns untranslated targets; badgesForProfile translates", () => {
    const reg = new OilRigTriggerRegistry(LABELS);
    const { profile } = rigWithStep({ timerTarget: "LargeOilRig" });
    expect(reg.targetsForProfile(profile).get(12345)).toBe("LargeOilRig");
    expect(reg.badgesForProfile(profile).get(12345)).toBe(LABELS.large);
  });
});

describe("AlertTemplateService", () => {
  const defaults = (key: string): string | null =>
    ({
      raid_alert: "Base under attack at {0}!",
      malformed_default: "ok {0}",
    })[key] ?? null;

  it("dotNetFormat handles slots and literal braces like string.Format", () => {
    expect(dotNetFormat("{0} vs {1}", ["a", "b"])).toBe("a vs b");
    expect(dotNetFormat("{{literal}} {0}", ["x"])).toBe("{literal} x");
    expect(() => dotNetFormat("{5}", ["only"])).toThrow();
  });

  it("override wins over resource fallback; remove falls back again", () => {
    const file = join(mkdtempSync(join(tmpdir(), "rpd-alerts-")), "custom_alerts.json");
    const svc = new AlertTemplateService({ filePath: file, culture: "de-DE", defaults });

    expect(svc.getFormattedAlert("raid_alert", ["12:34"])).toBe("Base under attack at 12:34!");
    svc.setOverride("raid_alert", "RAID {0}!!");
    expect(svc.hasOverride("raid_alert")).toBe(true);
    expect(svc.getFormattedAlert("raid_alert", ["12:34"])).toBe("RAID 12:34!!");

    svc.removeOverride("raid_alert");
    expect(svc.hasOverride("raid_alert")).toBe(false);
    rmSync(file, { force: true });
  });

  it("survives restarts (file round-trip) and corrupt files (empty start)", () => {
    const dir = mkdtempSync(join(tmpdir(), "rpd-alerts-"));
    const file = join(dir, "custom_alerts.json");
    const a = new AlertTemplateService({ filePath: file, culture: "en", defaults });
    a.setOverride("raid_alert", "custom {0}");
    expect(JSON.parse(readFileSync(file, "utf8"))).toEqual({ en: { raid_alert: "custom {0}" } });

    const b = new AlertTemplateService({ filePath: file, culture: "en", defaults });
    expect(b.getFormattedAlert("raid_alert", ["now"])).toBe("custom now");

    writeFileSync(file, "{corrupt", "utf8");
    const logs: string[] = [];
    const c = new AlertTemplateService({
      filePath: file,
      culture: "en",
      defaults,
      log: (_l, m) => logs.push(m),
    });
    expect(c.getAlertTemplate("raid_alert")).toBe("Base under attack at {0}!"); // falls back to resource
    expect(logs.join()).toContain("unreadable");

    // Malformed CUSTOM template with no default translation: C# string.Format("") succeeds
    // returning "" — so the observable legacy behavior is an EMPTY alert, not the raw template.
    c.setOverride("malformed_custom_only", "broken {7}");
    expect(c.getFormattedAlert("malformed_custom_only", ["x"])).toBe("");
    rmSync(dir, { recursive: true, force: true });
  });
});
