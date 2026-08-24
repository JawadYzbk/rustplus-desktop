/**
 * `logic/*` channel handlers — Logic Engine control + rule persistence.
 * saveRules keeps stored STEPS for unchanged rule ids (the header-only editor must not be
 * able to destroy step definitions); brand-new ids start with empty steps.
 */
import type { z } from "zod";
import {
  logicGetRules,
  logicGetRule,
  logicGetTimers,
  logicSaveRules,
  type DeviceNodeDto,
} from "@rpd/shared";
import { parseDevices, serializeDevices } from "./services/devices/server-profile.js";
import { countActualDevicesTree } from "./services/devices/device-data.js";
import { parseLogicRule, serializeLogicRule, type LogicRule, type LogicStep } from "./services/automation/logic-rule.js";
import type { LogicEngineService } from "./services/automation/engine-service.js";
import { parsePairingLink } from "./services/rustplus/pairing-parser.js";

export interface ProfileBridgeDeps {
  profiles: {
    list(): Array<{ Name: string; Host: string; Port: number; SteamId64: string }>;
    matchKey(p: { Host: string; Port: number; SteamId64: string }): string;
    upsert(profile: { Name: string; Description?: string; Host: string; Port: number; SteamId64: string } & Record<string, unknown>): void;
    devicesFor(key: string): Record<string, unknown>[] | null;
    saveDevices(key: string, devices: Record<string, unknown>[]): boolean;
    field(key: string, name: string): unknown;
    setField(key: string, name: string, value: unknown): boolean;
  };
  /** Mutable active-profile reference (engine + connection context read it). */
  activeRef: { key: string | null };
}

export function buildProfileHandlers(deps: ProfileBridgeDeps): {
  "profile/list": () => { profiles: Array<{ matchKey: string; name: string; host: string; port: number; steamId64: string; deviceCount: number }> };
  "profile/pair": (req: { link: string; name?: string }) => { activated: boolean; profile: { matchKey: string; name: string; host: string; port: number; steamId64: string; deviceCount: number } };
  "profile/getDevices": (req: { matchKey: string }) => { devices: DeviceNodeDto[]; found: boolean };
  "profile/saveDevices": (req: { matchKey: string; devices: unknown[] }) => { saved: boolean };
  "profile/activate": (req: { matchKey: string }) => { activated: boolean };
} {
  return {
    "profile/list": () => ({
      profiles: deps.profiles.list().map((p) => {
        const key = deps.profiles.matchKey(p);
        const raw = deps.profiles.devicesFor(key) ?? [];
        return {
          matchKey: key,
          name: p.Name,
          host: p.Host,
          port: p.Port,
          steamId64: p.SteamId64,
          deviceCount: countActualDevicesTree(parseDevices(raw)),
        };
      }),
    }),

    "profile/pair": (req) => {
      const payload = parsePairingLink(req.link);
      if (!payload?.SteamId64 || !payload.PlayerToken || !/^\d{17}$/.test(payload.SteamId64)) {
        throw new Error("Invalid Rust+ pairing link. Paste the complete rustplus:// link.");
      }
      const name = req.name?.trim() || payload.ServerName?.trim() || `${payload.Host}:${payload.Port}`;
      const profile = {
        Name: name,
        Description: payload.ServerDescription ?? undefined,
        Host: payload.Host,
        Port: payload.Port,
        SteamId64: payload.SteamId64,
        PlayerToken: payload.PlayerToken,
      };
      deps.profiles.upsert(profile);
      const matchKey = deps.profiles.matchKey(profile);
      deps.activeRef.key = matchKey;
      return {
        activated: true,
        profile: {
          matchKey,
          name,
          host: payload.Host,
          port: payload.Port,
          steamId64: payload.SteamId64,
          deviceCount: countActualDevicesTree(parseDevices(deps.profiles.devicesFor(matchKey) ?? [])),
        },
      };
    },

    "profile/getDevices": (req) => {
      const raw = deps.profiles.devicesFor(req.matchKey);
      if (raw === null) return { devices: [], found: false };
      return { devices: parseDevices(raw) as unknown as DeviceNodeDto[], found: true };
    },

    "profile/saveDevices": (req) => ({
      saved: deps.profiles.saveDevices(req.matchKey, serializeDevices(req.devices as never)),
    }),

    "profile/activate": (req) => {
      const known =
        deps.profiles.list().findIndex((p) => deps.profiles.matchKey(p) === req.matchKey) >= 0;
      if (known) deps.activeRef.key = req.matchKey;
      return { activated: known };
    },
  };
}

export interface EngineDeps {
  engine: {
    status(): ReturnType<LogicEngineService["status"]>;
    requestStop(): void;
    runRule(ruleId: string): Promise<boolean>;
    rulesFor(matchKey: string): unknown[];
    isEngineActiveFor(matchKey: string): boolean;
    saveRulesFor(
      matchKey: string,
      headers: Array<z.infer<typeof logicSaveRules["request"]>["rules"][number]>,
      isEngineActive: boolean,
    ): boolean;
    ruleFor(matchKey: string, ruleId: string): LogicRule | null;
    saveFullRuleFor(matchKey: string, rule: LogicRule): boolean;
    timersFor(matchKey: string): Array<{
      id: string;
      name: string;
      command: string;
      endTimeUtcMs: number;
      enableCountdownAudio: boolean;
      enableAlarmAudio: boolean;
    }>;
    removeTimerFor(matchKey: string, id: string): boolean;
    tryAddTimerFor(
      matchKey: string,
      name: string,
      hours: number,
      minutes: number,
      seconds: number,
    ): { ok: true; id: string } | { ok: false; reason: "limit" | "letter" | "duration" };
  };
}

type RuleHeader = z.infer<typeof logicGetRules["response"]>["rules"][number];

export function buildLogicHandlers(engine: EngineDeps["engine"]): {
  "logic/status": () => ReturnType<LogicEngineService["status"]>;
  "logic/stop": () => { stopped: boolean };
  "logic/run": (req: { ruleId: string }) => Promise<{ accepted: boolean }>;
  "logic/getRules": (req: { matchKey: string }) => { found: boolean; isEngineActive: boolean; rules: RuleHeader[] };
  "logic/saveRules": (req: z.infer<typeof logicSaveRules["request"]>) => { saved: boolean };
  "logic/getRule": (req: { matchKey: string; ruleId: string }) => z.infer<(typeof logicGetRule)["response"]>;
  "logic/saveRule": (req: { matchKey: string; rule: Record<string, unknown> }) => { saved: boolean };
  "logic/getTimers": (req: { matchKey: string }) => z.infer<(typeof logicGetTimers)["response"]>;
  "logic/addTimer": (req: {
    matchKey: string;
    name: string;
    hours: number;
    minutes: number;
    seconds: number;
  }) => { ok: boolean; id: string; reason: "limit" | "letter" | "duration" | null };
  "logic/removeTimer": (req: { matchKey: string; id: string }) => { removed: boolean };
} {
  return {
    "logic/status": () => engine.status(),
    "logic/stop": () => {
      engine.requestStop();
      return { stopped: true };
    },
    "logic/run": async (req) => ({ accepted: await engine.runRule(req.ruleId) }),
    "logic/getRules": (req) => {
      const raw = engine.rulesFor(req.matchKey);
      const rules = raw.map((r) => parseLogicRule((r ?? {}) as Record<string, unknown>));
      return {
        found: true,
        isEngineActive: engine.isEngineActiveFor(req.matchKey),
        rules: rules.map((r): RuleHeader => ({
          id: r.id,
          name: r.name,
          isEnabled: r.isEnabled,
          isLoopEnabled: r.isLoopEnabled,
          loopCount: r.loopCount,
          triggerType: r.triggerType,
          triggerEntityId: r.triggerEntityId,
          triggerCommand: r.triggerCommand,
          triggerRuleId: r.triggerRuleId,
          triggerState: r.triggerState,
          conditionOperator: r.conditionOperator,
          conditionDeviceEntityId: r.conditionDeviceEntityId,
          conditionDeviceState: r.conditionDeviceState,
          stepCount: r.steps.length,
        })),
      };
    },
    "logic/saveRules": (req) => ({
      saved: engine.saveRulesFor(req.matchKey, req.rules, req.isEngineActive),
    }),

    "logic/getRule": (req) => {
      const rule = engine.ruleFor(req.matchKey, req.ruleId);
      return { found: rule !== null, rule: rule ? ruleToDto(rule) : null };
    },

    "logic/saveRule": (req) => {
      // parseLogicRule applies the C# defaults/clamps/unknown-enum tolerance on load.
      const rule = parseLogicRule(req.rule as unknown as Record<string, unknown>);
      return { saved: engine.saveFullRuleFor(req.matchKey, rule) };
    },

    "logic/getTimers": (req: { matchKey: string }) => ({ timers: engine.timersFor(req.matchKey) }),
    "logic/addTimer": (req: { matchKey: string; name: string; hours: number; minutes: number; seconds: number }) => {
      const r = engine.tryAddTimerFor(req.matchKey, req.name, req.hours, req.minutes, req.seconds);
      return r.ok ? { ok: true, id: r.id, reason: null } : { ok: false, id: "", reason: r.reason };
    },
    "logic/removeTimer": (req: { matchKey: string; id: string }) => ({
      removed: engine.removeTimerFor(req.matchKey, req.id),
    }),
  };
}

type StepDto = NonNullable<z.infer<(typeof logicGetRule)["response"]>["rule"]>["steps"][number];

/** LogicStep → bridge DTO (recursive for conditionalSteps). */
function stepToDto(s: LogicStep): StepDto {
  return {
    stepType: s.stepType,
    timerMinutes: s.timerMinutes,
    timerTarget: s.timerTarget,
    timerName: s.timerName,
    showCrateOnMap: s.showCrateOnMap,
    alarmTextHint: s.alarmTextHint,
    waitSeconds: s.waitSeconds,
    targetEntityId: s.targetEntityId,
    targetGroupName: s.targetGroupName,
    toggleState: s.toggleState,
    conditionOperator: s.conditionOperator,
    conditionDeviceIdsCsv: s.conditionDeviceIdsCsv,
    conditionalSteps: s.conditionalSteps.map(stepToDto),
  };
}

function ruleToDto(r: LogicRule): NonNullable<z.infer<(typeof logicGetRule)["response"]>["rule"]> {
  return {
    id: r.id,
    name: r.name,
    isEnabled: r.isEnabled,
    isLoopEnabled: r.isLoopEnabled,
    loopCount: r.loopCount,
    triggerType: r.triggerType,
    triggerEntityId: r.triggerEntityId,
    triggerCommand: r.triggerCommand,
    triggerRuleId: r.triggerRuleId,
    triggerState: r.triggerState,
    conditionOperator: r.conditionOperator,
    conditionDeviceEntityId: r.conditionDeviceEntityId,
    conditionDeviceState: r.conditionDeviceState,
    steps: r.steps.map(stepToDto),
  };
}
