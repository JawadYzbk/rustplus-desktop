/** Typed bridge for the profile-scoped Device Automation editor. */
import type { z } from "zod";
import { deviceAutomationGetRules, deviceAutomationSaveRules } from "@rpd/shared";
import {
  parseAutomationRule,
  serializeAutomationRule,
  type DeviceAutomationRule,
} from "./services/devices/server-profile.js";

type RuleDto = z.infer<typeof deviceAutomationGetRules["response"]>["rules"][number];
type SaveRequest = z.infer<typeof deviceAutomationSaveRules["request"]>;

export interface DeviceAutomationProfileDeps {
  profiles: {
    list(): Array<{ Name: string; Host: string; Port: number; SteamId64: string }>;
    matchKey(profile: { Host: string; Port: number; SteamId64: string }): string;
    field(key: string, name: string): unknown;
    setField(key: string, name: string, value: unknown): boolean;
  };
}

export function buildDeviceAutomationHandlers(deps: DeviceAutomationProfileDeps): {
  "deviceAutomation/getRules": (req: { matchKey: string }) => {
    found: boolean;
    isActive: boolean;
    rules: RuleDto[];
  };
  "deviceAutomation/saveRules": (req: SaveRequest) => { saved: boolean };
} {
  const knownProfile = (key: string): boolean =>
    deps.profiles.list().some((profile) => deps.profiles.matchKey(profile) === key);

  return {
    "deviceAutomation/getRules": (req) => {
      if (!knownProfile(req.matchKey)) return { found: false, isActive: false, rules: [] };
      const raw = deps.profiles.field(req.matchKey, "DeviceAutomationRules");
      const rules = Array.isArray(raw)
        ? raw.map((rule) => toDto(parseAutomationRule((rule ?? {}) as Record<string, unknown>)))
        : [];
      return {
        found: true,
        isActive: deps.profiles.field(req.matchKey, "IsDeviceAutomationActive") === true,
        rules,
      };
    },

    "deviceAutomation/saveRules": (req) => {
      if (!knownProfile(req.matchKey)) return { saved: false };
      const rules = req.rules.map((rule) => serializeAutomationRule(fromDto(rule)));
      const rulesSaved = deps.profiles.setField(req.matchKey, "DeviceAutomationRules", rules);
      const activeSaved = deps.profiles.setField(req.matchKey, "IsDeviceAutomationActive", req.isActive);
      return { saved: rulesSaved && activeSaved };
    },
  };
}

function fromDto(rule: RuleDto): DeviceAutomationRule {
  return parseAutomationRule({
    Id: rule.id,
    Name: rule.name,
    IsEnabled: rule.isEnabled,
    IsExpanded: rule.isExpanded,
    ConditionType: rule.conditionType,
    PlayerMatchMode: rule.playerMatchMode,
    SpecificPlayerSteamId: rule.specificPlayerSteamId,
    LocationEntityId: rule.locationEntityId,
    DistanceMeters: rule.distanceMeters,
    StartTime: rule.startTime,
    EndTime: rule.endTime,
    TargetEntityId: rule.targetEntityId,
    MatchedState: rule.matchedState,
    UnmatchedState: rule.unmatchedState,
  });
}

function toDto(rule: DeviceAutomationRule): RuleDto {
  return {
    id: rule.id,
    name: rule.name,
    isEnabled: rule.isEnabled,
    isExpanded: rule.isExpanded,
    conditionType: rule.conditionType as RuleDto["conditionType"],
    playerMatchMode: rule.playerMatchMode as RuleDto["playerMatchMode"],
    specificPlayerSteamId: rule.specificPlayerSteamId,
    locationEntityId: rule.locationEntityId,
    distanceMeters: rule.distanceMeters,
    startTime: rule.startTime,
    endTime: rule.endTime,
    targetEntityId: rule.targetEntityId,
    matchedState: rule.matchedState,
    unmatchedState: rule.unmatchedState,
  };
}
