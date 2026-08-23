/**
 * DeviceAutomationEvaluator — line-faithful port of Services/DeviceAutomationEvaluator.cs:
 * proximity matching (Any/All/Specific × Online/Offline) and the overnight time-window check
 * with wrap-around. The C# DEBUG Verify() asserts are golden tests here.
 *
 * SteamId is a string in TS (u64 exceeds Number.MAX_SAFE_INTEGER).
 */
export interface AutomationRuleInput {
  playerMatchMode: string; // AnyOnline | AllOnline | AnyOffline | AllOffline | SpecificOnline | SpecificOffline
  specificPlayerSteamId: string;
  distanceMeters: number;
  startTime: string;
  endTime: string;
}

export interface PlayerSnapshot {
  steamId: string;
  isOnline: boolean;
  x: number | null;
  y: number | null;
}

/** DistanceMeters setter parity: Math.Max(1, value). */
export const clampDistanceMeters = (v: number): number => Math.max(1, v);

export function isProximityMatch(
  rule: Pick<AutomationRuleInput, "playerMatchMode" | "specificPlayerSteamId" | "distanceMeters">,
  anchorX: number,
  anchorY: number,
  players: readonly PlayerSnapshot[],
): boolean {
  const selected =
    rule.playerMatchMode.startsWith("Specific")
      ? players.filter((p) => p.steamId === rule.specificPlayerSteamId)
      : [...players];

  if (rule.playerMatchMode === "AnyOffline") return selected.some((p) => !p.isOnline);
  if (rule.playerMatchMode === "AllOffline") return selected.length > 0 && selected.every((p) => !p.isOnline);
  if (rule.playerMatchMode === "SpecificOffline") return selected.length === 1 && !selected[0]!.isOnline;

  const isNear = (p: PlayerSnapshot): boolean =>
    p.isOnline &&
    p.x !== null &&
    p.y !== null &&
    Math.sqrt((p.x - anchorX) ** 2 + (p.y - anchorY) ** 2) <= rule.distanceMeters;

  return rule.playerMatchMode === "AllOnline"
    ? selected.length > 0 && selected.every(isNear)
    : selected.some(isNear);
}

export function isTimeMatch(rule: Pick<AutomationRuleInput, "startTime" | "endTime">, currentTime: string): boolean {
  const r = tryGetTimeMatch(rule, currentTime);
  return r.parsed && r.matched;
}

/** TryGetTimeMatch parity — parsed=false means the time strings were unparseable at all. */
export function tryGetTimeMatch(
  rule: Pick<AutomationRuleInput, "startTime" | "endTime">,
  currentTime: string,
): { parsed: boolean; matched: boolean } {
  const start = tryMinutes(rule.startTime);
  const end = tryMinutes(rule.endTime);
  const current = tryMinutes(currentTime);
  if (start === null || end === null || current === null) return { parsed: false, matched: false };

  // start == end means "always"; start < end is a normal window; start > end wraps midnight.
  const matched =
    start === end || (start < end ? current >= start && current < end : current >= start || current < end);
  return { parsed: true, matched };
}

/** TimeSpan.TryParse(invariant) → total minutes mod 1440; null when unparseable. */
function tryMinutes(value: string): number | null {
  const m = /^(\d{1,2}):(\d{2})(?::(\d{2}))?$/.exec(value.trim());
  if (!m) return null;
  const h = Number(m[1]);
  const min = Number(m[2]);
  const sec = m[3] !== undefined ? Number(m[3]) : 0;
  if (min > 59 || sec > 59) return null;
  const totalMinutes = Math.floor((h * 3600 + min * 60 + sec) / 60);
  return ((totalMinutes % 1440) + 1440) % 1440;
}
