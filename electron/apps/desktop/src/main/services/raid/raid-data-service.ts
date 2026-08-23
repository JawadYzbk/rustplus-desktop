/**
 * Raid data service — port of Services/Raid/RaidDataService.cs: load the embedded dataset,
 * strip unsupported boat targets (and their matrix entries), then validate shape/numbers.
 * The C# reads Assets/Data/raid-data.json from the app base directory; we resolve the same
 * asset relative to this module and fall back to cwd (vitest runs from apps/desktop).
 */
import { existsSync, readFileSync } from "node:fs";
import { join } from "node:path";
import type { RaidDataSet, RaidResourceCost, RaidSource, RaidTarget } from "./raid-models.js";

function findRaidDataPath(): string {
  const candidates = [
    join(process.cwd(), "assets", "data", "raid-data.json"),
    // out/main → apps/desktop/assets/data
    join(__dirname, "..", "..", "assets", "data", "raid-data.json"),
    // src/main/services/raid → apps/desktop/assets/data
    join(__dirname, "..", "..", "..", "..", "assets", "data", "raid-data.json"),
  ];
  for (const p of candidates) {
    if (existsSync(p)) return p;
  }
  throw new Error("The packaged raid-data.json asset is missing.");
}

const asFinite = (v: unknown): number => (typeof v === "number" && Number.isFinite(v) ? v : 0);
const asStr = (v: unknown): string => (typeof v === "string" ? v : "");

/** The shipped dataset is camelCase; C# reads it case-insensitively — accept both casings. */
const pick = (obj: Record<string, unknown>, pascal: string): unknown =>
  obj[pascal] ?? obj[pascal.charAt(0).toLowerCase() + pascal.slice(1)];

function parseMatrix(raw: unknown): Map<number, Map<number, number>> {
  const out = new Map<number, Map<number, number>>();
  if (raw && typeof raw === "object") {
    for (const [src, inner] of Object.entries(raw as Record<string, unknown>)) {
      const row = new Map<number, number>();
      if (inner && typeof inner === "object") {
        for (const [tgt, val] of Object.entries(inner as Record<string, unknown>)) {
          row.set(Number(tgt), asFinite(val));
        }
      }
      out.set(Number(src), row);
    }
  }
  return out;
}

export function parseRaidData(json: string): RaidDataSet {
  const raw = JSON.parse(json) as Record<string, unknown>;
  const sources = Array.isArray(pick(raw, "Sources")) ? (pick(raw, "Sources") as Record<string, unknown>[]) : [];
  const targets = Array.isArray(pick(raw, "Targets")) ? (pick(raw, "Targets") as Record<string, unknown>[]) : [];
  const data: RaidDataSet = {
    schemaVersion: asFinite(pick(raw, "SchemaVersion")),
    generatedAt: asStr(pick(raw, "GeneratedAt")),
    sources: sources.map((s): RaidSource => ({
      sourceId: asFinite(pick(s, "SourceId")),
      prefabName: asStr(pick(s, "PrefabName")),
      itemId: typeof pick(s, "ItemId") === "number" ? (pick(s, "ItemId") as number) : null,
      itemShortname: asStr(pick(s, "ItemShortname")),
      itemSlug: asStr(pick(s, "ItemSlug")),
      itemCategorySlug: asStr(pick(s, "ItemCategorySlug")),
      displayName: asStr(pick(s, "DisplayName")),
      kind: asStr(pick(s, "Kind")),
      rawDamage: asFinite(pick(s, "RawDamage")),
      damageTypes: Object.fromEntries(
        Object.entries((pick(s, "DamageTypes") ?? {}) as Record<string, unknown>).map(([k, v]) => [k, asFinite(v)]),
      ),
      craftCost:
        pick(s, "CraftCost") == null
          ? null
          : (pick(s, "CraftCost") as Record<string, unknown>[]).map((c): RaidResourceCost => ({
              shortname: asStr(pick(c, "Shortname")),
              itemId: asFinite(pick(c, "ItemId")),
              displayName: asStr(pick(c, "DisplayName")),
              amount: asFinite(pick(c, "Amount")),
            })),
      workbenchLevelRequired:
        typeof pick(s, "WorkbenchLevelRequired") === "number" ? (pick(s, "WorkbenchLevelRequired") as number) : null,
    })),
    targets: targets.map((t): RaidTarget => ({
      targetId: asFinite(pick(t, "TargetId")),
      prefabName: asStr(pick(t, "PrefabName")),
      itemId: typeof pick(t, "ItemId") === "number" ? (pick(t, "ItemId") as number) : null,
      itemShortname: typeof pick(t, "ItemShortname") === "string" ? (pick(t, "ItemShortname") as string) : null,
      itemSlug: typeof pick(t, "ItemSlug") === "string" ? (pick(t, "ItemSlug") as string) : null,
      itemCategorySlug:
        typeof pick(t, "ItemCategorySlug") === "string" ? (pick(t, "ItemCategorySlug") as string) : null,
      buildingSlug: typeof pick(t, "BuildingSlug") === "string" ? (pick(t, "BuildingSlug") as string) : null,
      buildingImage: typeof pick(t, "BuildingImage") === "string" ? (pick(t, "BuildingImage") as string) : null,
      displayName: asStr(pick(t, "DisplayName")),
      buildingTier: typeof pick(t, "BuildingTier") === "string" ? (pick(t, "BuildingTier") as string) : null,
      componentType: asStr(pick(t, "ComponentType")),
      startHealth: asFinite(pick(t, "StartHealth")),
    })),
    damagePerHit: parseMatrix(pick(raw, "DamagePerHit")),
    hits: parseMatrix(pick(raw, "Hits")),
  };
  return data;
}

/** RemoveUnsupportedTargets parity: "/building boat/" prefabs are not raidable in-app. */
export function removeUnsupportedTargets(data: RaidDataSet): void {
  const ids = new Set<number>();
  for (const t of data.targets) {
    if (t.prefabName.toLowerCase().includes("/building boat/")) ids.add(t.targetId);
  }
  if (ids.size === 0) return;
  data.targets = data.targets.filter((t) => !ids.has(t.targetId));
  for (const row of data.damagePerHit.values()) {
    for (const id of ids) row.delete(id);
  }
  for (const row of data.hits.values()) {
    for (const id of ids) row.delete(id);
  }
}

class InvalidDataError extends Error {}
export { InvalidDataError };

/** Validate parity — schema version, presence, ids, finite non-negative numbers, matrices. */
export function validateRaidData(data: RaidDataSet): void {
  if (data.schemaVersion !== 1) {
    throw new InvalidDataError(`Unsupported raid data schema version ${data.schemaVersion}.`);
  }
  if (data.sources.length === 0 || data.targets.length === 0 || data.hits.size === 0 || data.damagePerHit.size === 0) {
    throw new InvalidDataError("Raid data must contain sources, targets, hit counts, and damage values.");
  }

  const sourceIds = new Set<number>();
  for (const source of data.sources) {
    if (source.sourceId <= 0 || sourceIds.has(source.sourceId) || source.displayName.trim().length === 0) {
      throw new InvalidDataError("Raid data contains an invalid or duplicate source.");
    }
    sourceIds.add(source.sourceId);
    const badCraft =
      source.craftCost?.some((c) => c.shortname.trim().length === 0 || !Number.isFinite(c.amount) || c.amount < 0) ===
      true;
    const badDamageTypes = Object.values(source.damageTypes).some((v) => !Number.isFinite(v) || v < 0);
    if (!Number.isFinite(source.rawDamage) || source.rawDamage < 0 || badDamageTypes || badCraft) {
      throw new InvalidDataError(`Raid source '${source.displayName}' contains malformed numeric data.`);
    }
  }

  const targetIds = new Set<number>();
  for (const target of data.targets) {
    if (
      target.targetId <= 0 ||
      targetIds.has(target.targetId) ||
      target.displayName.trim().length === 0 ||
      !Number.isFinite(target.startHealth) ||
      target.startHealth <= 0
    ) {
      throw new InvalidDataError("Raid data contains an invalid or duplicate target.");
    }
    targetIds.add(target.targetId);
  }

  validateMatrix(data.damagePerHit, sourceIds, targetIds, (v) => Number.isFinite(v) && v > 0, "damage");
  validateMatrix(data.hits, sourceIds, targetIds, (v) => v > 0, "hit count");
}

function validateMatrix(
  matrix: Map<number, Map<number, number>>,
  sourceIds: Set<number>,
  targetIds: Set<number>,
  validValue: (v: number) => boolean,
  valueName: string,
): void {
  for (const [sourceId, values] of matrix) {
    if (!sourceIds.has(sourceId)) {
      throw new InvalidDataError(`Raid matrix references unknown source ${sourceId}.`);
    }
    for (const [targetId, value] of values) {
      if (!targetIds.has(targetId) || !validValue(value)) {
        throw new InvalidDataError(`Raid matrix contains an invalid ${valueName} for target ${targetId}.`);
      }
    }
  }
}

/** LoadAsync parity. */
export function loadRaidData(): RaidDataSet {
  const json = readFileSync(findRaidDataPath(), "utf-8");
  const data = parseRaidData(json);
  removeUnsupportedTargets(data);
  validateRaidData(data);
  return data;
}
