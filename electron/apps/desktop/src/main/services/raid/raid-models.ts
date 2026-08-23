/**
 * Raid data models — port of Models/Raid/RaidDataModels.cs (headless parts; WPF image
 * bindings are UI concerns). Field names follow the dataset JSON (PascalCase in the file,
 * camelCase here via explicit mapping in raid-data-service.ts).
 */

export interface RaidSource {
  sourceId: number;
  prefabName: string;
  itemId: number | null;
  itemShortname: string;
  itemSlug: string;
  itemCategorySlug: string;
  displayName: string;
  kind: string;
  rawDamage: number;
  damageTypes: Record<string, number>;
  craftCost: RaidResourceCost[] | null;
  workbenchLevelRequired: number | null;
}

export interface RaidTarget {
  targetId: number;
  prefabName: string;
  itemId: number | null;
  itemShortname: string | null;
  itemSlug: string | null;
  itemCategorySlug: string | null;
  buildingSlug: string | null;
  buildingImage: string | null;
  displayName: string;
  buildingTier: string | null;
  componentType: string;
  startHealth: number;
}

/** Category parity — ComponentType switch from RaidDataModels.cs. */
export function raidTargetCategory(t: Pick<RaidTarget, "componentType" | "itemCategorySlug">): string {
  switch (t.componentType) {
    case "Door":
    case "Gate":
      return "Doors & gates";
    case "BuildingBlock":
    case "SimpleBuildingBlock":
      return "Building structures";
    case "Barricade":
      return t.itemCategorySlug === "traps" ? "Traps" : "Barricades";
    case "BaseOven":
    case "BoxStorage":
      return "Deployables";
    default:
      return "Other";
  }
}

export interface RaidResourceCost {
  shortname: string;
  itemId: number;
  displayName: string;
  amount: number;
}

export interface RaidDataSet {
  schemaVersion: number;
  generatedAt: string;
  sources: RaidSource[];
  targets: RaidTarget[];
  damagePerHit: Map<number, Map<number, number>>;
  hits: Map<number, Map<number, number>>;
}

export interface RaidMethodResult {
  source: RaidSource;
  requiredItems: number;
  damagePerItem: number;
  totalDamage: number;
  overkill: number;
  resources: RaidResourceTotal[];
  hasCraftCost: boolean;
}

export interface RaidResourceTotal {
  shortname: string;
  itemId: number;
  displayName: string;
  amount: number;
}

export interface RaidItemTotal {
  source: RaidSource;
  amount: number;
}

export interface RaidPlanEntry {
  targetId: number;
  quantity: number;
  sourceId: number;
}

export type RaidComparisonMode = "LowestSulfur" | "LowestTotalResources" | "FewestRaidItems" | "Custom";
