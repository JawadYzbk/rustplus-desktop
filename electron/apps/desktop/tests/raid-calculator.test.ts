/**
 * Ports of the MSTest trio for already-ported deterministic cores:
 *  - RaidCalculatorTests.cs (9 methods) over the REAL embedded raid-data.json
 *  - SettingsSearchMatcherTests.cs (1)
 * TutorialTests' ProgressStore cases land with the tutorials stage (registry definitions
 * depend on WPF target IDs that have no Electron counterpart yet).
 */
import { mkdtempSync, readFileSync, rmSync } from "node:fs";
import { tmpdir } from "node:os";
import { join } from "node:path";
import { beforeEach, describe, expect, it } from "vitest";
import {
  InvalidDataError,
  loadRaidData,
  parseRaidData,
  removeUnsupportedTargets,
  validateRaidData,
} from "../src/main/services/raid/raid-data-service.js";
import { RaidCalculatorEngine } from "../src/main/services/raid/raid-calculator.js";
import { RaidPlanStore } from "../src/main/services/raid/raid-plan-store.js";
import { raidTargetCategory } from "../src/main/services/raid/raid-models.js";
import { settingsSearchMatches } from "../src/renderer/src/lib/settings-search.js";
import type { RaidDataSet } from "../src/main/services/raid/raid-models.js";

let data: RaidDataSet;
let engine: RaidCalculatorEngine;

beforeEach(() => {
  data = loadRaidData(); // LoadAsync parity: strips boats + validates
  engine = new RaidCalculatorEngine(data);
});

const byName = <T extends { displayName: string }>(items: T[], name: string): T => {
  const hit = items.find((i) => i.displayName === name);
  if (!hit) throw new Error(`fixture missing ${name}`);
  return hit;
};

describe("RaidCalculator (port of RaidCalculatorTests.cs)", () => {
  it("SingleTargetUsesDatasetHitCountAndCraftCost", () => {
    const target = byName(data.targets, "Armored Door");
    const source = byName(data.sources, "Timed Explosive Charge");

    const result = engine
      .getMethods(target)
      .find((m) => m.source.sourceId === source.sourceId);
    expect(result).toBeDefined();
    expect(result!.requiredItems).toBe(data.hits.get(source.sourceId)!.get(target.targetId));
    for (const cost of source.craftCost ?? []) {
      expect(
        result!.resources.find((r) => r.shortname === cost.shortname)?.amount,
      ).toBe(cost.amount * result!.requiredItems);
    }
  });

  it("TargetQuantityMultipliesWholeDatasetHitCount", () => {
    const target = byName(data.targets, "Armored Door");
    const source = byName(data.sources, "Rocket");
    const datasetHits = data.hits.get(source.sourceId)!.get(target.targetId)!;

    const result = engine
      .getMethods(target, 3)
      .find((m) => m.source.sourceId === source.sourceId)!;

    expect(result.requiredItems).toBe(datasetHits * 3);
    // Dataset rounding is whole raid items — hits == ceil(health / damagePerItem).
    expect(Math.ceil(target.startHealth / result.damagePerItem)).toBe(datasetHits);
  });

  it("MultipleTargetsAndMethodsAggregateSharedResources", () => {
    const firstTarget = byName(data.targets, "Armored Door");
    const secondTarget = data.targets.find((t) => t.targetId !== firstTarget.targetId)!;
    const firstSource = byName(data.sources, "Timed Explosive Charge");
    const secondSource = byName(data.sources, "Rocket");
    const first = engine.getMethods(firstTarget, 2).find((m) => m.source.sourceId === firstSource.sourceId)!;
    const second = engine.getMethods(secondTarget).find((m) => m.source.sourceId === secondSource.sourceId)!;

    const totals = RaidCalculatorEngine.aggregate([first, second]);
    const expectedSulfur =
      first.resources.filter((r) => r.shortname === "sulfur").reduce((s, r) => s + r.amount, 0) +
      second.resources.filter((r) => r.shortname === "sulfur").reduce((s, r) => s + r.amount, 0);
    expect(totals.find((r) => r.shortname === "sulfur")?.amount).toBe(expectedSulfur);

    const raidItems = RaidCalculatorEngine.aggregateItems([first, second, first]);
    expect(raidItems.find((i) => i.source.sourceId === first.source.sourceId)?.amount).toBe(
      first.requiredItems * 2,
    );
  });

  it("SmartCombinationUsesWholeItemsAndReachesTargetHealth", () => {
    const metalWall = data.targets.find((t) => t.displayName === "Wall (Metal)")!;
    const selectedSources = data.sources
      .filter((s) => s.itemShortname === "explosive.timed" || s.itemShortname === "ammo.rifle.explosive")
      .map((s) => s.sourceId);
    const mix = engine.getBestCombination(metalWall, selectedSources, "LowestSulfur");

    expect(mix).toHaveLength(2);
    expect(mix.find((p) => p.source.itemShortname === "explosive.timed")?.requiredItems).toBe(3);
    expect(mix.reduce((s, p) => s + p.totalDamage, 0)).toBeGreaterThanOrEqual(metalWall.startHealth);
  });

  it("InvalidQuantityIsClampedAndUnknownMethodIsUnavailable", () => {
    const target = data.targets[0]!;
    const zero = engine.getMethods(target, 0);
    expect(zero.map((m) => m.requiredItems)).toEqual(engine.getMethods(target, 1).map((m) => m.requiredItems));
    expect(zero.some((m) => m.source.sourceId === Number.MAX_SAFE_INTEGER)).toBe(false);
  });

  it("OptionalNullCraftCostIsAcceptedButMalformedNumbersAreRejected", () => {
    validateRaidData(data); // real dataset passes
    expect(data.sources.some((s) => s.craftCost === null)).toBe(true);

    const source = data.sources[0]!;
    const malformed: RaidDataSet = {
      schemaVersion: 1,
      generatedAt: "",
      sources: [{ ...source, rawDamage: Number.NaN }],
      targets: [data.targets[0]!],
      damagePerHit: new Map([[source.sourceId, new Map([[data.targets[0]!.targetId, 1]])]]),
      hits: new Map([[source.sourceId, new Map([[data.targets[0]!.targetId, 1]])]]),
    };
    expect(() => validateRaidData(malformed)).toThrow(InvalidDataError);
  });

  it("PersistenceRoundTripsPlanEntries", () => {
    const directory = mkdtempSync(join(tmpdir(), "rpd-raid-test-"));
    try {
      const store = new RaidPlanStore(join(directory, "plan.json"));
      const expected = { targetId: data.targets[0]!.targetId, quantity: 3, sourceId: data.sources[0]!.sourceId };
      store.save([expected]);

      const restored = store.load();
      expect(restored).toEqual([expected]);
      // Atomic write leaves no tmp residue.
      expect(() => readFileSync(join(directory, "plan.json.tmp"))).toThrow();
    } finally {
      rmSync(directory, { recursive: true, force: true });
    }
  });

  it("TargetCategoriesAndSearchTermsComeFromDatasetFields", () => {
    const door = data.targets.find((t) => t.componentType === "Door")!;
    const matches = data.targets.filter((t) =>
      `${t.displayName} ${t.componentType} ${t.buildingTier ?? ""} ${t.itemCategorySlug ?? ""}`
        .toLowerCase()
        .includes(door.displayName.toLowerCase()),
    );
    // Category getter parity (ComponentType switch → "Doors & gates").
    expect(
      raidTargetCategory({ componentType: door.componentType, itemCategorySlug: door.itemCategorySlug }),
    ).toBe("Doors & gates");
    expect(matches.length).toBeGreaterThan(0);
    expect(matches.some((t) => t.targetId === door.targetId)).toBe(true);
  });

  it("LoadAsync_UnsupportedBoatTargets_RemovesTargetsAndMatrixEntries", () => {
    expect(data.targets.some((t) => t.prefabName.toLowerCase().includes("/building boat/"))).toBe(false);

    // Raw dataset may contain boats; removal must strip both targets and their matrix rows.
    const raw = parseRaidData(readFileSync(join(process.cwd(), "assets", "data", "raid-data.json"), "utf-8"));
    removeUnsupportedTargets(raw);
    expect(raw.targets.some((t) => t.prefabName.toLowerCase().includes("/building boat/"))).toBe(false);
    const boatIds = new Set(
      parseRaidData(readFileSync(join(process.cwd(), "assets", "data", "raid-data.json"), "utf-8"))
        .targets.filter((t) => t.prefabName.toLowerCase().includes("/building boat/"))
        .map((t) => t.targetId),
    );
    for (const row of raw.hits.values()) {
      for (const id of boatIds) expect(row.has(id)).toBe(false);
    }
    validateRaidData(raw);
  });
});

describe("SettingsSearchMatcher (port of SettingsSearchMatcherTests.cs)", () => {
  it("MatchesEverySearchTermAcrossTitleAndKeywords", () => {
    expect(settingsSearchMatches("gpu scale", "Map Performance", "GPU rendering scale cache")).toBe(true);
    expect(settingsSearchMatches("gpu alerts", "Map Performance", "GPU rendering scale cache")).toBe(false);
  });
});
