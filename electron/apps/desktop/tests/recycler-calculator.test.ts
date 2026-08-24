import { beforeAll, describe, expect, it } from "vitest";
import { calculateRecycler, loadRecyclerItems, type RecyclerItem } from "../src/main/services/recycler/recycler-service.js";

let items: RecyclerItem[];

beforeAll(() => {
  items = loadRecyclerItems();
});

describe("RecyclerCalculator", () => {
  it("loads the embedded recyclable catalog and preserves categories", () => {
    expect(items.length).toBeGreaterThan(700);
    expect(items.find((item) => item.shortName === "ammo.shotgun")?.category).toBe("Ammunition");
  });

  it("matches wild and safe-zone probability math plus stack timing", () => {
    const result = calculateRecycler(items, { "ammo.shotgun": 10 });
    const metal = result.outputs.find((output) => output.shortName === "metal.fragments")!;
    const gunpowder = result.outputs.find((output) => output.shortName === "gunpowder")!;

    expect(metal.wild).toMatchObject({ expected: 15, guaranteed: 10, chance: 10, chancePercent: 50 });
    expect(gunpowder.safe).toMatchObject({ expected: 20, guaranteed: 20, chance: 0, chancePercent: 0 });
    expect(result.wildSeconds).toBe(10); // ceil(10 / ceil(64 * .1)) * 5
    expect(result.safeSeconds).toBe(16);
  });
});
