/**
 * Raid calculator engine — port of Services/RaidCalculatorEngine.cs (L1-164).
 * GetMethods / Recommend / GetBestCombination (integer DP over scaled damage with
 * hundredth-HP fallback) / Aggregate / AggregateItems. Hit counts from raid-data.json are
 * authoritative and already rounded to whole raid items; quantity multiplies AFTER rounding.
 */
import type {
  RaidComparisonMode,
  RaidItemTotal,
  RaidMethodResult,
  RaidResourceTotal,
  RaidSource,
  RaidTarget,
} from "./raid-models.js";
import type { RaidDataSet } from "./raid-models.js";

interface CombinationState {
  firstCost: number;
  secondCost: number;
  items: number;
  actualDamage: number;
  previousDamage: number;
  methodIndex: number;
}

function isBetterThan(a: CombinationState, b: CombinationState): boolean {
  return (
    a.firstCost < b.firstCost ||
    (a.firstCost === b.firstCost &&
      (a.secondCost < b.secondCost ||
        (a.secondCost === b.secondCost &&
          (a.items < b.items || (a.items === b.items && a.actualDamage < b.actualDamage)))))
  );
}

export class RaidCalculatorEngine {
  private readonly sources = new Map<number, RaidSource>();

  constructor(private readonly data: RaidDataSet) {
    for (const source of data.sources) this.sources.set(source.sourceId, source);
  }

  /** GetMethods parity (L14-38). */
  getMethods(target: RaidTarget, targetQuantity = 1): RaidMethodResult[] {
    const quantity = Math.max(1, targetQuantity);
    const methods: RaidMethodResult[] = [];
    for (const [sourceId, hitCounts] of this.data.hits) {
      const hits = hitCounts.get(target.targetId);
      if (hits === undefined || hits <= 0) continue;
      const source = this.sources.get(sourceId);
      if (!source) continue;

      // Multiply only after dataset rounding so multi-target plans cannot under-count.
      const requiredItems = hits * quantity;
      const damage = this.data.damagePerHit.get(sourceId)?.get(target.targetId) ?? 0;
      const resources: RaidResourceTotal[] =
        source.craftCost === null
          ? []
          : source.craftCost.map((cost) => ({
              shortname: cost.shortname,
              itemId: cost.itemId,
              displayName: cost.displayName,
              amount: cost.amount * requiredItems,
            }));
      const totalDamage = damage * requiredItems;
      methods.push({
        source,
        requiredItems,
        damagePerItem: damage,
        totalDamage,
        overkill: Math.max(0, totalDamage - target.startHealth * quantity),
        resources,
        hasCraftCost: source.craftCost !== null,
      });
    }
    return methods;
  }

  /** Recommend parity (L41-56). */
  static recommend(methods: RaidMethodResult[], mode: RaidComparisonMode): RaidMethodResult | null {
    if (methods.length === 0 || mode === "Custom") return null;
    switch (mode) {
      case "LowestSulfur":
        return [...methods]
          .filter((m) => m.hasCraftCost)
          .sort((a, b) => {
            const sa = a.resources.find((c) => c.shortname === "sulfur")?.amount ?? 0;
            const sb = b.resources.find((c) => c.shortname === "sulfur")?.amount ?? 0;
            return sa - sb || a.requiredItems - b.requiredItems;
          })[0] ?? null;
      case "LowestTotalResources":
        return [...methods]
          .filter((m) => m.hasCraftCost)
          .sort((a, b) => {
            const ta = a.resources.reduce((sum, c) => sum + c.amount, 0);
            const tb = b.resources.reduce((sum, c) => sum + c.amount, 0);
            return ta - tb || a.requiredItems - b.requiredItems;
          })[0] ?? null;
      case "FewestRaidItems":
        return (
          [...methods].sort(
            (a, b) => a.requiredItems - b.requiredItems || Number(a.hasCraftCost ? 0 : 1) - Number(b.hasCraftCost ? 0 : 1),
          )[0] ?? null
        );
      default:
        return null;
    }
  }

  private static gcd(left: number, right: number): number {
    let a = left;
    let b = right;
    while (b !== 0) {
      [a, b] = [b, a % b];
    }
    return Math.abs(a);
  }

  private static createResult(method: RaidMethodResult, count: number, targetHealth: number): RaidMethodResult {
    const resources: RaidResourceTotal[] = method.hasCraftCost
      ? (method.source.craftCost ?? []).map((cost) => ({
          shortname: cost.shortname,
          itemId: cost.itemId,
          displayName: cost.displayName,
          amount: cost.amount * count,
        }))
      : [];
    const totalDamage = method.damagePerItem * count;
    return {
      source: method.source,
      requiredItems: count,
      damagePerItem: method.damagePerItem,
      totalDamage,
      overkill: Math.max(0, totalDamage - targetHealth),
      resources,
      hasCraftCost: method.hasCraftCost,
    };
  }

  /** GetBestCombination parity (L59-113): DP over scaled damage, ponytail fallback at scale=100. */
  getBestCombination(
    target: RaidTarget,
    sourceIds: readonly number[],
    mode: RaidComparisonMode,
    targetQuantity = 1,
  ): RaidMethodResult[] {
    const wanted = new Set(sourceIds);
    const methods = this.getMethods(target)
      .filter((m) => wanted.has(m.source.sourceId))
      .filter((m) => mode === "FewestRaidItems" || m.hasCraftCost);
    if (methods.length === 0) return [];

    let scale = 10_000;
    let scaledDamage = methods.map((m) => Math.max(1, Math.round(m.damagePerItem * scale)));
    let divisor = scaledDamage.reduce((acc, v) => RaidCalculatorEngine.gcd(acc, v));
    let health = Math.max(1, Math.ceil((target.startHealth * scale) / divisor));
    if (health > 2_000_000) {
      // ponytail: hundredth-HP fallback caps memory (C# comment verbatim).
      scale = 100;
      scaledDamage = methods.map((m) => Math.max(1, Math.floor(m.damagePerItem * scale)));
      divisor = scaledDamage.reduce((acc, v) => RaidCalculatorEngine.gcd(acc, v));
      health = Math.max(1, Math.ceil((target.startHealth * scale) / divisor));
    }
    const damage = scaledDamage.map((v) => Math.max(1, Math.floor(v / divisor)));
    const best: Array<CombinationState | undefined> = new Array(health + 1).fill(undefined);
    best[0] = { firstCost: 0, secondCost: 0, items: 0, actualDamage: 0, previousDamage: -1, methodIndex: -1 };

    for (let dealt = 0; dealt < health; dealt++) {
      const current = best[dealt];
      if (current === undefined) continue;
      for (let methodIndex = 0; methodIndex < methods.length; methodIndex++) {
        const method = methods[methodIndex]!;
        const nextDamage = Math.min(health, dealt + damage[methodIndex]!);
        const sulfur = method.source.craftCost?.find((r) => r.shortname.toLowerCase() === "sulfur")?.amount ?? 0;
        const totalResources = method.source.craftCost?.reduce((s, r) => s + r.amount, 0) ?? 0;
        const costs: [number, number] =
          mode === "LowestTotalResources"
            ? [totalResources, sulfur]
            : mode === "FewestRaidItems"
              ? [1, sulfur]
              : [sulfur, totalResources];
        const candidate: CombinationState = {
          firstCost: current.firstCost + costs[0],
          secondCost: current.secondCost + costs[1],
          items: current.items + 1,
          actualDamage: current.actualDamage + method.damagePerItem,
          previousDamage: dealt,
          methodIndex,
        };
        const incumbent = best[nextDamage];
        if (incumbent === undefined || isBetterThan(candidate, incumbent)) best[nextDamage] = candidate;
      }
    }
    if (best[health] === undefined) return [];

    const counts = new Array<number>(methods.length).fill(0);
    for (let state = health; state > 0; ) {
      const step = best[state]!;
      counts[step.methodIndex]!++;
      state = step.previousDamage;
    }
    const quantity = Math.max(1, targetQuantity);
    const out: RaidMethodResult[] = [];
    methods.forEach((method, index) => {
      const count = counts[index]! * quantity;
      if (count > 0) out.push(RaidCalculatorEngine.createResult(method, count, target.startHealth * quantity));
    });
    return out;
  }

  /** Aggregate parity — shared-resource totals grouped by shortname (case-insensitive). */
  static aggregate(methods: readonly RaidMethodResult[]): RaidResourceTotal[] {
    const groups = new Map<string, { total: RaidResourceTotal; amount: number }>();
    for (const method of methods) {
      for (const resource of method.resources) {
        const key = resource.shortname.toLowerCase();
        const g = groups.get(key);
        if (g) g.amount += resource.amount;
        else groups.set(key, { total: resource, amount: resource.amount });
      }
    }
    return [...groups.values()]
      .map((g): RaidResourceTotal => ({ ...g.total, amount: g.amount }))
      .sort((a, b) => {
        const sa = a.shortname.toLowerCase() === "sulfur" ? 1 : 0;
        const sb = b.shortname.toLowerCase() === "sulfur" ? 1 : 0;
        return sb - sa || a.displayName.localeCompare(b.displayName);
      });
  }

  /** AggregateItems parity — per-source item totals ordered by amount desc, display name asc. */
  static aggregateItems(methods: readonly RaidMethodResult[]): RaidItemTotal[] {
    const groups = new Map<number, { source: RaidSource; amount: number }>();
    for (const method of methods) {
      const g = groups.get(method.source.sourceId);
      if (g) g.amount += method.requiredItems;
      else groups.set(method.source.sourceId, { source: method.source, amount: method.requiredItems });
    }
    return [...groups.values()].sort((a, b) => b.amount - a.amount || a.source.displayName.localeCompare(b.source.displayName));
  }
}
