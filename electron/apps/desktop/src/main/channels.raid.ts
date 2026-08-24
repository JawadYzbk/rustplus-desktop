import { raidCalculate, raidGetData } from "@rpd/shared";
import { RaidCalculatorEngine } from "./services/raid/raid-calculator.js";
import { raidTargetCategory, type RaidDataSet } from "./services/raid/raid-models.js";

export interface RaidBridgeDeps {
  data: RaidDataSet;
  engine: RaidCalculatorEngine;
}

export function buildRaidHandlers(deps: RaidBridgeDeps): {
  "raid/getData": () => ReturnType<typeof raidGetData["response"]["parse"]>;
  "raid/calculate": (req: ReturnType<typeof raidCalculate["request"]["parse"]>) => ReturnType<typeof raidCalculate["response"]["parse"]>;
} {
  const sources = deps.data.sources.map((source) => ({ ...source }));
  const targets = deps.data.targets.map((target) => ({
    ...target,
    category: raidTargetCategory(target),
  }));

  return {
    "raid/getData": () => ({ sources, targets }),
    "raid/calculate": (req) => {
      const target = deps.data.targets.find((item) => item.targetId === req.targetId);
      if (!target) throw new Error(`Raid target ${req.targetId} was not found.`);
      const methods = deps.engine.getMethods(target, req.targetQuantity);
      const selected = deps.engine.getBestCombination(target, req.sourceIds, req.mode, req.targetQuantity);
      return {
        methods,
        recommended: RaidCalculatorEngine.recommend(methods, req.mode),
        combination: selected,
        resources: RaidCalculatorEngine.aggregate(selected),
        items: RaidCalculatorEngine.aggregateItems(selected),
      };
    },
  };
}
