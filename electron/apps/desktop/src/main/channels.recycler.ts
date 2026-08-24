import { recyclerCalculate } from "@rpd/shared";
import type { z } from "zod";
import { calculateRecycler, type RecyclerItem } from "./services/recycler/recycler-service.js";

export interface RecyclerBridgeDeps {
  items: RecyclerItem[];
}

export function buildRecyclerHandlers(deps: RecyclerBridgeDeps): {
  "recycler/getData": () => { items: Array<Pick<RecyclerItem, "id" | "shortName" | "displayName" | "category" | "stackSize">> };
  "recycler/calculate": (request: z.infer<typeof recyclerCalculate["request"]>) => z.infer<typeof recyclerCalculate["response"]>;
} {
  const items = deps.items.map(({ id, shortName, displayName, category, stackSize }) => ({ id, shortName, displayName, category, stackSize }));
  return {
    "recycler/getData": () => ({ items }),
    "recycler/calculate": (request) => {
      const quantities = Object.fromEntries(request.quantities.map((entry) => [entry.shortName, entry.quantity]));
      return calculateRecycler(deps.items, quantities);
    },
  };
}
