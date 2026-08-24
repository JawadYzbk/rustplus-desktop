import { deathsClear, deathsGetStats } from "@rpd/shared";
import type { DeathService } from "./services/deaths/death-service.js";

export function buildDeathHandlers(deaths: DeathService): {
  "deaths/getStats": (request: ReturnType<typeof deathsGetStats["request"]["parse"]>) => ReturnType<typeof deathsGetStats["response"]["parse"]>;
  "deaths/clear": () => ReturnType<typeof deathsClear["response"]["parse"]>;
} {
  return {
    "deaths/getStats": (request) => deaths.stats(request),
    "deaths/clear": () => ({ cleared: deaths.clear() }),
  };
}
