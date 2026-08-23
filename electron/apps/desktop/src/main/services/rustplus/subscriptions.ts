/**
 * Subscription orchestrator — port of PrimeSubscriptionsAsync/PokeEntityAsync semantics
 * (audit RUSTPLUS_CONNECTIVITY §4, §8):
 *  - subscribe once per entity per connection (subOnce set resets on every connect);
 *  - after subscribing, "poke" = immediate GetEntityInfo so device state populates without waiting
 *    for the next broadcast;
 *  - priming is sequential with a 100 ms gap between entities and a 5 s per-entity budget.
 */
import { rq } from "./protocol.js";

export const PRIME_BUDGET_MS = 5_000;
export const PRIME_GAP_MS = 100;

export interface SubscriptionCore {
  needsSubscribeOnce(entityId: number): boolean;
  markSubscribed(entityId: number): void;
}

export interface SubscriptionDeps {
  core: SubscriptionCore;
  /** Wired to the live instance; must reject on timeout/failure. */
  send(data: Record<string, unknown>, timeoutMs?: number): Promise<unknown>;
  sleep?: (ms: number) => Promise<void>;
}

export class SubscriptionOrchestrator {
  private readonly sleep: (ms: number) => Promise<void>;

  constructor(private readonly deps: SubscriptionDeps) {
    this.sleep = deps.sleep ?? ((ms) => new Promise((r) => setTimeout(r, ms)));
  }

  /** Sequentially subscribe+poke each entity not yet subscribed on this connection. */
  async prime(entityIds: readonly number[]): Promise<void> {
    for (let i = 0; i < entityIds.length; i++) {
      const id = entityIds[i]!;
      if (!this.deps.core.needsSubscribeOnce(id)) continue;

      await this.deps.send(rq.subscribeEntity(id), PRIME_BUDGET_MS);
      // Poke: best effort — a failed poke never fails priming.
      try {
        await this.deps.send(rq.getEntityInfo(id), PRIME_BUDGET_MS);
      } catch {
        /* poke is opportunistic */
      }
      this.deps.core.markSubscribed(id);

      if (i < entityIds.length - 1) await this.sleep(PRIME_GAP_MS);
    }
  }

  /** Re-subscribe a single entity (e.g. after server-side drop); no-ops if already subscribed. */
  async ensure(entityId: number): Promise<boolean> {
    if (!this.deps.core.needsSubscribeOnce(entityId)) return false;
    await this.deps.send(rq.subscribeEntity(entityId), PRIME_BUDGET_MS);
    this.deps.core.markSubscribed(entityId);
    return true;
  }
}
