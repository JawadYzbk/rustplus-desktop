/**
 * Connection timing contracts ported 1:1 from RustPlusClientReal.cs (audit RUSTPLUS_CONNECTIVITY §5).
 * All classes are pure/injectable-clock so golden tests pin the exact legacy numbers.
 */

/** Token bucket — legacy: cap 50, refill 25 tokens/s, busy-wait loops of 333 ms (AcquireTokenAsync L52-94). */
export class RateLimiter {
  static readonly CAPACITY = 50;
  static readonly REFILL_PER_SECOND = 25;
  static readonly WAIT_STEP_MS = 333;

  private tokens: number = RateLimiter.CAPACITY;
  private lastRefill: number;

  constructor(
    private readonly clock: Clock = realClock,
    private readonly capacity = RateLimiter.CAPACITY,
    private readonly refillPerSecond = RateLimiter.REFILL_PER_SECOND,
  ) {
    this.lastRefill = clock.now();
  }

  /** Resolves when a token is consumed; waits in 333 ms steps while the bucket is empty. */
  async acquire(): Promise<void> {
    for (;;) {
      this.refill();
      if (this.tokens >= 1) {
        this.tokens -= 1;
        return;
      }
      await this.clock.sleep(RateLimiter.WAIT_STEP_MS);
    }
  }

  /** Test hook: current token count after refill. */
  peek(): number {
    this.refill();
    return this.tokens;
  }

  private refill(): void {
    const now = this.clock.now();
    const elapsed = Math.max(0, now - this.lastRefill);
    if (elapsed <= 0) return;
    this.tokens = Math.min(this.capacity, this.tokens + (elapsed / 1000) * this.refillPerSecond);
    this.lastRefill = now;
  }
}

export interface Clock {
  now(): number;
  sleep(ms: number): Promise<void>;
}

export const realClock: Clock = {
  now: () => Date.now(),
  sleep: (ms) => new Promise((r) => setTimeout(r, ms)),
};

/** Error classification — legacy CheckConnectionLost (L96-136), inner-exception walk included. */
export type LossKind = "immediate" | "timeout" | "other";

const IMMEDIATE_PATTERNS = ["not connected", "connection closed", "socket", "eof", "unable to read"];
const TIMEOUT_PATTERNS = ["timed out", "timeout", "canceled", "cancelled"];

export function classifyError(err: unknown): LossKind {
  // Legacy walks InnerException chain; in JS we walk `cause` + message.
  const messages: string[] = [];
  let cur: unknown = err;
  for (let depth = 0; cur !== null && cur !== undefined && depth < 8; depth++) {
    if (typeof cur === "object") {
      const e = cur as { message?: unknown; name?: unknown; cause?: unknown };
      if (typeof e.message === "string") messages.push(e.message.toLowerCase());
      if (typeof e.name === "string" && /timeout|abort/i.test(e.name)) return "timeout";
      cur = e.cause;
    } else {
      messages.push(String(cur).toLowerCase());
      break;
    }
  }
  const joined = messages.join(" | ");
  if (IMMEDIATE_PATTERNS.some((p) => joined.includes(p))) return "immediate";
  if (TIMEOUT_PATTERNS.some((p) => joined.includes(p))) return "timeout";
  return "other";
}

/** 5 consecutive timeouts → ConnectionLost; immediate-loss errors fire instantly; success resets. */
export class TimeoutDetector {
  static readonly CONSECUTIVE_LIMIT = 5;

  private consecutive = 0;

  constructor(private readonly limit = TimeoutDetector.CONSECUTIVE_LIMIT) {}

  /** Returns true when the connection should be declared lost. */
  record(err: unknown): boolean {
    const kind = classifyError(err);
    if (kind === "immediate") {
      this.consecutive = 0; // legacy resets the counter on immediate loss
      return true;
    }
    if (kind === "timeout") {
      this.consecutive += 1;
      return this.consecutive >= this.limit;
    }
    return false;
  }

  /** A valid response resets the counter (legacy L7470). */
  success(): void {
    this.consecutive = 0;
  }

  get current(): number {
    return this.consecutive;
  }
}

/** Exponential backoff — legacy OnConnectionLost: 2 s ×2 → max 60 s. */
export class BackoffPolicy {
  private current: number;

  constructor(
    private readonly initialMs = 2_000,
    private readonly maxMs = 60_000,
  ) {
    this.current = initialMs;
  }

  next(): number {
    const value = this.current;
    this.current = Math.min(this.maxMs, this.current * 2);
    return value;
  }

  reset(): void {
    this.current = this.initialMs;
  }

  get peek(): number {
    return this.current;
  }
}
