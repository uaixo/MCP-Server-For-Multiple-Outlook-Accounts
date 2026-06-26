/**
 * Bounded, jittered retry with the no-duplicate-side-effect policy
 * (NFR-REL-2/3).
 *
 * Retry decisions key off {@link GraphError.category} and the request's
 * {@link RetryClass}:
 * - `safe` (read / organise): retry on 429, transient 5xx, timeouts, and
 *   transport failures.
 * - `nonDuplicable` (send / draft-create): retry ONLY on a pre-processing 429
 *   (a throttle that ran before any side effect). Never retry an ambiguous
 *   failure (5xx / timeout / transport) that may have already taken effect —
 *   this is what guarantees no duplicate sends.
 *
 * A `Retry-After` delay from the server is honoured when present.
 */

import type { RetryClass } from "../domain/contracts.js";
import { GraphError } from "./errors.js";

export interface RetryOptions {
  readonly maxRetries?: number;
  readonly baseDelayMs?: number;
  readonly maxDelayMs?: number;
  readonly sleep?: (ms: number) => Promise<void>;
  readonly random?: () => number;
}

const DEFAULTS = {
  maxRetries: 3,
  baseDelayMs: 500,
  maxDelayMs: 8_000,
} as const;

const defaultSleep = (ms: number): Promise<void> => new Promise((r) => setTimeout(r, ms));

export function shouldRetry(err: GraphError, retryClass: RetryClass): boolean {
  if (err.category === "rateLimited") return true; // pre-processing throttle — safe for both classes.
  if (retryClass === "nonDuplicable") return false; // never retry an ambiguous failure (NFR-REL-3).
  return err.category === "server" || err.category === "timeout" || err.category === "transport";
}

/** Full-jitter exponential backoff, capped, in milliseconds. */
export function backoffMs(
  attempt: number,
  base: number,
  max: number,
  random: () => number,
): number {
  const ceiling = Math.min(max, base * 2 ** attempt);
  return Math.floor(ceiling / 2 + random() * (ceiling / 2));
}

export async function withRetry<T>(
  retryClass: RetryClass,
  attempt: () => Promise<T>,
  opts: RetryOptions = {},
): Promise<T> {
  const maxRetries = opts.maxRetries ?? DEFAULTS.maxRetries;
  const base = opts.baseDelayMs ?? DEFAULTS.baseDelayMs;
  const max = opts.maxDelayMs ?? DEFAULTS.maxDelayMs;
  const sleep = opts.sleep ?? defaultSleep;
  const random = opts.random ?? Math.random;

  let attemptNo = 0;
  for (;;) {
    try {
      return await attempt();
    } catch (e) {
      const err = e instanceof GraphError ? e : undefined;
      if (!err || attemptNo >= maxRetries || !shouldRetry(err, retryClass)) {
        throw e;
      }
      const delay = err.retryAfterMs ?? backoffMs(attemptNo, base, max, random);
      await sleep(delay);
      attemptNo++;
    }
  }
}
