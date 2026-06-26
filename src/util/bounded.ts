/**
 * Bounded-concurrency runner (NFR-REL-4).
 *
 * Organise (C8) fans a single neutral request out to one Graph call per message
 * in a conversation. Firing them all at once would risk tripping Graph's rate
 * limits, so this caps how many run concurrently while preserving input order
 * in the results array.
 */

import type { ConcurrencyLimiter } from "../domain/contracts.js";

/** Default max in-flight Graph calls for a per-message fan-out. */
export const ORGANISE_CONCURRENCY = 5;

export class BoundedConcurrency implements ConcurrencyLimiter {
  private readonly limit: number;

  constructor(limit: number = ORGANISE_CONCURRENCY) {
    if (!Number.isInteger(limit) || limit < 1) {
      throw new Error(`Concurrency limit must be a positive integer (got ${limit}).`);
    }
    this.limit = limit;
  }

  async run<T>(tasks: Array<() => Promise<T>>): Promise<T[]> {
    const results = new Array<T>(tasks.length);
    let next = 0;
    let failed = false;

    const worker = async (): Promise<void> => {
      for (;;) {
        // Stop dispatching new tasks once any task has failed, so a partial
        // fan-out (e.g. organize_mail) doesn't keep applying side effects after
        // the first error. In-flight tasks still settle; the rejection wins.
        if (failed) return;
        const index = next++;
        if (index >= tasks.length) return;
        try {
          results[index] = await tasks[index]!();
        } catch (e) {
          failed = true;
          throw e;
        }
      }
    };

    const workerCount = Math.min(this.limit, tasks.length);
    await Promise.all(Array.from({ length: workerCount }, () => worker()));
    return results;
  }
}
