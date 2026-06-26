import { describe, expect, it } from "vitest";
import { BoundedConcurrency } from "../src/util/bounded.js";

describe("BoundedConcurrency (NFR-REL-4)", () => {
  it("preserves result order regardless of completion order", async () => {
    const limiter = new BoundedConcurrency(3);
    const tasks = [30, 10, 20, 5].map((ms, i) => async () => {
      await new Promise((r) => setTimeout(r, ms));
      return i;
    });
    expect(await limiter.run(tasks)).toEqual([0, 1, 2, 3]);
  });

  it("never exceeds the configured concurrency", async () => {
    const limiter = new BoundedConcurrency(2);
    let inFlight = 0;
    let peak = 0;
    const tasks = Array.from({ length: 8 }, () => async () => {
      inFlight += 1;
      peak = Math.max(peak, inFlight);
      await new Promise((r) => setTimeout(r, 5));
      inFlight -= 1;
      return null;
    });
    await limiter.run(tasks);
    expect(peak).toBeLessThanOrEqual(2);
  });

  it("handles an empty task list", async () => {
    const limiter = new BoundedConcurrency(4);
    expect(await limiter.run([])).toEqual([]);
  });

  it("rejects an invalid limit", () => {
    expect(() => new BoundedConcurrency(0)).toThrow(/positive integer/i);
  });

  it("propagates the first task rejection", async () => {
    const limiter = new BoundedConcurrency(2);
    const tasks = [
      async () => 1,
      async () => {
        throw new Error("boom");
      },
    ];
    await expect(limiter.run(tasks)).rejects.toThrow(/boom/);
  });

  it("stops dispatching new tasks after the first failure", async () => {
    const limiter = new BoundedConcurrency(1); // serial: failure happens before later tasks
    let ran = 0;
    const tasks = [
      async () => {
        ran += 1;
        throw new Error("stop");
      },
      async () => {
        ran += 1;
        return 0;
      },
      async () => {
        ran += 1;
        return 0;
      },
    ];
    await expect(limiter.run(tasks)).rejects.toThrow(/stop/);
    expect(ran).toBe(1); // the two tasks after the failure never start
  });
});
