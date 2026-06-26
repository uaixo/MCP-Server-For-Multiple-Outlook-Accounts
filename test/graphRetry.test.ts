import { describe, expect, it, vi } from "vitest";
import { GraphError } from "../src/graph/errors.js";
import { backoffMs, shouldRetry, withRetry } from "../src/graph/retry.js";

const noSleep = async (): Promise<void> => undefined;
const opts = { sleep: noSleep, random: () => 0, baseDelayMs: 10, maxDelayMs: 100 };

describe("retry policy (NFR-REL-2/3)", () => {
  it("shouldRetry: 429 retries for both classes", () => {
    const e = new GraphError("x", "rateLimited", 429);
    expect(shouldRetry(e, "safe")).toBe(true);
    expect(shouldRetry(e, "nonDuplicable")).toBe(true);
  });

  it("shouldRetry: ambiguous failures retry only for safe (no duplicate sends)", () => {
    for (const cat of ["server", "timeout", "transport"] as const) {
      expect(shouldRetry(new GraphError("x", cat), "safe")).toBe(true);
      expect(shouldRetry(new GraphError("x", cat), "nonDuplicable")).toBe(false);
    }
  });

  it("shouldRetry: client/auth never retry", () => {
    expect(shouldRetry(new GraphError("x", "client", 400), "safe")).toBe(false);
    expect(shouldRetry(new GraphError("x", "auth", 401), "safe")).toBe(false);
  });

  it("retries a safe transient failure then succeeds", async () => {
    const attempt = vi
      .fn()
      .mockRejectedValueOnce(new GraphError("x", "server", 503))
      .mockResolvedValueOnce("ok");
    await expect(withRetry("safe", attempt, opts)).resolves.toBe("ok");
    expect(attempt).toHaveBeenCalledTimes(2);
  });

  it("does NOT retry a nonDuplicable ambiguous failure", async () => {
    const attempt = vi.fn().mockRejectedValue(new GraphError("x", "server", 503));
    await expect(withRetry("nonDuplicable", attempt, opts)).rejects.toBeInstanceOf(GraphError);
    expect(attempt).toHaveBeenCalledTimes(1);
  });

  it("retries a nonDuplicable 429 (pre-processing throttle)", async () => {
    const attempt = vi
      .fn()
      .mockRejectedValueOnce(new GraphError("x", "rateLimited", 429))
      .mockResolvedValueOnce("sent");
    await expect(withRetry("nonDuplicable", attempt, opts)).resolves.toBe("sent");
    expect(attempt).toHaveBeenCalledTimes(2);
  });

  it("honours Retry-After for the backoff delay", async () => {
    const sleep = vi.fn(async () => undefined);
    const attempt = vi
      .fn()
      .mockRejectedValueOnce(new GraphError("x", "rateLimited", 429, 1234))
      .mockResolvedValueOnce("ok");
    await withRetry("safe", attempt, { ...opts, sleep });
    expect(sleep).toHaveBeenCalledWith(1234);
  });

  it("gives up after maxRetries", async () => {
    const attempt = vi.fn().mockRejectedValue(new GraphError("x", "server", 503));
    await expect(withRetry("safe", attempt, { ...opts, maxRetries: 2 })).rejects.toBeInstanceOf(
      GraphError,
    );
    expect(attempt).toHaveBeenCalledTimes(3); // initial + 2 retries
  });

  it("backoffMs stays within the cap", () => {
    for (let i = 0; i < 10; i++) {
      const d = backoffMs(i, 500, 8000, () => 1);
      expect(d).toBeLessThanOrEqual(8000);
      expect(d).toBeGreaterThanOrEqual(0);
    }
  });
});
