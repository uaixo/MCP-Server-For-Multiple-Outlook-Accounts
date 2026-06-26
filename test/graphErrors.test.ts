import { describe, expect, it } from "vitest";
import {
  GraphError,
  categorizeStatus,
  errorFromResponse,
  errorFromThrown,
  parseRetryAfter,
} from "../src/graph/errors.js";

describe("Graph error mapping (FR-ERR-1)", () => {
  it("categorizes HTTP status codes", () => {
    expect(categorizeStatus(401)).toBe("auth");
    expect(categorizeStatus(403)).toBe("auth");
    expect(categorizeStatus(429)).toBe("rateLimited");
    expect(categorizeStatus(408)).toBe("timeout");
    expect(categorizeStatus(503)).toBe("server");
    expect(categorizeStatus(400)).toBe("client");
    expect(categorizeStatus(200)).toBe("unknown");
  });

  it("parses Retry-After as delta-seconds and as an HTTP date", () => {
    expect(parseRetryAfter("5")).toBe(5000);
    expect(parseRetryAfter(null)).toBeUndefined();
    const now = Date.now();
    const future = new Date(now + 2000).toUTCString();
    const ms = parseRetryAfter(future, now);
    expect(ms).toBeGreaterThanOrEqual(0);
    expect(ms).toBeLessThanOrEqual(2000);
  });

  it("maps a 403 response to an actionable auth error", async () => {
    const res = new Response(JSON.stringify({ error: { message: "Forbidden" } }), { status: 403 });
    const err = await errorFromResponse(res);
    expect(err).toBeInstanceOf(GraphError);
    expect(err.category).toBe("auth");
    expect(err.message).toMatch(/re-connect/i);
  });

  it("maps a 429 response with Retry-After to a rate-limit error", async () => {
    const res = new Response("{}", { status: 429, headers: { "retry-after": "3" } });
    const err = await errorFromResponse(res);
    expect(err.category).toBe("rateLimited");
    expect(err.retryAfterMs).toBe(3000);
  });

  it("includes the Graph message for a 400 client error", async () => {
    const res = new Response(JSON.stringify({ error: { message: "Bad filter" } }), { status: 400 });
    const err = await errorFromResponse(res);
    expect(err.category).toBe("client");
    expect(err.message).toMatch(/Bad filter/);
  });

  it("maps thrown transport/timeout errors", () => {
    expect(errorFromThrown(Object.assign(new Error("x"), { name: "TimeoutError" })).category).toBe(
      "timeout",
    );
    expect(errorFromThrown(Object.assign(new Error("x"), { name: "AbortError" })).category).toBe(
      "timeout",
    );
    expect(errorFromThrown(new TypeError("fetch failed")).category).toBe("transport");
    const passthrough = new GraphError("x", "server", 500);
    expect(errorFromThrown(passthrough)).toBe(passthrough);
  });
});
