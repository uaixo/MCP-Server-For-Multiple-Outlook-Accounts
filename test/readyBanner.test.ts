import { describe, expect, it } from "vitest";
import { formatReadyBanner } from "../src/index.js";

describe("formatReadyBanner (NFR-OPS-2 / NFR-SEC-6)", () => {
  it("lists connected account identities", () => {
    expect(formatReadyBanner(["a@x.com", "b@y.com"])).toBe(
      "[outlook-mcp] ready on stdio — connected accounts: a@x.com, b@y.com\n",
    );
  });

  it("says 'none' when no accounts are connected", () => {
    expect(formatReadyBanner([])).toBe("[outlook-mcp] ready on stdio — connected accounts: none\n");
  });
});
