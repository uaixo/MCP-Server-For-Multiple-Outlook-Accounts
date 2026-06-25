import { describe, expect, it } from "vitest";
import { DEFAULTS, ENV_KEYS, loadConfig } from "../src/config.js";

/**
 * Scaffold smoke tests for the config loader (spec §12 / provider-mapping §5).
 * These prove the test harness runs and the env→config mapping is correct.
 * Capability-level tests (mocked Graph/MSAL) are added in the build phase.
 */
describe("loadConfig", () => {
  it("applies reference defaults when env is empty", () => {
    const cfg = loadConfig({}, ":");
    expect(cfg.lockTimeoutMs).toBe(DEFAULTS.lockTimeoutMs);
    expect(cfg.requestTimeoutMs).toBe(DEFAULTS.requestTimeoutMs);
    expect(cfg.oauthCredentialsPath).toBeUndefined();
    // Path attachments disabled by default (NFR-SEC-3).
    expect(cfg.attachmentsAllowList).toEqual([]);
    expect(cfg.dataDir).toMatch(/\.outlook-mcp$/);
  });

  it("reads and validates overrides", () => {
    const cfg = loadConfig(
      {
        [ENV_KEYS.dataDir]: "/data/outlook",
        [ENV_KEYS.oauthCredentials]: "/data/outlook/app.json",
        [ENV_KEYS.attachmentsDir]: "/a:/b",
        [ENV_KEYS.lockTimeoutMs]: "5000",
        [ENV_KEYS.requestTimeoutMs]: "15000",
      },
      ":",
    );
    expect(cfg.dataDir).toBe("/data/outlook");
    expect(cfg.oauthCredentialsPath).toBe("/data/outlook/app.json");
    expect(cfg.attachmentsAllowList).toEqual(["/a", "/b"]);
    expect(cfg.lockTimeoutMs).toBe(5000);
    expect(cfg.requestTimeoutMs).toBe(15000);
  });

  it("rejects a non-positive integer timeout", () => {
    expect(() => loadConfig({ [ENV_KEYS.requestTimeoutMs]: "0" }, ":")).toThrow(
      ENV_KEYS.requestTimeoutMs,
    );
    expect(() => loadConfig({ [ENV_KEYS.lockTimeoutMs]: "abc" }, ":")).toThrow(
      ENV_KEYS.lockTimeoutMs,
    );
  });
});
