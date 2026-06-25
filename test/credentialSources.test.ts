import { mkdtemp, rm, writeFile } from "node:fs/promises";
import { tmpdir } from "node:os";
import { join } from "node:path";
import { afterEach, beforeEach, describe, expect, it } from "vitest";
import {
  OUTLOOK_SCOPES,
  consentScopes,
  getCredentialSource,
  loadCredentialSources,
} from "../src/auth/credentialSources.js";
import type { Config } from "../src/config.js";

function baseConfig(dataDir: string): Config {
  return {
    dataDir,
    oauthCredentialsPath: undefined,
    attachmentsAllowList: [],
    lockTimeoutMs: 12000,
    requestTimeoutMs: 30000,
  };
}

describe("credentialSources (FR-ID-5/6, FR-AUTH-10)", () => {
  let dir: string;
  beforeEach(async () => {
    dir = await mkdtemp(join(tmpdir(), "creds-"));
  });
  afterEach(async () => {
    await rm(dir, { recursive: true, force: true });
  });

  it("strips reserved OIDC scopes for the consent request", () => {
    expect(OUTLOOK_SCOPES).toContain("offline_access");
    expect(consentScopes(OUTLOOK_SCOPES)).toEqual(["Mail.ReadWrite", "Mail.Send", "User.Read"]);
  });

  it("pins a single source from OUTLOOK_OAUTH_CREDENTIALS and defaults the tenant", async () => {
    const path = join(dir, "acme.json");
    await writeFile(path, JSON.stringify({ clientId: "abc-123" }));
    const sources = await loadCredentialSources({ ...baseConfig(dir), oauthCredentialsPath: path });

    expect(sources).toHaveLength(1);
    expect(sources[0]).toMatchObject({ id: "acme", clientId: "abc-123", tenant: "common" });
    expect(sources[0]?.scopes).toEqual([...OUTLOOK_SCOPES]);
  });

  it("auto-discovers all credentials*.json sources, sorted", async () => {
    await writeFile(join(dir, "credentials.json"), JSON.stringify({ clientId: "c1" }));
    await writeFile(
      join(dir, "credentials-acme.json"),
      JSON.stringify({ clientId: "c2", tenant: "tenant-xyz", id: "acme" }),
    );
    await writeFile(join(dir, "tokens.json"), "{}"); // must be ignored

    // Sorted by filename: "credentials-acme.json" < "credentials.json" ('-' < '.').
    const sources = await loadCredentialSources(baseConfig(dir));
    expect(sources.map((s) => s.id)).toEqual(["acme", "credentials"]);
    expect(sources.find((s) => s.id === "acme")).toMatchObject({
      clientId: "c2",
      tenant: "tenant-xyz",
    });
  });

  it("rejects a config missing clientId with an actionable message", async () => {
    const path = join(dir, "credentials.json");
    await writeFile(path, JSON.stringify({ tenant: "common" }));
    await expect(
      loadCredentialSources({ ...baseConfig(dir), oauthCredentialsPath: path }),
    ).rejects.toThrow(/clientId/);
  });

  it("resolves a source by id", async () => {
    await writeFile(join(dir, "credentials.json"), JSON.stringify({ clientId: "c1" }));
    const found = await getCredentialSource(baseConfig(dir), "credentials");
    expect(found?.clientId).toBe("c1");
    expect(await getCredentialSource(baseConfig(dir), "missing")).toBeUndefined();
  });
});
