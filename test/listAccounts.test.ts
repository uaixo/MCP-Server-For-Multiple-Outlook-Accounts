import { describe, expect, it } from "vitest";
import { listAccounts } from "../src/capabilities/listAccounts.js";
import type { Account } from "../src/domain/types.js";
import type { AccountRegistry } from "../src/domain/contracts.js";

function fakeRegistry(accounts: Account[]): AccountRegistry {
  return {
    list: async () => accounts,
    resolve: async () => accounts[0]!,
  };
}

const acct = (id: string, source = "app1"): Account => ({
  id: id.toLowerCase(),
  displayId: id,
  credentialSourceId: source,
});

describe("listAccounts (C1, FR-C1-1..3)", () => {
  it("returns a non-error connect hint when none are connected (FR-C1-2)", async () => {
    const result = await listAccounts(fakeRegistry([]));
    expect(result.structured.account_count).toBe(0);
    expect(result.structured.accounts).toEqual([]);
    expect(result.summary).toMatch(/connect/i);
  });

  it("lists a single connected account with its credential source", async () => {
    const result = await listAccounts(fakeRegistry([acct("User@Example.com", "acme")]));
    expect(result.structured.account_count).toBe(1);
    expect(result.structured.accounts[0]).toEqual({
      account: "User@Example.com",
      credential_source: "acme",
    });
    expect(result.summary).toContain("User@Example.com");
    expect(result.summary).toContain("acme");
  });

  it("lists several accounts", async () => {
    const result = await listAccounts(fakeRegistry([acct("a@x.com"), acct("b@y.com")]));
    expect(result.structured.account_count).toBe(2);
    expect(result.structured.accounts.map((a) => a.account)).toEqual(["a@x.com", "b@y.com"]);
  });
});
