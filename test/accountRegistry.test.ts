import { describe, expect, it } from "vitest";
import { AccountSelectionError, FileAccountRegistry } from "../src/auth/accountRegistry.js";
import type { Account } from "../src/domain/types.js";
import type { TokenStore } from "../src/domain/contracts.js";

/** Minimal fake store returning a fixed account set. */
function fakeStore(accounts: Account[]): TokenStore {
  return {
    list: async () => accounts,
    readCache: async () => undefined,
    upsert: async () => undefined,
    remove: async () => undefined,
  };
}

const acct = (id: string): Account => ({ id, displayId: id, credentialSourceId: "app1" });

describe("FileAccountRegistry.resolve (FR-ID-2/3/4)", () => {
  it("errors with connect guidance when none are connected", async () => {
    const reg = new FileAccountRegistry(fakeStore([]));
    await expect(reg.resolve()).rejects.toBeInstanceOf(AccountSelectionError);
    await expect(reg.resolve()).rejects.toThrow(/connect/i);
  });

  it("defaults to the only connected account", async () => {
    const reg = new FileAccountRegistry(fakeStore([acct("a@x.com")]));
    expect((await reg.resolve()).id).toBe("a@x.com");
  });

  it("requires disambiguation when several are connected", async () => {
    const reg = new FileAccountRegistry(fakeStore([acct("a@x.com"), acct("b@y.com")]));
    await expect(reg.resolve()).rejects.toThrow(/specify one/i);
    await expect(reg.resolve()).rejects.toThrow(/a@x.com/);
  });

  it("matches a named account case-insensitively", async () => {
    const reg = new FileAccountRegistry(fakeStore([acct("a@x.com"), acct("b@y.com")]));
    expect((await reg.resolve("A@X.COM")).id).toBe("a@x.com");
  });

  it("errors listing accounts when a named account is unknown", async () => {
    const reg = new FileAccountRegistry(fakeStore([acct("a@x.com")]));
    await expect(reg.resolve("nope@z.com")).rejects.toThrow(/Unknown account/i);
    await expect(reg.resolve("nope@z.com")).rejects.toThrow(/a@x.com/);
  });
});
