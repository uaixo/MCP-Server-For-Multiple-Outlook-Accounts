import { afterEach, describe, expect, it, vi } from "vitest";
import type { CredentialSource } from "../src/domain/contracts.js";

// Shared mock fns, hoisted so the vi.mock factory can reference them.
const mocks = vi.hoisted(() => ({
  acquireTokenInteractive: vi.fn(),
  acquireTokenSilent: vi.fn(),
  serialize: vi.fn(() => '{"cache":1}'),
  deserialize: vi.fn(),
  getAllAccounts: vi.fn(async () => [] as Array<{ username: string }>),
}));

vi.mock("@azure/msal-node", () => ({
  // A class (not vi.fn) so it is usable with `new` under vitest 4.
  PublicClientApplication: class {
    acquireTokenInteractive = mocks.acquireTokenInteractive;
    acquireTokenSilent = mocks.acquireTokenSilent;
    getTokenCache() {
      return {
        serialize: mocks.serialize,
        deserialize: mocks.deserialize,
        getAllAccounts: mocks.getAllAccounts,
      };
    }
  },
}));

const { interactiveConnect, acquireToken } = await import("../src/auth/msalClient.js");

const source: CredentialSource = {
  id: "app1",
  clientId: "client-123",
  tenant: "common",
  scopes: ["Mail.ReadWrite", "Mail.Send", "User.Read", "offline_access"],
};

const noopBrowser = async (): Promise<void> => undefined;

afterEach(() => {
  vi.clearAllMocks();
  mocks.serialize.mockReturnValue('{"cache":1}');
});

describe("interactiveConnect (FR-AUTH-3..7)", () => {
  it("returns the authenticated identity and serialized cache, requesting only resource scopes", async () => {
    mocks.acquireTokenInteractive.mockResolvedValue({
      account: { username: "User@Example.com" },
      accessToken: "tok",
    });

    const result = await interactiveConnect(source, { openBrowser: noopBrowser });
    expect(result.identity).toBe("User@Example.com");
    expect(result.serializedCache).toBe('{"cache":1}');

    const req = mocks.acquireTokenInteractive.mock.calls[0]?.[0];
    expect(req.scopes).toEqual(["Mail.ReadWrite", "Mail.Send", "User.Read"]); // offline_access stripped
  });

  it("aborts when the authenticated account cannot be determined (FR-AUTH-6)", async () => {
    mocks.acquireTokenInteractive.mockResolvedValue({ account: null, accessToken: "tok" });
    await expect(interactiveConnect(source, { openBrowser: noopBrowser })).rejects.toThrow(
      /determine the authenticated account/i,
    );
  });

  it("bounds the consent wait and fails on timeout (FR-AUTH-7)", async () => {
    mocks.acquireTokenInteractive.mockReturnValue(new Promise(() => undefined)); // never resolves
    await expect(
      interactiveConnect(source, { openBrowser: noopBrowser, timeoutMs: 20 }),
    ).rejects.toThrow(/Timed out/i);
  });
});

describe("acquireToken (FR-ID-5 silent refresh)", () => {
  it("refreshes via the issuing app registration and returns the access token", async () => {
    mocks.getAllAccounts.mockResolvedValue([{ username: "User@Example.com" }]);
    mocks.acquireTokenSilent.mockResolvedValue({ accessToken: "fresh-token" });

    const result = await acquireToken(
      { id: "user@example.com", displayId: "User@Example.com", credentialSourceId: "app1" },
      source,
      '{"cache":1}',
      { timeoutMs: 1000 },
    );
    expect(result.accessToken).toBe("fresh-token");
    expect(mocks.deserialize).toHaveBeenCalledWith('{"cache":1}');
  });

  it("throws an actionable re-connect error when the account is not in the cache", async () => {
    mocks.getAllAccounts.mockResolvedValue([]);
    await expect(
      acquireToken(
        { id: "user@example.com", displayId: "User@Example.com", credentialSourceId: "app1" },
        source,
        "{}",
        { timeoutMs: 1000 },
      ),
    ).rejects.toThrow(/re-connect/i);
  });
});
