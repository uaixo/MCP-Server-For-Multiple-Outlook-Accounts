import { describe, expect, it, vi } from "vitest";
import { FetchGraphClient, GRAPH_BASE } from "../src/graph/client.js";
import { GraphError } from "../src/graph/errors.js";
import type { Account } from "../src/domain/types.js";

const account: Account = { id: "a@x.com", displayId: "a@x.com", credentialSourceId: "app1" };

function client(fetchImpl: typeof fetch) {
  return new FetchGraphClient({
    requestTimeoutMs: 30_000,
    getToken: async () => "tok-123",
    fetchImpl,
    retry: { maxRetries: 0, sleep: async () => undefined },
  });
}

describe("FetchGraphClient (NFR-REL-1, FR-ERR-1)", () => {
  it("builds the URL + auth header and returns parsed JSON", async () => {
    const fetchImpl = vi.fn(
      async () =>
        new Response(JSON.stringify({ value: [1, 2] }), {
          status: 200,
          headers: { "content-type": "application/json" },
        }),
    ) as unknown as typeof fetch;

    const c = client(fetchImpl);
    const out = await c.request<{ value: number[] }>(account, {
      method: "GET",
      path: "/me/messages",
      query: { $top: 5, $select: "id" },
      retryClass: "safe",
    });
    expect(out.value).toEqual([1, 2]);

    const [url, init] = (fetchImpl as unknown as ReturnType<typeof vi.fn>).mock.calls[0];
    const parsed = new URL(url as string);
    expect(parsed.origin + parsed.pathname).toBe(`${GRAPH_BASE}/me/messages`);
    expect(parsed.searchParams.get("$top")).toBe("5");
    expect(parsed.searchParams.get("$select")).toBe("id");
    expect((init as RequestInit).headers).toMatchObject({ Authorization: "Bearer tok-123" });
    expect((init as RequestInit).signal).toBeInstanceOf(AbortSignal);
  });

  it("uses an absolute nextLink path as-is", async () => {
    const next = `${GRAPH_BASE}/me/messages?$skiptoken=abc`;
    const fetchImpl = vi.fn(
      async () => new Response(JSON.stringify({ value: [] }), { status: 200 }),
    );
    const c = client(fetchImpl as unknown as typeof fetch);
    await c.request(account, { method: "GET", path: next, retryClass: "safe" });
    expect((fetchImpl as ReturnType<typeof vi.fn>).mock.calls[0][0]).toBe(next);
  });

  it("maps a non-OK response to a GraphError", async () => {
    const fetchImpl = vi.fn(
      async () => new Response(JSON.stringify({ error: { message: "nope" } }), { status: 401 }),
    );
    const c = client(fetchImpl as unknown as typeof fetch);
    await expect(
      c.request(account, { method: "GET", path: "/me", retryClass: "safe" }),
    ).rejects.toMatchObject({ name: "GraphError", category: "auth" });
  });

  it("returns undefined for 204 No Content", async () => {
    const fetchImpl = vi.fn(async () => new Response(null, { status: 204 }));
    const c = client(fetchImpl as unknown as typeof fetch);
    const out = await c.request(account, {
      method: "PATCH",
      path: "/me/messages/1",
      body: { isRead: true },
      retryClass: "safe",
    });
    expect(out).toBeUndefined();
  });

  it("refuses an absolute URL on a non-Graph host without fetching or minting a token (SSRF guard)", async () => {
    const fetchImpl = vi.fn(async () => new Response("{}", { status: 200 }));
    const getToken = vi.fn(async () => "tok-secret");
    const c = new FetchGraphClient({
      requestTimeoutMs: 30_000,
      getToken,
      fetchImpl: fetchImpl as unknown as typeof fetch,
      retry: { maxRetries: 3, sleep: async () => undefined },
    });

    await expect(
      c.request(account, {
        method: "GET",
        path: "https://evil.example/collect",
        retryClass: "safe",
      }),
    ).rejects.toThrow(/non-Microsoft-Graph host/i);

    // The token is never minted and no request is ever sent to the attacker host.
    expect(getToken).not.toHaveBeenCalled();
    expect(fetchImpl).not.toHaveBeenCalled();
  });

  it("still accepts an absolute nextLink on the Graph origin", async () => {
    const fetchImpl = vi.fn(
      async () => new Response(JSON.stringify({ value: [] }), { status: 200 }),
    );
    const c = client(fetchImpl as unknown as typeof fetch);
    await c.request(account, {
      method: "GET",
      path: `${GRAPH_BASE}/me/messages?$skiptoken=abc`,
      retryClass: "safe",
    });
    expect((fetchImpl as ReturnType<typeof vi.fn>).mock.calls).toHaveLength(1);
  });

  it("maps a 2xx body that isn't JSON to a GraphError instead of throwing raw", async () => {
    const fetchImpl = vi.fn(async () => new Response("<html>not json</html>", { status: 200 }));
    const c = client(fetchImpl as unknown as typeof fetch);
    await expect(
      c.request(account, { method: "GET", path: "/me", retryClass: "safe" }),
    ).rejects.toMatchObject({ name: "GraphError" });
  });

  it("returns undefined for 202 Accepted with an empty body (sendMail)", async () => {
    const fetchImpl = vi.fn(async () => new Response("", { status: 202 }));
    const c = client(fetchImpl as unknown as typeof fetch);
    const out = await c.request(account, {
      method: "POST",
      path: "/me/sendMail",
      body: { message: {}, saveToSentItems: true },
      retryClass: "nonDuplicable",
    });
    expect(out).toBeUndefined();
  });

  it("maps a thrown timeout to a GraphError(timeout)", async () => {
    const fetchImpl = vi.fn(async () => {
      throw Object.assign(new Error("aborted"), { name: "TimeoutError" });
    });
    const c = client(fetchImpl as unknown as typeof fetch);
    await expect(
      c.request(account, { method: "GET", path: "/me", retryClass: "safe" }),
    ).rejects.toMatchObject({ name: "GraphError", category: "timeout" });
  });
});
