import { describe, expect, it, vi } from "vitest";
import { searchConversations } from "../src/capabilities/searchConversations.js";
import type { Account } from "../src/domain/types.js";
import type { AccountRegistry, GraphClient, GraphRequest } from "../src/domain/contracts.js";

const account: Account = { id: "a@x.com", displayId: "a@x.com", credentialSourceId: "app1" };

const registry: AccountRegistry = {
  list: async () => [account],
  resolve: async () => account,
};

function graphReturning(response: unknown): { graph: GraphClient; lastReq: () => GraphRequest } {
  const request = vi.fn(async () => response);
  return {
    graph: { request } as unknown as GraphClient,
    lastReq: () => request.mock.calls.at(-1)![1] as GraphRequest,
  };
}

describe("searchConversations (C2)", () => {
  it("requires a non-empty query (FR-C2-1)", async () => {
    const { graph } = graphReturning({ value: [] });
    await expect(searchConversations({ registry, graph }, { query: "  " })).rejects.toThrow(
      /query is required/i,
    );
  });

  it("dedupes messages into conversation summaries and returns the cursor", async () => {
    const { graph, lastReq } = graphReturning({
      value: [
        {
          id: "m1",
          conversationId: "c1",
          subject: "Hi",
          from: { emailAddress: { name: "Bob", address: "bob@x.com" } },
          receivedDateTime: "2026-06-01T10:00:00Z",
          bodyPreview: "hello",
        },
        { id: "m2", conversationId: "c1", subject: "Hi", bodyPreview: "older" },
        { id: "m3", conversationId: "c2", subject: "Other", bodyPreview: "snip" },
      ],
      "@odata.nextLink": "https://graph.microsoft.com/v1.0/me/messages?$skiptoken=Z",
    });

    const result = await searchConversations({ registry, graph }, { query: "hello", pageSize: 50 });
    expect(result.structured.count).toBe(2);
    expect(result.structured.conversations.map((c) => c.conversation_id)).toEqual(["c1", "c2"]);
    expect(result.structured.conversations[0]).toMatchObject({
      subject: "Hi",
      sender: "Bob <bob@x.com>",
      snippet: "hello",
    });
    expect(result.structured.next_page_token).toContain("$skiptoken=Z");

    // Free-text query → $search, no $orderby, $top honoured.
    const req = lastReq();
    expect(req.path).toBe("/me/messages");
    expect(req.query).toMatchObject({ $search: '"hello"', $top: 50 });
    expect(req.query?.$orderby).toBeUndefined();
  });

  it("builds an OData $filter (with $orderby) for structured queries", async () => {
    const { graph, lastReq } = graphReturning({ value: [] });
    await searchConversations({ registry, graph }, { query: "is:unread" });
    expect(lastReq().query).toMatchObject({
      $filter: "isRead eq false",
      $orderby: "receivedDateTime desc",
    });
  });

  it("fetches the opaque nextLink directly when given a page token", async () => {
    const { graph, lastReq } = graphReturning({ value: [] });
    const token = "https://graph.microsoft.com/v1.0/me/messages?$skiptoken=ABC";
    await searchConversations({ registry, graph }, { query: "hello", pageToken: token });
    expect(lastReq().path).toBe(token);
    expect(lastReq().query).toBeUndefined();
  });

  it("propagates an unsupported-operator translation error", async () => {
    const { graph } = graphReturning({ value: [] });
    await expect(searchConversations({ registry, graph }, { query: "foo:bar" })).rejects.toThrow(
      /Unsupported search operator/i,
    );
  });
});
