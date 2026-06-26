import { describe, expect, it, vi } from "vitest";
import { readConversation } from "../src/capabilities/readConversation.js";
import { CONVERSATION_BODY_CHAR_CAP } from "../src/output/contract.js";
import type { Account } from "../src/domain/types.js";
import type { AccountRegistry, GraphClient, GraphRequest } from "../src/domain/contracts.js";

const account: Account = { id: "a@x.com", displayId: "a@x.com", credentialSourceId: "app1" };
const registry: AccountRegistry = { list: async () => [account], resolve: async () => account };

function graphReturning(response: unknown): { graph: GraphClient; lastReq: () => GraphRequest } {
  const request = vi.fn(async () => response);
  return {
    graph: { request } as unknown as GraphClient,
    lastReq: () => request.mock.calls.at(-1)![1] as GraphRequest,
  };
}

describe("readConversation (C3)", () => {
  it("requires a conversation id", async () => {
    const { graph } = graphReturning({ value: [] });
    await expect(readConversation({ registry, graph }, { conversationId: " " })).rejects.toThrow(
      /conversation_id is required/i,
    );
  });

  it("maps messages, renders HTML bodies to text, and derives labels", async () => {
    const { graph, lastReq } = graphReturning({
      value: [
        {
          id: "m1",
          conversationId: "c1",
          subject: "Re: Hi",
          from: { emailAddress: { name: "Bob", address: "bob@x.com" } },
          toRecipients: [{ emailAddress: { address: "me@x.com" } }],
          receivedDateTime: "2026-06-02T10:00:00Z",
          body: { contentType: "html", content: "<p>Hello <b>there</b></p>" },
          categories: ["Work"],
          isRead: false,
        },
      ],
      "@odata.count": 1,
    });

    const result = await readConversation({ registry, graph }, { conversationId: "c1" });
    expect(result.structured.message_count).toBe(1);
    const m = result.structured.messages[0]!;
    expect(m.from).toBe("Bob <bob@x.com>");
    expect(m.to).toEqual(["me@x.com"]);
    expect(m.body_text).toBe("Hello there");
    expect(m.labels).toEqual(["Work", "Unread"]);
    expect(result.structured.truncated).toBe(false);

    // Filter on conversationId, newest first, with $count consistency header.
    const req = lastReq();
    expect(req.query?.$filter).toBe("conversationId eq 'c1'");
    expect(req.query?.$orderby).toBe("receivedDateTime desc");
    expect(req.headers).toMatchObject({ ConsistencyLevel: "eventual" });
  });

  it("flags truncation and omitted count when more messages exist (FR-C3-2)", async () => {
    const { graph } = graphReturning({
      value: [{ id: "m1", conversationId: "c1", body: { contentType: "text", content: "hi" } }],
      "@odata.count": 5,
    });
    const result = await readConversation({ registry, graph }, { conversationId: "c1" });
    expect(result.structured.truncated).toBe(true);
    expect(result.structured.omitted_message_count).toBe(4);
  });

  it("bounds combined body characters (NFR-PERF-2)", async () => {
    const big = "x".repeat(15_000);
    const { graph } = graphReturning({
      value: [
        { id: "m1", conversationId: "c1", body: { contentType: "text", content: big } },
        { id: "m2", conversationId: "c1", body: { contentType: "text", content: big } },
      ],
      "@odata.count": 2,
    });
    const result = await readConversation({ registry, graph }, { conversationId: "c1" });
    const total = result.structured.messages.reduce((n, m) => n + m.body_text.length, 0);
    expect(total).toBeLessThanOrEqual(CONVERSATION_BODY_CHAR_CAP);
    expect(result.structured.truncated).toBe(true);
  });
});
