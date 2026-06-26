import { describe, expect, it, vi } from "vitest";
import { organizeMail } from "../src/capabilities/organizeMail.js";
import { BoundedConcurrency } from "../src/util/bounded.js";
import type { Account } from "../src/domain/types.js";
import type { AccountRegistry, GraphClient, GraphRequest } from "../src/domain/contracts.js";

const account: Account = { id: "a@x.com", displayId: "a@x.com", credentialSourceId: "app1" };
const registry: AccountRegistry = { list: async () => [account], resolve: async () => account };
const limiter = new BoundedConcurrency(5);

/** Graph mock: each handler may answer a request by returning a value (else undefined). */
function makeGraph(handlers: Array<(req: GraphRequest) => unknown>) {
  const calls: GraphRequest[] = [];
  const request = vi.fn(async (_acct: Account, req: GraphRequest) => {
    calls.push(req);
    for (const h of handlers) {
      const r = h(req);
      if (r !== undefined) return r;
    }
    return undefined; // PATCH / move success (no body)
  });
  return { graph: { request } as unknown as GraphClient, calls };
}

const deps = (graph: GraphClient) => ({ registry, graph, limiter });

describe("organizeMail (C8) — validation (FR-C8-1/2)", () => {
  const { graph } = makeGraph([]);

  it("requires exactly one target", async () => {
    await expect(organizeMail(deps(graph), { markRead: true })).rejects.toThrow(/exactly one/i);
    await expect(
      organizeMail(deps(graph), { conversationId: "c", messageId: "m", markRead: true }),
    ).rejects.toThrow(/exactly one/i);
  });

  it("requires at least one change", async () => {
    await expect(organizeMail(deps(graph), { messageId: "m1" })).rejects.toThrow(
      /no changes requested/i,
    );
  });
});

describe("organizeMail (C8) — single message", () => {
  it("merges categories and read-state into one PATCH (FR-C8-6)", async () => {
    const { graph, calls } = makeGraph([
      (req) => (req.method === "GET" ? { id: "m1", categories: ["Old"], isRead: true } : undefined),
    ]);

    const result = await organizeMail(deps(graph), {
      messageId: "m1",
      addLabels: ["Work"],
      removeLabels: ["Old"],
      markRead: false,
    });

    const patch = calls.find((c) => c.method === "PATCH")!;
    expect(patch.path).toBe("/me/messages/m1");
    expect(patch.body).toEqual({ categories: ["Work"], isRead: false });

    expect(result.structured.message_count).toBe(1);
    expect(result.structured.labels).toEqual(["Work"]);
    expect(result.structured.marked_read).toBe(false);
  });

  it("emits a move op for archive", async () => {
    const { graph, calls } = makeGraph([
      (req) => (req.method === "GET" ? { id: "m1", categories: [], isRead: true } : undefined),
    ]);

    const result = await organizeMail(deps(graph), { messageId: "m1", archive: true });

    const move = calls.find((c) => c.path === "/me/messages/m1/move")!;
    expect(move.method).toBe("POST");
    expect(move.body).toEqual({ destinationId: "archive" });
    expect(result.structured.archived).toBe(true);
  });
});

describe("organizeMail (C8) — whole conversation (FR-C8-4)", () => {
  it("applies per message and reports the union of resulting labels", async () => {
    const { graph, calls } = makeGraph([
      (req) =>
        req.method === "GET" && req.path === "/me/messages"
          ? {
              value: [
                { id: "m1", categories: ["A"], isRead: false },
                { id: "m2", categories: ["B"], isRead: false },
              ],
            }
          : undefined,
    ]);

    const result = await organizeMail(deps(graph), {
      conversationId: "conv-1",
      addLabels: ["Shared"],
      markRead: true,
    });

    // One enumeration GET, then a PATCH per message.
    const patches = calls.filter((c) => c.method === "PATCH");
    expect(patches).toHaveLength(2);
    expect(patches.map((p) => p.path).sort()).toEqual(["/me/messages/m1", "/me/messages/m2"]);

    expect(result.structured.message_count).toBe(2);
    // Union across messages: A, B, plus the added Shared — sorted.
    expect(result.structured.labels).toEqual(["A", "B", "Shared"]);
    expect(result.structured.target).toEqual({ type: "conversation", id: "conv-1" });
  });

  it("errors when the conversation has no messages", async () => {
    const { graph } = makeGraph([
      (req) => (req.path === "/me/messages" ? { value: [] } : undefined),
    ]);
    await expect(
      organizeMail(deps(graph), { conversationId: "empty", markRead: true }),
    ).rejects.toThrow(/no messages found/i);
  });
});
