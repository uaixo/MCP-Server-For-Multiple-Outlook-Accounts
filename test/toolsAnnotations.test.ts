import { describe, expect, it } from "vitest";
import { Client } from "@modelcontextprotocol/sdk/client/index.js";
import { InMemoryTransport } from "@modelcontextprotocol/sdk/inMemory.js";
import { createServer } from "../src/index.js";
import { BoundedConcurrency } from "../src/util/bounded.js";
import type { Account } from "../src/domain/types.js";
import type { AccountRegistry, AttachmentReader, GraphClient } from "../src/domain/contracts.js";

const account: Account = { id: "a@x.com", displayId: "a@x.com", credentialSourceId: "app1" };
const registry: AccountRegistry = { list: async () => [account], resolve: async () => account };
const graph = { request: async () => ({}) } as unknown as GraphClient;
const attachments = {
  read: async () => ({
    filename: "f",
    mimeType: "application/octet-stream",
    bytes: new Uint8Array(),
  }),
} as AttachmentReader;
const limiter = new BoundedConcurrency(5);

async function listTools() {
  const server = createServer({ registry, graph, attachments, limiter });
  const [clientTransport, serverTransport] = InMemoryTransport.createLinkedPair();
  const client = new Client({ name: "test", version: "0" });
  await Promise.all([client.connect(clientTransport), server.connect(serverTransport)]);
  const { tools } = await client.listTools();
  await client.close();
  return new Map(tools.map((t) => [t.name, t]));
}

describe("tool safety annotations (NFR-OPS-4)", () => {
  it("registers all phase 1–4 tools", async () => {
    const tools = await listTools();
    for (const name of [
      "list_accounts",
      "search_conversations",
      "read_conversation",
      "create_draft",
      "send_message",
      "list_labels",
      "create_label",
      "organize_mail",
    ]) {
      expect(tools.has(name)).toBe(true);
    }
  });

  it("marks send_message destructive and create_draft non-destructive (FR-C5-3)", async () => {
    const tools = await listTools();

    const send = tools.get("send_message")!;
    expect(send.annotations?.destructiveHint).toBe(true);
    expect(send.annotations?.readOnlyHint).toBe(false);
    expect(send.annotations?.idempotentHint).toBe(false);

    const draft = tools.get("create_draft")!;
    expect(draft.annotations?.destructiveHint).toBe(false);
    expect(draft.annotations?.readOnlyHint).toBe(false);

    // Read tools remain read-only and non-destructive.
    expect(tools.get("read_conversation")!.annotations?.readOnlyHint).toBe(true);
  });

  it("marks organize_mail destructive + idempotent and list_labels read-only (FR-C8-5)", async () => {
    const tools = await listTools();

    const organise = tools.get("organize_mail")!;
    expect(organise.annotations?.destructiveHint).toBe(true);
    expect(organise.annotations?.idempotentHint).toBe(true);
    expect(organise.annotations?.readOnlyHint).toBe(false);

    expect(tools.get("list_labels")!.annotations?.readOnlyHint).toBe(true);

    // create_label is an additive write, not destructive.
    const createLabel = tools.get("create_label")!;
    expect(createLabel.annotations?.readOnlyHint).toBe(false);
    expect(createLabel.annotations?.destructiveHint).toBe(false);
  });
});
