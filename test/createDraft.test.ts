import { describe, expect, it, vi } from "vitest";
import { createDraft } from "../src/capabilities/createDraft.js";
import type { Account } from "../src/domain/types.js";
import type {
  AccountRegistry,
  AttachmentReader,
  GraphClient,
  GraphRequest,
} from "../src/domain/contracts.js";

const account: Account = { id: "a@x.com", displayId: "a@x.com", credentialSourceId: "app1" };
const registry: AccountRegistry = { list: async () => [account], resolve: async () => account };

const attachments: AttachmentReader = {
  read: async (input) => ({
    filename: input.filename ?? "file.bin",
    mimeType: input.mimeType ?? "application/octet-stream",
    bytes: new Uint8Array(Buffer.from(input.contentBase64 ?? "", "base64")),
  }),
};

function graphQueue(responses: unknown[]) {
  const calls: GraphRequest[] = [];
  let i = 0;
  const request = vi.fn(async (_acct: Account, req: GraphRequest) => {
    calls.push(req);
    return responses[i++];
  });
  return { graph: { request } as unknown as GraphClient, calls };
}

describe("createDraft (C4)", () => {
  it("creates a draft via POST /me/messages without sending (FR-C4-1/2)", async () => {
    const { graph, calls } = graphQueue([
      { id: "draft1", conversationId: "c9", webLink: "https://outlook/draft1" },
    ]);

    const result = await createDraft(
      { registry, graph, attachments },
      { to: ["b@x.com"], cc: ["c@x.com"], subject: "Hello", body: "hi there" },
    );

    expect(calls).toHaveLength(1);
    const req = calls[0]!;
    expect(req.method).toBe("POST");
    expect(req.path).toBe("/me/messages");
    expect(req.retryClass).toBe("nonDuplicable"); // NFR-REL-3
    const body = req.body as { toRecipients: unknown[]; subject: string };
    expect(body.subject).toBe("Hello");
    expect(body.toRecipients).toEqual([{ emailAddress: { address: "b@x.com" } }]);

    expect(result.structured.draft_id).toBe("draft1");
    expect(result.structured.conversation_id).toBe("c9");
    expect(result.structured.web_link).toBe("https://outlook/draft1");
    expect(result.structured.is_reply).toBe(false);
    expect(result.structured.has_attachments).toBe(false);
    expect(result.summary).toMatch(/not sent/i);
  });

  it("attaches resolved files (FR-C4-3)", async () => {
    const spy = vi.spyOn(attachments, "read");
    const { graph, calls } = graphQueue([{ id: "draft2" }]);

    const result = await createDraft(
      { registry, graph, attachments },
      {
        to: ["b@x.com"],
        body: "see attached",
        attachments: [
          { filename: "note.txt", contentBase64: Buffer.from("hi").toString("base64") },
        ],
      },
    );

    expect(spy).toHaveBeenCalledOnce();
    const body = calls[0]!.body as { attachments: Array<{ name: string; "@odata.type": string }> };
    expect(body.attachments[0]!.name).toBe("note.txt");
    expect(body.attachments[0]!["@odata.type"]).toBe("#microsoft.graph.fileAttachment");
    expect(result.structured.has_attachments).toBe(true);
    spy.mockRestore();
  });

  it("derives reply threading + Re: subject from the conversation (FR-C4-4)", async () => {
    const { graph, calls } = graphQueue([
      { value: [{ id: "m0", subject: "Project plan", internetMessageId: "<orig@mail>" }] },
      { id: "draft3", conversationId: "conv-1" },
    ]);

    const result = await createDraft(
      { registry, graph, attachments },
      { to: ["b@x.com"], body: "my reply", replyToConversationId: "conv-1" },
    );

    // First a lookup GET, then the draft POST.
    expect(calls[0]!.method).toBe("GET");
    expect(calls[0]!.query?.$filter).toBe("conversationId eq 'conv-1'");
    const body = calls[1]!.body as {
      subject: string;
      internetMessageHeaders: Array<{ name: string; value: string }>;
    };
    expect(body.subject).toBe("Re: Project plan");
    expect(body.internetMessageHeaders).toContainEqual({
      name: "In-Reply-To",
      value: "<orig@mail>",
    });
    expect(result.structured.is_reply).toBe(true);
  });
});
