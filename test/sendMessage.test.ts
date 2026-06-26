import { describe, expect, it, vi } from "vitest";
import { sendMessage } from "../src/capabilities/sendMessage.js";
import { FetchGraphClient } from "../src/graph/client.js";
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

describe("sendMessage (C5)", () => {
  it("sends via a single POST /me/sendMail (FR-C5-1/2)", async () => {
    const { graph, calls } = graphQueue([undefined]); // sendMail → 202, no body

    const result = await sendMessage(
      { registry, graph, attachments },
      { to: ["b@x.com"], subject: "Ping", body: "hi" },
    );

    expect(calls).toHaveLength(1);
    const req = calls[0]!;
    expect(req.method).toBe("POST");
    expect(req.path).toBe("/me/sendMail");
    expect(req.retryClass).toBe("nonDuplicable"); // NFR-REL-3 / FR-C5-4
    const body = req.body as { message: { subject: string }; saveToSentItems: boolean };
    expect(body.saveToSentItems).toBe(true);
    expect(body.message.subject).toBe("Ping");

    expect(result.structured.sent).toBe(true);
    expect(result.structured.to).toEqual(["b@x.com"]);
    expect(result.summary).toMatch(/^Sent/);
  });

  it("does NOT duplicate-send under a transient failure (FR-C5-4 / NFR-REL-3)", async () => {
    // Real client so the actual retry policy applies. A 5xx is an *ambiguous*
    // failure for a send: it must be attempted exactly once and then surfaced.
    const fetchImpl = vi.fn(async () => new Response("boom", { status: 503 }));
    const graph = new FetchGraphClient({
      requestTimeoutMs: 1000,
      getToken: async () => "tok",
      fetchImpl: fetchImpl as unknown as typeof fetch,
      retry: { baseDelayMs: 0, sleep: async () => undefined }, // would-be retries are instant
    });

    await expect(
      sendMessage({ registry, graph, attachments }, { to: ["b@x.com"], body: "hi" }),
    ).rejects.toThrow();

    expect(fetchImpl).toHaveBeenCalledOnce(); // attempted exactly once — no duplicate
  });

  it("retries a pre-processing 429 then succeeds (no duplicate, NFR-REL-3)", async () => {
    let n = 0;
    const fetchImpl = vi.fn(async () => {
      n += 1;
      if (n === 1) return new Response("slow down", { status: 429 });
      return new Response("", { status: 202 });
    });
    const graph = new FetchGraphClient({
      requestTimeoutMs: 1000,
      getToken: async () => "tok",
      fetchImpl: fetchImpl as unknown as typeof fetch,
      retry: { baseDelayMs: 0, sleep: async () => undefined },
    });

    const result = await sendMessage(
      { registry, graph, attachments },
      { to: ["b@x.com"], body: "hi" },
    );
    expect(result.structured.sent).toBe(true);
    expect(fetchImpl).toHaveBeenCalledTimes(2); // 429 retried, then delivered once
  });

  it("threads a reply via the conversation lookup (FR-C5-1)", async () => {
    const { graph, calls } = graphQueue([
      { value: [{ id: "m0", subject: "Plan", internetMessageId: "<x@mail>" }] },
      undefined,
    ]);

    await sendMessage(
      { registry, graph, attachments },
      { to: ["b@x.com"], body: "reply", replyToConversationId: "conv-9" },
    );

    expect(calls[0]!.method).toBe("GET");
    expect(calls[1]!.path).toBe("/me/sendMail");
    const body = calls[1]!.body as { message: { subject: string } };
    expect(body.message.subject).toBe("Re: Plan");
  });
});
