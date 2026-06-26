import { describe, expect, it, vi } from "vitest";
import { GraphAttachmentUploader } from "../src/mail/uploadSession.js";
import { UPLOAD_CHUNK_BYTES } from "../src/output/contract.js";
import type { Account } from "../src/domain/types.js";
import type { GraphClient, GraphRequest, ResolvedAttachment } from "../src/domain/contracts.js";

const account: Account = { id: "a@x.com", displayId: "a@x.com", credentialSourceId: "app1" };

const attachment = (bytes: number): ResolvedAttachment => ({
  filename: "big.bin",
  mimeType: "application/octet-stream",
  bytes: new Uint8Array(bytes),
});

function graphReturning(value: unknown) {
  const request = vi.fn(async (_a: Account, _req: GraphRequest) => value);
  return { graph: { request } as unknown as GraphClient, request };
}

describe("GraphAttachmentUploader", () => {
  it("creates a session, then PUTs chunks with Content-Range and no auth header", async () => {
    const { graph, request } = graphReturning({ uploadUrl: "https://upload.example/abc" });
    const puts: Array<{ url: string; init: RequestInit }> = [];
    const fetchImpl = vi.fn(async (url: string, init: RequestInit) => {
      puts.push({ url, init });
      return new Response("", { status: 200 });
    });

    const uploader = new GraphAttachmentUploader({
      graph,
      requestTimeoutMs: 1000,
      fetchImpl: fetchImpl as unknown as typeof fetch,
    });
    const size = UPLOAD_CHUNK_BYTES + 100; // → two chunks
    await uploader.upload(account, "msg1", attachment(size));

    // createUploadSession went through the Graph client with the right item.
    const sessionReq = request.mock.calls[0]![1];
    expect(sessionReq.method).toBe("POST");
    expect(sessionReq.path).toBe("/me/messages/msg1/attachments/createUploadSession");
    expect((sessionReq.body as { AttachmentItem: { size: number } }).AttachmentItem.size).toBe(
      size,
    );

    // Two chunk PUTs to the upload URL, with correct Content-Range and NO token.
    expect(puts).toHaveLength(2);
    expect(puts[0]!.url).toBe("https://upload.example/abc");
    expect(puts[0]!.init.method).toBe("PUT");
    const h0 = puts[0]!.init.headers as Record<string, string>;
    const h1 = puts[1]!.init.headers as Record<string, string>;
    expect(h0["Content-Range"]).toBe(`bytes 0-${UPLOAD_CHUNK_BYTES - 1}/${size}`);
    expect(h1["Content-Range"]).toBe(`bytes ${UPLOAD_CHUNK_BYTES}-${size - 1}/${size}`);
    expect(h0["Authorization"]).toBeUndefined(); // pre-authed URL — token must NOT leak
  });

  it("refuses a non-HTTPS upload URL", async () => {
    const { graph } = graphReturning({ uploadUrl: "http://insecure.example/x" });
    const uploader = new GraphAttachmentUploader({
      graph,
      requestTimeoutMs: 1000,
      fetchImpl: vi.fn() as unknown as typeof fetch,
    });
    await expect(uploader.upload(account, "m", attachment(10))).rejects.toThrow(/non-HTTPS/i);
  });

  it("retries a transient (5xx) chunk failure with bounded backoff, then gives up", async () => {
    const { graph } = graphReturning({ uploadUrl: "https://upload.example/x" });
    const fetchImpl = vi.fn(async () => new Response("nope", { status: 500 }));
    const uploader = new GraphAttachmentUploader({
      graph,
      requestTimeoutMs: 1000,
      fetchImpl: fetchImpl as unknown as typeof fetch,
      retry: { maxRetries: 3, sleep: async () => undefined }, // instant
    });
    await expect(uploader.upload(account, "m", attachment(10))).rejects.toThrow(
      /upload of attachment "big.bin" failed/i,
    );
    expect(fetchImpl).toHaveBeenCalledTimes(4); // 1 try + 3 retries
  });

  it("retries a transient chunk failure then succeeds", async () => {
    const { graph } = graphReturning({ uploadUrl: "https://upload.example/x" });
    let n = 0;
    const fetchImpl = vi.fn(async () => {
      n += 1;
      return n === 1 ? new Response("", { status: 503 }) : new Response("", { status: 200 });
    });
    const uploader = new GraphAttachmentUploader({
      graph,
      requestTimeoutMs: 1000,
      fetchImpl: fetchImpl as unknown as typeof fetch,
      retry: { sleep: async () => undefined },
    });
    await uploader.upload(account, "m", attachment(10)); // resolves
    expect(fetchImpl).toHaveBeenCalledTimes(2); // 503 retried once, then 200
  });

  it("does NOT retry a non-transient 4xx (e.g. expired session)", async () => {
    const { graph } = graphReturning({ uploadUrl: "https://upload.example/x" });
    const fetchImpl = vi.fn(async () => new Response("gone", { status: 404 }));
    const uploader = new GraphAttachmentUploader({
      graph,
      requestTimeoutMs: 1000,
      fetchImpl: fetchImpl as unknown as typeof fetch,
      retry: { sleep: async () => undefined },
    });
    await expect(uploader.upload(account, "m", attachment(10))).rejects.toThrow(/HTTP 404/);
    expect(fetchImpl).toHaveBeenCalledTimes(1); // not retried
  });
});
