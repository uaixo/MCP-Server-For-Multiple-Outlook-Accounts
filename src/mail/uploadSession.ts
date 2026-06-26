/**
 * Large-attachment upload via a Microsoft Graph upload session (provider-mapping
 * §3, attachment notes). Lifts the ~3 MB inline `fileAttachment` ceiling toward
 * the message size cap.
 *
 * Flow for one attachment on an existing draft message:
 *   1. `POST /me/messages/{id}/attachments/createUploadSession` (via the Graph
 *      client) → a Graph-issued `uploadUrl`; and
 *   2. PUT the bytes to that URL in 320 KiB-aligned chunks.
 *
 * Egress note (security): the `uploadUrl` is on a NON-Graph host, so the chunk
 * PUTs cannot go through the origin-pinned Graph client. They are sent here with
 * a dedicated `fetch`, and deliberately carry NO `Authorization` header — the URL
 * is pre-authenticated by Graph. We require https so bytes never go in the clear,
 * and only ever PUT to the exact URL Graph returned (never a caller-supplied one).
 *
 * No-duplicate-send: this only uploads to a DRAFT; the irreversible send is a
 * separate `…/send` call gated by the `nonDuplicable` retry class, so retrying a
 * chunk PUT (idempotent by `Content-Range`) can never cause a duplicate delivery.
 */

import type {
  Account,
  AttachmentUploader,
  GraphClient,
  ResolvedAttachment,
} from "../domain/contracts.js";
import type { GraphUploadSession } from "../graph/types.js";
import { UPLOAD_CHUNK_BYTES } from "../output/contract.js";

export interface UploaderDeps {
  readonly graph: GraphClient;
  readonly requestTimeoutMs: number;
  /** Injectable for tests; defaults to global `fetch`. Used ONLY for chunk PUTs. */
  readonly fetchImpl?: typeof fetch;
}

export class GraphAttachmentUploader implements AttachmentUploader {
  constructor(private readonly deps: UploaderDeps) {}

  async upload(account: Account, messageId: string, attachment: ResolvedAttachment): Promise<void> {
    const total = attachment.bytes.byteLength;

    // 1) Create the upload session (a Graph endpoint → via the origin-pinned client).
    // Creating a session on a draft has no irreversible effect, so it is `safe`.
    const session = await this.deps.graph.request<GraphUploadSession>(account, {
      method: "POST",
      path: `/me/messages/${encodeURIComponent(messageId)}/attachments/createUploadSession`,
      body: {
        AttachmentItem: {
          attachmentType: "file",
          name: attachment.filename,
          size: total,
          contentType: attachment.mimeType,
        },
      },
      retryClass: "safe",
    });

    const uploadUrl = session.uploadUrl;
    let parsed: URL;
    try {
      parsed = new URL(uploadUrl);
    } catch {
      throw new Error("Microsoft Graph returned an invalid attachment upload URL.");
    }
    if (parsed.protocol !== "https:") {
      throw new Error("Refusing to upload an attachment over a non-HTTPS upload URL.");
    }

    // 2) PUT the bytes in chunks. No Authorization header — the URL is pre-authed.
    const doFetch = this.deps.fetchImpl ?? fetch;
    for (let start = 0; start < total; start += UPLOAD_CHUNK_BYTES) {
      const end = Math.min(start + UPLOAD_CHUNK_BYTES, total);
      const chunk = attachment.bytes.subarray(start, end);

      let res: Response;
      try {
        res = await doFetch(uploadUrl, {
          method: "PUT",
          headers: {
            "Content-Type": "application/octet-stream",
            "Content-Range": `bytes ${start}-${end - 1}/${total}`,
          },
          body: chunk,
          signal: AbortSignal.timeout(this.deps.requestTimeoutMs),
        });
      } catch (e) {
        throw new Error(
          `Upload of attachment "${attachment.filename}" failed: ` +
            (e instanceof Error ? e.message : String(e)),
        );
      }
      if (!res.ok) {
        throw new Error(
          `Upload of attachment "${attachment.filename}" failed (HTTP ${res.status}).`,
        );
      }
    }
  }
}
