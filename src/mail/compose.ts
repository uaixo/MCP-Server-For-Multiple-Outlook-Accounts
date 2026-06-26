/**
 * Build the outgoing Graph message JSON shared by create_draft (C4) and
 * send_message (C5) — FR-C4-1/4/5, FR-C5-1, NFR-PERF-3, NFR-SEC-5.
 *
 * This module is pure (no Graph or filesystem I/O): it takes the caller's
 * composition inputs plus already-resolved attachment bytes and returns the
 * message resource ready to POST, together with the computed raw size. Keeping
 * it pure makes the recipient parsing, `Re:` defaulting, reply threading, and
 * size guard fully unit-testable without a network.
 *
 * Threading (FR-C4-4 / FR-C5-1): when replying we set the RFC 5322
 * `In-Reply-To`/`References` headers from the original message's
 * `internetMessageId`, and default an omitted subject to the original prefixed
 * with `Re:`. (Graph's acceptance of these headers is confirmed live.)
 */

import type { ComposeInput, ResolvedAttachment } from "../domain/contracts.js";
import { MAX_INLINE_ATTACHMENT_BYTES, MAX_OUTGOING_MESSAGE_BYTES } from "../output/contract.js";
import {
  formatRecipient,
  toGraphRecipient,
  type GraphFileAttachment,
  type GraphInternetMessageHeader,
  type GraphOutgoingMessage,
  type GraphRecipient,
} from "../graph/types.js";
import { sanitizeHeaderValue } from "./sanitize.js";

export interface ComposeOptions {
  /** Override the outgoing-size limit (bytes); defaults to {@link MAX_OUTGOING_MESSAGE_BYTES}. */
  readonly maxBytes?: number;
  /** Original-message context when this is a reply (drives `Re:` default + threading). */
  readonly reply?: {
    readonly subject?: string;
    readonly internetMessageId?: string;
  };
}

export interface ComposedMessage {
  readonly message: GraphOutgoingMessage;
  /** Raw body + ALL attachment bytes (inline + upload), validated against the limit (NFR-PERF-3). */
  readonly sizeBytes: number;
  /** Human-readable recipient strings for the tool summary. */
  readonly recipients: { to: string[]; cc: string[]; bcc: string[] };
  /**
   * Attachments too large to inline (> {@link MAX_INLINE_ATTACHMENT_BYTES}). They
   * are NOT in `message.attachments`; the write capability uploads each to the
   * created draft via an upload session.
   */
  readonly uploadAttachments: ResolvedAttachment[];
}

const ADDRESS_RE = /^[^\s@]+@[^\s@]+\.[^\s@]+$/;

/** Parse a recipient — a bare `addr` or a `Display Name <addr>` form (FR-C4-5). */
export function parseRecipient(raw: string): GraphRecipient {
  // Strip header-injection characters first so a CR/LF can't break the parse
  // (and can't smuggle a header through a display name) — NFR-SEC-5.
  const sanitized = sanitizeHeaderValue(raw);
  if (!sanitized) throw new Error("A recipient cannot be empty.");

  const angle = /^(.*)<([^<>]+)>$/.exec(sanitized);
  const name = angle ? angle[1]!.trim() : undefined;
  const address = (angle ? angle[2]! : sanitized).trim();

  if (!ADDRESS_RE.test(address)) {
    throw new Error(`Invalid recipient address: ${JSON.stringify(raw)}.`);
  }
  return toGraphRecipient(address, name || undefined);
}

function parseList(values: readonly string[] | undefined): GraphRecipient[] {
  return (values ?? []).map(parseRecipient);
}

/** Default an omitted reply subject to the original prefixed with `Re:` (FR-C4-4). */
function ensureRePrefix(subject: string): string {
  const s = subject.trim();
  return /^re:/i.test(s) ? s : `Re: ${s}`;
}

function resolveSubject(input: ComposeInput, opts: ComposeOptions): string | undefined {
  const explicit = input.subject ? sanitizeHeaderValue(input.subject) : "";
  if (explicit) return explicit;
  const replySubject = opts.reply?.subject ? sanitizeHeaderValue(opts.reply.subject) : "";
  if (replySubject) return ensureRePrefix(replySubject);
  return undefined;
}

function toFileAttachment(a: ResolvedAttachment): GraphFileAttachment {
  return {
    "@odata.type": "#microsoft.graph.fileAttachment",
    name: a.filename,
    contentType: a.mimeType,
    contentBytes: Buffer.from(a.bytes).toString("base64"),
  };
}

function threadingHeaders(opts: ComposeOptions): GraphInternetMessageHeader[] | undefined {
  const id = opts.reply?.internetMessageId ? sanitizeHeaderValue(opts.reply.internetMessageId) : "";
  if (!id) return undefined;
  return [
    { name: "In-Reply-To", value: id },
    { name: "References", value: id },
  ];
}

export function composeMessage(
  input: ComposeInput,
  attachments: readonly ResolvedAttachment[],
  opts: ComposeOptions = {},
): ComposedMessage {
  const to = parseList(input.to);
  if (to.length === 0) throw new Error("At least one recipient (to) is required."); // FR-C4-1
  const cc = parseList(input.cc);
  const bcc = parseList(input.bcc);

  const subject = resolveSubject(input, opts);
  const body = input.body ?? "";

  // Split by the inline limit: small files ride inline in the message request;
  // larger ones are uploaded to the draft via an upload session afterwards.
  const inline = attachments.filter((a) => a.bytes.byteLength <= MAX_INLINE_ATTACHMENT_BYTES);
  const uploadAttachments = attachments.filter(
    (a) => a.bytes.byteLength > MAX_INLINE_ATTACHMENT_BYTES,
  );
  const fileAttachments = inline.map(toFileAttachment);
  const headers = threadingHeaders(opts);

  // Validate raw size locally before any Graph call (NFR-PERF-3) — counting ALL
  // attachments (inline + upload), since the whole message is bounded.
  const maxBytes = opts.maxBytes ?? MAX_OUTGOING_MESSAGE_BYTES;
  const attachmentBytes = attachments.reduce((n, a) => n + a.bytes.byteLength, 0);
  const sizeBytes = Buffer.byteLength(body, "utf8") + attachmentBytes;
  if (sizeBytes > maxBytes) {
    const mb = (n: number) => (n / (1024 * 1024)).toFixed(1);
    throw new Error(
      `Outgoing message is ${mb(sizeBytes)} MB, over the ${mb(maxBytes)} MB limit. ` +
        "Reduce the body or attachment size.",
    );
  }

  const message: GraphOutgoingMessage = {
    ...(subject !== undefined ? { subject } : {}),
    body: { contentType: input.isHtml ? "HTML" : "Text", content: body },
    toRecipients: to,
    ...(cc.length ? { ccRecipients: cc } : {}),
    ...(bcc.length ? { bccRecipients: bcc } : {}),
    ...(fileAttachments.length ? { attachments: fileAttachments } : {}),
    ...(headers ? { internetMessageHeaders: headers } : {}),
  };

  // Each parsed recipient always has an address, so formatRecipient is non-empty;
  // the `?? ""` is only to satisfy its `string | undefined` return type.
  const display = (r: GraphRecipient): string => formatRecipient(r) ?? "";
  return {
    message,
    sizeBytes,
    recipients: {
      to: to.map(display),
      cc: cc.map(display),
      bcc: bcc.map(display),
    },
    uploadAttachments,
  };
}
