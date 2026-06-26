/**
 * Output-contract helpers and resource bounds (FR-OUT-*, NFR-PERF-1/2).
 *
 * The structured payload remains the authoritative, complete result; only the
 * human-readable summary is clamped to a character budget (FR-OUT-2).
 */

/** Max characters for a human-readable summary before clamping (NFR-PERF-1). */
export const RESPONSE_CHAR_BUDGET = 25_000;

/** Max messages returned for a single conversation read, newest kept (NFR-PERF-2 / FR-C3-2). */
export const CONVERSATION_MESSAGE_CAP = 100;

/** Max combined body characters across a conversation read (NFR-PERF-2 / FR-C3-2). */
export const CONVERSATION_BODY_CHAR_CAP = 20_000;

/** Default search page size and its bounded maximum (FR-C2-2). */
export const DEFAULT_PAGE_SIZE = 20;
export const MAX_PAGE_SIZE = 100;

/**
 * Max raw size (body + attachment bytes) for an outgoing message, validated
 * locally before the Graph call so oversize sends fail fast with a clear error
 * rather than after an upload (NFR-PERF-3). Reference limit ~25 MB; the
 * effective mailbox policy is confirmed against the live API by the operator.
 */
export const MAX_OUTGOING_MESSAGE_BYTES = 25 * 1024 * 1024;

/**
 * Threshold at/below which an attachment is sent INLINE as a `fileAttachment` in
 * the message request (one call). Microsoft Graph caps an inline `fileAttachment`
 * at ~3 MB; larger files are uploaded to the draft via an upload session instead
 * (mail/uploadSession.ts). The total message is still bounded by
 * {@link MAX_OUTGOING_MESSAGE_BYTES}.
 */
export const MAX_INLINE_ATTACHMENT_BYTES = 3 * 1024 * 1024;

/**
 * Chunk size for upload-session PUTs. Microsoft Graph requires every chunk except
 * the last to be a multiple of 320 KiB; this is 10 × 320 KiB (~3.1 MB).
 */
export const UPLOAD_CHUNK_BYTES = 10 * 320 * 1024;

/** Clamp a string to a character budget, appending an ellipsis when truncated. */
export function clampText(
  text: string,
  max: number = RESPONSE_CHAR_BUDGET,
): { text: string; truncated: boolean } {
  if (text.length <= max) return { text, truncated: false };
  return { text: `${text.slice(0, Math.max(0, max - 1))}…`, truncated: true };
}

/** Normalise a requested page size to [1, MAX_PAGE_SIZE], defaulting when unset (FR-C2-2). */
export function clampPageSize(requested: number | undefined): number {
  if (requested === undefined || !Number.isFinite(requested)) return DEFAULT_PAGE_SIZE;
  return Math.min(MAX_PAGE_SIZE, Math.max(1, Math.floor(requested)));
}
