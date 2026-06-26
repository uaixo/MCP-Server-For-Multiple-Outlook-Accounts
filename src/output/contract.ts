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
