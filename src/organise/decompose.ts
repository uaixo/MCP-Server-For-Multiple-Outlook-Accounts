/**
 * The C8 label-decomposition fan-out (FR-C8-6; architecture §6, provider-mapping
 * §3.1) — the core Outlook porting risk.
 *
 * Gmail collapses tagging, foldering, and read-state into one "label". Outlook
 * keeps them separate, so ONE neutral organise intent becomes a COMBINATION of
 * Graph operations against a target message:
 *   - category add/remove + read-state → a single `PATCH /me/messages/{id}`
 *     (Graph's category PATCH replaces the whole array, so we merge against the
 *     message's current categories here); and
 *   - a move (archive / trash / junk) → `POST /me/messages/{id}/move` to the
 *     Archive, Deleted Items, or Junk Email well-known folder.
 *
 * This module is pure (no I/O): the capability fetches each target message's
 * current state, calls `decompose`, and runs the returned operations.
 */

import type {
  GraphOperation,
  GraphRequest,
  OrganiseIntent,
  OrganiseTargetMessage,
} from "../domain/contracts.js";

/** Graph well-known folder ids for the move destinations (provider-mapping §3.1). */
export const ARCHIVE_FOLDER_ID = "archive";
export const DELETED_ITEMS_FOLDER_ID = "deleteditems";
export const JUNK_FOLDER_ID = "junkemail";

/**
 * The single move destination an intent requests, if any. archive/trash/junk are
 * mutually exclusive (validated by the capability), so at most one applies; the
 * order here is just a deterministic tie-break.
 */
export function moveDestination(
  intent: OrganiseIntent,
): { readonly id: string; readonly label: string } | undefined {
  if (intent.archive) return { id: ARCHIVE_FOLDER_ID, label: "Archive" };
  if (intent.trash) return { id: DELETED_ITEMS_FOLDER_ID, label: "Deleted Items" };
  if (intent.junk) return { id: JUNK_FOLDER_ID, label: "Junk Email" };
  return undefined;
}

/**
 * Merge a category change against the message's current categories. Adds win
 * insertion order; a name in both add and remove ends up removed. The result is
 * the full array Graph's PATCH will set (FR-C8 idempotent: re-applying converges).
 */
export function mergeCategories(
  current: readonly string[],
  add: readonly string[],
  remove: readonly string[],
): string[] {
  const result = new Map<string, string>();
  for (const c of current) result.set(c, c);
  for (const a of add) result.set(a, a);
  for (const r of remove) result.delete(r);
  return [...result.values()];
}

/** Categories a message will carry after the intent is applied (for the union report). */
export function resultingCategories(
  message: OrganiseTargetMessage,
  intent: OrganiseIntent,
): string[] {
  return mergeCategories(
    message.categories ?? [],
    intent.addLabelIds ?? [],
    intent.removeLabelIds ?? [],
  );
}

function describePatch(body: { categories?: string[]; isRead?: boolean }): string {
  const parts: string[] = [];
  if (body.categories !== undefined) parts.push(`categories=[${body.categories.join(", ")}]`);
  if (body.isRead !== undefined) parts.push(body.isRead ? "mark read" : "mark unread");
  return parts.join("; ");
}

export function decompose(
  message: OrganiseTargetMessage,
  intent: OrganiseIntent,
): GraphOperation[] {
  const ops: GraphOperation[] = [];
  const path = `/me/messages/${encodeURIComponent(message.id)}`;

  // 1) categories[] + isRead collapse into one PATCH.
  const add = intent.addLabelIds ?? [];
  const remove = intent.removeLabelIds ?? [];
  const wantCategoryChange = add.length > 0 || remove.length > 0;

  const body: { categories?: string[]; isRead?: boolean } = {};
  if (wantCategoryChange) {
    body.categories = mergeCategories(message.categories ?? [], add, remove);
  }
  if (intent.markRead !== undefined) body.isRead = intent.markRead;
  if (body.categories !== undefined || body.isRead !== undefined) {
    const req: GraphRequest = { method: "PATCH", path, body, retryClass: "safe" };
    ops.push({ description: `${describePatch(body)} on ${message.id}`, request: req });
  }

  // 2) archive/trash/junk = move to a well-known folder (Graph has no combined
  // modify+move call). At most one applies (validated upstream).
  const move = moveDestination(intent);
  if (move) {
    const req: GraphRequest = {
      method: "POST",
      path: `${path}/move`,
      body: { destinationId: move.id },
      retryClass: "safe",
    };
    ops.push({ description: `move ${message.id} to ${move.label}`, request: req });
  }

  return ops;
}
