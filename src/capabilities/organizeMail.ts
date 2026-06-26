/**
 * C8 — Organise mail (FR-C8-1..6) — the unified "label" operation.
 *
 * Applies one neutral organise intent (add/remove category tags, mark
 * read/unread, archive) to exactly one target — a single message OR a whole
 * conversation. Because Outlook splits the neutral label concept, the work fans
 * out (organise/decompose.ts): per target message we PATCH categories[]/isRead
 * and/or POST a move. A conversation has no single modify call, so we enumerate
 * its message ids and apply per message, bounding concurrency (NFR-REL-4), then
 * report the UNION of resulting labels across the conversation (FR-C8-4).
 *
 * Validation runs before any Graph call (FR-C8-1/2 / FR-ERR-3): exactly one
 * target, and at least one change.
 *
 * Partial application: a conversation fan-out is not transactional. If a
 * per-message op fails, messages processed before it keep their change and the
 * call rejects (no structured result). Because the operation is idempotent,
 * simply re-running the same request safely converges the whole conversation.
 *
 * Annotations (NFR-OPS-4): write, destructive (removals/moves are non-additive),
 * idempotent (re-applying converges), open-world.
 */

import type {
  AccountRegistry,
  ConcurrencyLimiter,
  GraphClient,
  GraphRequest,
  OrganiseIntent,
  OrganiseTargetMessage,
} from "../domain/contracts.js";
import type { ToolResult } from "../domain/types.js";
import type { GraphMessage } from "../graph/types.js";
import { collectPaged } from "../graph/paginate.js";
import { decompose, resultingCategories } from "../organise/decompose.js";
import { clampText } from "../output/contract.js";

/** Max messages organised for a single conversation target. */
export const ORGANISE_MESSAGE_CAP = 500;

const SELECT = "id,categories,isRead";

export interface OrganizeMailArgs {
  readonly account?: string;
  readonly conversationId?: string;
  readonly messageId?: string;
  readonly addLabels?: string[];
  readonly removeLabels?: string[];
  readonly markRead?: boolean;
  readonly archive?: boolean;
}

export interface OrganizeMailStructured {
  readonly account: string;
  readonly target: { readonly type: "conversation" | "message"; readonly id: string };
  readonly message_count: number;
  readonly truncated: boolean;
  readonly added: string[];
  readonly removed: string[];
  readonly marked_read?: boolean;
  readonly archived?: boolean;
  /** Union of category labels applied across the targeted message(s) (FR-C8-4). */
  readonly labels: string[];
}

export interface OrganizeMailDeps {
  readonly registry: AccountRegistry;
  readonly graph: GraphClient;
  readonly limiter: ConcurrencyLimiter;
}

function buildIntent(args: OrganizeMailArgs): OrganiseIntent {
  return {
    addLabelIds: args.addLabels,
    removeLabelIds: args.removeLabels,
    markRead: args.markRead,
    archive: args.archive,
  };
}

function hasChange(args: OrganizeMailArgs): boolean {
  return (
    (args.addLabels?.length ?? 0) > 0 ||
    (args.removeLabels?.length ?? 0) > 0 ||
    args.markRead !== undefined ||
    args.archive === true
  );
}

export async function organizeMail(
  deps: OrganizeMailDeps,
  args: OrganizeMailArgs,
): Promise<ToolResult<OrganizeMailStructured>> {
  const account = await deps.registry.resolve(args.account);

  // FR-C8-1: exactly one target.
  const conversationId = args.conversationId?.trim();
  const messageId = args.messageId?.trim();
  if (!!conversationId === !!messageId) {
    throw new Error("Provide exactly one of conversation_id or message_id.");
  }
  // FR-C8-2: at least one change.
  if (!hasChange(args)) {
    throw new Error(
      "No changes requested. Provide at least one of add_labels, remove_labels, mark_read, or archive.",
    );
  }

  const intent = buildIntent(args);

  // Resolve the target message(s) with the state decompose needs.
  let messages: OrganiseTargetMessage[];
  let truncated = false;
  if (messageId) {
    const m = await deps.graph.request<GraphMessage>(account, {
      method: "GET",
      path: `/me/messages/${encodeURIComponent(messageId)}`,
      query: { $select: SELECT },
      retryClass: "safe",
    });
    messages = [{ id: m.id, categories: m.categories, isRead: m.isRead }];
  } else {
    const escaped = conversationId!.replace(/'/g, "''");
    const req: GraphRequest = {
      method: "GET",
      path: "/me/messages",
      query: {
        $filter: `conversationId eq '${escaped}'`,
        $select: SELECT,
        // Deterministic order so that, if a conversation exceeds the cap, it is
        // the newest messages that are organised (not a provider-arbitrary set).
        $orderby: "receivedDateTime desc",
        $top: 100,
      },
      retryClass: "safe",
    };
    const page = await collectPaged<GraphMessage>(deps.graph, account, req, ORGANISE_MESSAGE_CAP);
    truncated = page.truncated;
    messages = page.items.map((m) => ({ id: m.id, categories: m.categories, isRead: m.isRead }));
    if (messages.length === 0) {
      throw new Error(`No messages found for conversation ${conversationId}.`);
    }
  }

  // Fan out: one task per message runs its ops in order (a move changes the id,
  // so PATCH must precede move); the limiter bounds cross-message concurrency.
  const tasks = messages.map((message) => async () => {
    for (const op of decompose(message, intent)) {
      await deps.graph.request(account, op.request);
    }
    return resultingCategories(message, intent);
  });
  const perMessageLabels = await deps.limiter.run(tasks);

  const labels = [...new Set(perMessageLabels.flat())].sort((a, b) => a.localeCompare(b));

  const structured: OrganizeMailStructured = {
    account: account.displayId,
    target: messageId
      ? { type: "message", id: messageId }
      : { type: "conversation", id: conversationId! },
    message_count: messages.length,
    truncated,
    added: args.addLabels ?? [],
    removed: args.removeLabels ?? [],
    ...(args.markRead !== undefined ? { marked_read: args.markRead } : {}),
    ...(args.archive ? { archived: true } : {}),
    labels,
  };

  const changes: string[] = [];
  if (structured.added.length) changes.push(`+[${structured.added.join(", ")}]`);
  if (structured.removed.length) changes.push(`-[${structured.removed.join(", ")}]`);
  if (structured.marked_read !== undefined)
    changes.push(structured.marked_read ? "mark read" : "mark unread");
  if (structured.archived) changes.push("archive");

  const summary = clampText(
    `Organised ${messages.length} message(s) in ${structured.target.type} ` +
      `${structured.target.id} of ${account.displayId}: ${changes.join(", ")}.` +
      (labels.length ? ` Labels now: ${labels.join(", ")}.` : "") +
      (truncated ? ` (Only the first ${ORGANISE_MESSAGE_CAP} messages were organised.)` : ""),
  ).text;

  return { summary, structured };
}
