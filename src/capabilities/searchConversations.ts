/**
 * C2 — Search conversations (FR-C2-1..6).
 *
 * Translates the query (search/translate.ts), pages `GET /me/messages`, and
 * groups the returned messages by `conversationId` into conversation summaries.
 * Because the list response already carries the summary fields (subject, from,
 * receivedDateTime, bodyPreview), no per-conversation detail fetch is needed —
 * so there is no fan-out to bound (FR-C2-6) and no per-entry fetch to degrade
 * (FR-C2-5) on this provider. Pagination returns Graph's opaque
 * `@odata.nextLink` as the cursor (FR-C2-3).
 *
 * Annotations (NFR-OPS-4): read-only, non-destructive, idempotent, open-world.
 */

import type { AccountRegistry, GraphClient, GraphRequest } from "../domain/contracts.js";
import type { ToolResult } from "../domain/types.js";
import { clampPageSize, clampText } from "../output/contract.js";
import { translate } from "../search/translate.js";
import { formatRecipient, type GraphListResponse, type GraphMessage } from "../graph/types.js";

export interface SearchArgs {
  readonly account?: string;
  readonly query: string;
  readonly pageSize?: number;
  readonly pageToken?: string;
}

export interface ConversationSummaryOut {
  readonly conversation_id: string;
  readonly subject?: string;
  readonly sender?: string;
  readonly date?: string;
  readonly snippet?: string;
}

export interface SearchStructured {
  readonly account: string;
  readonly query: string;
  readonly count: number;
  readonly conversations: ConversationSummaryOut[];
  readonly next_page_token?: string;
}

export interface SearchDeps {
  readonly registry: AccountRegistry;
  readonly graph: GraphClient;
}

const SELECT = "id,conversationId,subject,from,receivedDateTime,bodyPreview";

export async function searchConversations(
  deps: SearchDeps,
  args: SearchArgs,
): Promise<ToolResult<SearchStructured>> {
  const account = await deps.registry.resolve(args.account);

  const query = args.query?.trim();
  if (!query) throw new Error("A search query is required."); // FR-C2-1 / FR-ERR-3
  const pageSize = clampPageSize(args.pageSize);

  let req: GraphRequest;
  if (args.pageToken) {
    // Opaque next-page cursor: fetch the nextLink as-is (FR-C2-3).
    req = { method: "GET", path: args.pageToken, retryClass: "safe" };
  } else {
    const { filter, search } = translate(query); // may throw SearchQueryError (FR-ERR-3)
    const queryParams: Record<string, string | number> = { $select: SELECT, $top: pageSize };
    if (search) {
      queryParams.$search = `"${search}"`; // Graph orders $search by relevance ($orderby not allowed).
    } else {
      queryParams.$orderby = "receivedDateTime desc";
      if (filter) queryParams.$filter = filter;
    }
    req = { method: "GET", path: "/me/messages", query: queryParams, retryClass: "safe" };
  }

  const page = await deps.graph.request<GraphListResponse<GraphMessage>>(account, req);

  // Group by conversation, keeping first (newest, given the ordering) per conversation.
  const conversations: ConversationSummaryOut[] = [];
  const seen = new Set<string>();
  for (const m of page.value) {
    const key = m.conversationId ?? m.id;
    if (seen.has(key)) continue;
    seen.add(key);
    conversations.push({
      conversation_id: key,
      subject: m.subject,
      sender: formatRecipient(m.from),
      date: m.receivedDateTime,
      snippet: m.bodyPreview,
    });
  }

  const nextPageToken = page["@odata.nextLink"];
  const structured: SearchStructured = {
    account: account.displayId,
    query,
    count: conversations.length,
    conversations,
    ...(nextPageToken ? { next_page_token: nextPageToken } : {}),
  };

  const lines = conversations.map(
    (c) =>
      `- ${c.subject ?? "(no subject)"} — ${c.sender ?? "unknown"}` +
      `${c.date ? ` (${c.date})` : ""}${c.snippet ? `\n    ${c.snippet}` : ""}`,
  );
  const header =
    conversations.length === 0
      ? `No conversations matched in ${account.displayId}.`
      : `${conversations.length} conversation(s) in ${account.displayId}` +
        `${nextPageToken ? " (more available)" : ""}:`;
  const summary = clampText([header, ...lines].join("\n")).text;

  return { summary, structured };
}
