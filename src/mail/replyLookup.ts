/**
 * Look up the threading context for a reply (FR-C4-4 / FR-C5-1).
 *
 * Given a conversation id, fetch its most recent message to obtain the original
 * subject (for the `Re:` default) and `internetMessageId` (for the
 * `In-Reply-To`/`References` headers). This is the only Graph call the write
 * path makes purely for threading; the compose layer stays pure and consumes
 * the returned context.
 */

import type { Account, GraphClient, GraphRequest } from "../domain/contracts.js";
import type { GraphListResponse, GraphMessage } from "../graph/types.js";

export interface ReplyContext {
  readonly subject?: string;
  readonly internetMessageId?: string;
}

export async function fetchReplyContext(
  graph: GraphClient,
  account: Account,
  conversationId: string,
): Promise<ReplyContext> {
  const escaped = conversationId.replace(/'/g, "''");
  const req: GraphRequest = {
    method: "GET",
    path: "/me/messages",
    query: {
      $filter: `conversationId eq '${escaped}'`,
      $orderby: "receivedDateTime desc",
      $top: 1,
      $select: "id,subject,internetMessageId,conversationId",
    },
    retryClass: "safe",
  };
  const page = await graph.request<GraphListResponse<GraphMessage>>(account, req);
  const latest = page.value[0];
  if (!latest) {
    throw new Error(`No messages found for conversation ${conversationId} to reply to.`);
  }
  return { subject: latest.subject, internetMessageId: latest.internetMessageId };
}
