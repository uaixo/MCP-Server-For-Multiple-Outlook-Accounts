/**
 * C3 — Read a full conversation (FR-C3-1..3).
 *
 * Fetches the conversation's messages via `GET /me/messages?$filter=conversationId eq '…'`,
 * newest first, and returns each message's headers, plain-text body, and applied
 * organisation labels. The payload is bounded (FR-C3-2 / NFR-PERF-2): at most
 * CONVERSATION_MESSAGE_CAP messages (newest kept) and CONVERSATION_BODY_CHAR_CAP
 * combined body characters; a `truncated` flag and `omitted_message_count` report
 * what was dropped. HTML bodies are rendered to readable text (FR-C3-3).
 *
 * Annotations (NFR-OPS-4): read-only, non-destructive, idempotent, open-world.
 */

import type { AccountRegistry, GraphClient, GraphRequest } from "../domain/contracts.js";
import type { ToolResult } from "../domain/types.js";
import {
  CONVERSATION_BODY_CHAR_CAP,
  CONVERSATION_MESSAGE_CAP,
  clampText,
} from "../output/contract.js";
import { htmlToText } from "../util/html.js";
import { formatRecipient, type GraphListResponse, type GraphMessage } from "../graph/types.js";

export interface ReadArgs {
  readonly account?: string;
  readonly conversationId: string;
}

export interface MessageOut {
  readonly id: string;
  readonly from?: string;
  readonly to: string[];
  readonly cc: string[];
  readonly date?: string;
  readonly subject?: string;
  readonly body_text: string;
  readonly labels: string[];
}

export interface ReadStructured {
  readonly account: string;
  readonly conversation_id: string;
  readonly message_count: number;
  readonly truncated: boolean;
  readonly omitted_message_count?: number;
  readonly messages: MessageOut[];
}

export interface ReadDeps {
  readonly registry: AccountRegistry;
  readonly graph: GraphClient;
}

const SELECT =
  "id,conversationId,subject,from,toRecipients,ccRecipients,receivedDateTime,body,bodyPreview,categories,isRead";

function bodyToText(m: GraphMessage): string {
  const content = m.body?.content;
  if (content && m.body?.contentType === "html") return htmlToText(content);
  if (content) return content.trim();
  return m.bodyPreview?.trim() ?? "";
}

function labelsOf(m: GraphMessage): string[] {
  const labels = [...(m.categories ?? [])];
  if (m.isRead === false) labels.push("Unread");
  return labels;
}

export async function readConversation(
  deps: ReadDeps,
  args: ReadArgs,
): Promise<ToolResult<ReadStructured>> {
  const account = await deps.registry.resolve(args.account);

  const conversationId = args.conversationId?.trim();
  if (!conversationId) throw new Error("A conversation_id is required."); // FR-ERR-3

  const escaped = conversationId.replace(/'/g, "''");
  const req: GraphRequest = {
    method: "GET",
    path: "/me/messages",
    query: {
      $filter: `conversationId eq '${escaped}'`,
      $select: SELECT,
      $orderby: "receivedDateTime desc",
      $top: CONVERSATION_MESSAGE_CAP,
      $count: "true",
    },
    headers: { ConsistencyLevel: "eventual" }, // required for $count on messages
    retryClass: "safe",
  };

  const page = await deps.graph.request<GraphListResponse<GraphMessage>>(account, req);
  const fetched = page.value;
  const total = page["@odata.count"];

  // Bound combined body characters, newest first (FR-C3-2 / NFR-PERF-2).
  let used = 0;
  let bodyTruncated = false;
  const messages: MessageOut[] = fetched.map((m) => {
    const full = bodyToText(m);
    let body = full;
    if (used + full.length > CONVERSATION_BODY_CHAR_CAP) {
      body = full.slice(0, Math.max(0, CONVERSATION_BODY_CHAR_CAP - used));
      if (body.length < full.length) bodyTruncated = true;
    }
    used += body.length;
    return {
      id: m.id,
      from: formatRecipient(m.from),
      to: (m.toRecipients ?? []).map(formatRecipient).filter((x): x is string => !!x),
      cc: (m.ccRecipients ?? []).map(formatRecipient).filter((x): x is string => !!x),
      date: m.receivedDateTime,
      subject: m.subject,
      body_text: body,
      labels: labelsOf(m),
    };
  });

  const omittedByCount =
    total !== undefined
      ? Math.max(0, total - fetched.length)
      : page["@odata.nextLink"]
        ? undefined // more exist but exact count unknown
        : 0;
  const truncated =
    bodyTruncated || (omittedByCount ?? 1) > 0 || page["@odata.nextLink"] !== undefined;

  const structured: ReadStructured = {
    account: account.displayId,
    conversation_id: conversationId,
    message_count: messages.length,
    truncated,
    ...(omittedByCount && omittedByCount > 0 ? { omitted_message_count: omittedByCount } : {}),
    messages,
  };

  const blocks = messages.map((m) => {
    const head = [
      `From: ${m.from ?? "unknown"}`,
      m.to.length ? `To: ${m.to.join(", ")}` : undefined,
      m.date ? `Date: ${m.date}` : undefined,
      m.subject ? `Subject: ${m.subject}` : undefined,
      m.labels.length ? `Labels: ${m.labels.join(", ")}` : undefined,
    ]
      .filter(Boolean)
      .join("\n");
    return `${head}\n\n${m.body_text}`;
  });
  const header =
    `Conversation ${conversationId} in ${account.displayId} — ${messages.length} message(s)` +
    `${truncated ? " (truncated)" : ""}:`;
  const summary = clampText([header, ...blocks].join("\n\n---\n\n")).text;

  return { summary, structured };
}
