/**
 * C4 — Create a draft (FR-C4-1..5).
 *
 * Composes a message (recipients, subject, body, attachments, optional reply
 * threading) and persists it as a draft via `POST /me/messages` — which files
 * the message into Drafts with `isDraft = true` and does NOT send it (FR-C4-2).
 * The call is `nonDuplicable`: a draft create must not be replayed on an
 * ambiguous failure, so only a pre-processing 429 is retried (NFR-REL-3).
 *
 * Attachments at/under the inline limit ride inside that request; larger ones
 * are uploaded to the created draft via an upload session (mail/uploadSession.ts).
 *
 * Annotations (NFR-OPS-4): write, non-destructive (a draft is reversible),
 * non-idempotent, open-world.
 */

import type { ToolResult } from "../domain/types.js";
import type { GraphMessage } from "../graph/types.js";
import { clampText } from "../output/contract.js";
import {
  buildOutgoing,
  describeRecipients,
  type OutgoingArgs,
  type WriteDeps,
} from "./outgoing.js";

export type CreateDraftArgs = OutgoingArgs;

export interface CreateDraftStructured {
  readonly account: string;
  readonly draft_id: string;
  readonly conversation_id?: string;
  readonly web_link?: string;
  readonly to: string[];
  readonly cc: string[];
  readonly bcc: string[];
  readonly subject?: string;
  readonly has_attachments: boolean;
  readonly is_reply: boolean;
}

export async function createDraft(
  deps: WriteDeps,
  args: CreateDraftArgs,
): Promise<ToolResult<CreateDraftStructured>> {
  const account = await deps.registry.resolve(args.account);
  const { composed, isReply } = await buildOutgoing(deps, account, args);

  // POST /me/messages creates the message as a draft (isDraft=true); it is not
  // sent. Draft-create is non-duplicable (NFR-REL-3).
  const created = await deps.graph.request<GraphMessage>(account, {
    method: "POST",
    path: "/me/messages",
    body: composed.message,
    retryClass: "nonDuplicable",
  });

  // Upload any large attachments to the freshly-created draft.
  for (const attachment of composed.uploadAttachments) {
    await deps.uploader.upload(account, created.id, attachment);
  }

  const hasAttachments =
    (composed.message.attachments?.length ?? 0) + composed.uploadAttachments.length > 0;
  const structured: CreateDraftStructured = {
    account: account.displayId,
    draft_id: created.id,
    ...(created.conversationId ? { conversation_id: created.conversationId } : {}),
    ...(created.webLink ? { web_link: created.webLink } : {}),
    to: composed.recipients.to,
    cc: composed.recipients.cc,
    bcc: composed.recipients.bcc,
    subject: composed.message.subject,
    has_attachments: hasAttachments,
    is_reply: isReply,
  };

  const subjectText = composed.message.subject ? `"${composed.message.subject}"` : "(no subject)";
  const summary = clampText(
    `Draft ${subjectText} created in ${account.displayId} (${describeRecipients(
      composed.recipients,
    )}). It was saved to Drafts and not sent.`,
  ).text;

  return { summary, structured };
}
