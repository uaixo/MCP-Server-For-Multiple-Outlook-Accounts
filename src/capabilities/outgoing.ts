/**
 * Shared write-path plumbing for create_draft (C4) and send_message (C5).
 *
 * Both capabilities take the same composition inputs; the only difference is the
 * final Graph call (create a draft vs. send). This module resolves attachments
 * (allow-list guarded), derives reply threading, and composes the message — so
 * each capability is left to do just its one Graph operation.
 */

import type {
  Account,
  AccountRegistry,
  AttachmentInput,
  AttachmentReader,
  GraphClient,
} from "../domain/contracts.js";
import { composeMessage, type ComposedMessage } from "../mail/compose.js";
import { fetchReplyContext } from "../mail/replyLookup.js";

export interface WriteDeps {
  readonly registry: AccountRegistry;
  readonly graph: GraphClient;
  readonly attachments: AttachmentReader;
}

/** Composition inputs accepted by both write tools (camelCase; tools map snake_case). */
export interface OutgoingArgs {
  readonly account?: string;
  readonly to: string[];
  readonly cc?: string[];
  readonly bcc?: string[];
  readonly subject?: string;
  readonly body: string;
  readonly isHtml?: boolean;
  readonly attachments?: AttachmentInput[];
  /** Reply target: thread into this conversation and default the `Re:` subject. */
  readonly replyToConversationId?: string;
}

export interface PreparedOutgoing {
  readonly composed: ComposedMessage;
  readonly isReply: boolean;
}

/** Resolve attachments + reply context and compose the outgoing message. */
export async function buildOutgoing(
  deps: WriteDeps,
  account: Account,
  args: OutgoingArgs,
): Promise<PreparedOutgoing> {
  const resolved = await Promise.all((args.attachments ?? []).map((a) => deps.attachments.read(a)));

  const replyId = args.replyToConversationId?.trim();
  const reply = replyId ? await fetchReplyContext(deps.graph, account, replyId) : undefined;

  const composed = composeMessage(
    {
      to: args.to,
      cc: args.cc,
      bcc: args.bcc,
      subject: args.subject,
      body: args.body,
      isHtml: args.isHtml,
      replyToConversationId: replyId,
    },
    resolved,
    reply ? { reply } : {},
  );

  return { composed, isReply: reply !== undefined };
}

/** Compact `to`/`cc` description for a tool summary line. */
export function describeRecipients(recipients: { to: string[]; cc: string[] }): string {
  const parts = [`to ${recipients.to.join(", ")}`];
  if (recipients.cc.length) parts.push(`cc ${recipients.cc.join(", ")}`);
  return parts.join("; ");
}
