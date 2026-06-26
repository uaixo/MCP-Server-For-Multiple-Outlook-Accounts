/**
 * C5 — Send a message (FR-C5-1/2/4, NFR-REL-3).
 *
 * Composes the same inputs as create_draft. With no large attachments it
 * delivers via a SINGLE `POST /me/sendMail` call — one step, no ambiguous
 * window. When an attachment exceeds the inline limit it must be uploaded
 * against a message id, so the path becomes create-draft → upload → send.
 *
 * Either way the irreversible step (`sendMail`, or the final `…/send`) uses the
 * `nonDuplicable` retry class: only a pre-processing 429 is retried — never an
 * ambiguous 5xx/timeout that may already have sent — guaranteeing no duplicate
 * delivery under retry (FR-C5-4 / NFR-REL-3). If the two-step send fails
 * ambiguously, the draft is left in Drafts (unsent), never a duplicate.
 *
 * Sending is irreversible (FR-C5-2); the tool is annotated `destructiveHint:
 * true` at registration (FR-C5-3 / NFR-OPS-4) so the host can gate it.
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

export type SendMessageArgs = OutgoingArgs;

export interface SendMessageStructured {
  readonly account: string;
  readonly sent: true;
  readonly to: string[];
  readonly cc: string[];
  readonly bcc: string[];
  readonly subject?: string;
  readonly has_attachments: boolean;
  readonly is_reply: boolean;
}

export async function sendMessage(
  deps: WriteDeps,
  args: SendMessageArgs,
): Promise<ToolResult<SendMessageStructured>> {
  const account = await deps.registry.resolve(args.account);
  const { composed, isReply } = await buildOutgoing(deps, account, args);

  if (composed.uploadAttachments.length === 0) {
    // Fast path: a single send call. nonDuplicable so a retry can never
    // double-deliver (FR-C5-4 / NFR-REL-3). Returns 202 Accepted with no body.
    await deps.graph.request<void>(account, {
      method: "POST",
      path: "/me/sendMail",
      body: { message: composed.message, saveToSentItems: true },
      retryClass: "nonDuplicable",
    });
  } else {
    // Large attachments need a message id to upload against: create a draft,
    // upload, then send it. The final /send is the irreversible step and stays
    // nonDuplicable, so the no-duplicate guarantee holds (NFR-REL-3).
    const draft = await deps.graph.request<GraphMessage>(account, {
      method: "POST",
      path: "/me/messages",
      body: composed.message,
      retryClass: "nonDuplicable",
    });
    for (const attachment of composed.uploadAttachments) {
      await deps.uploader.upload(account, draft.id, attachment);
    }
    await deps.graph.request<void>(account, {
      method: "POST",
      path: `/me/messages/${encodeURIComponent(draft.id)}/send`,
      retryClass: "nonDuplicable",
    });
  }

  const structured: SendMessageStructured = {
    account: account.displayId,
    sent: true,
    to: composed.recipients.to,
    cc: composed.recipients.cc,
    bcc: composed.recipients.bcc,
    subject: composed.message.subject,
    has_attachments:
      (composed.message.attachments?.length ?? 0) + composed.uploadAttachments.length > 0,
    is_reply: isReply,
  };

  const subjectText = composed.message.subject ? `"${composed.message.subject}"` : "(no subject)";
  const summary = clampText(
    `Sent ${subjectText} from ${account.displayId} (${describeRecipients(composed.recipients)}).`,
  ).text;

  return { summary, structured };
}
