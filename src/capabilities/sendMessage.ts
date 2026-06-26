/**
 * C5 — Send a message (FR-C5-1/2/4, NFR-REL-3).
 *
 * Composes the same inputs as create_draft and delivers immediately via a
 * SINGLE `POST /me/sendMail` call. Using one call (rather than create-draft +
 * send) avoids an ambiguous two-step window, and the `nonDuplicable` retry
 * class means only a pre-processing 429 is retried — never an ambiguous 5xx or
 * timeout that may already have sent. Together these guarantee no duplicate
 * delivery under retry (FR-C5-4 / NFR-REL-3).
 *
 * Sending is irreversible (FR-C5-2); the tool is annotated `destructiveHint:
 * true` at registration (FR-C5-3 / NFR-OPS-4) so the host can gate it.
 */

import type { ToolResult } from "../domain/types.js";
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

  // Single send call; nonDuplicable so a retry can never double-deliver
  // (FR-C5-4 / NFR-REL-3). Returns 202 Accepted with no body.
  await deps.graph.request<void>(account, {
    method: "POST",
    path: "/me/sendMail",
    body: { message: composed.message, saveToSentItems: true },
    retryClass: "nonDuplicable",
  });

  const structured: SendMessageStructured = {
    account: account.displayId,
    sent: true,
    to: composed.recipients.to,
    cc: composed.recipients.cc,
    bcc: composed.recipients.bcc,
    subject: composed.message.subject,
    has_attachments: (composed.message.attachments?.length ?? 0) > 0,
    is_reply: isReply,
  };

  const subjectText = composed.message.subject ? `"${composed.message.subject}"` : "(no subject)";
  const summary = clampText(
    `Sent ${subjectText} from ${account.displayId} (${describeRecipients(composed.recipients)}).`,
  ).text;

  return { summary, structured };
}
