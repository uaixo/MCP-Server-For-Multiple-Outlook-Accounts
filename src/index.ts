#!/usr/bin/env node
/**
 * MCP server entry point (default transport: stdio, launched by the MCP host —
 * NFR-OPS-2).
 *
 * Phase 1: C1 list_accounts. Phase 2: C2 search_conversations, C3
 * read_conversation. Remaining capabilities (C4–C8) are added in later phases
 * (see doc/architecture.md §13).
 */

import { fileURLToPath } from "node:url";
import { z } from "zod";
import { McpServer } from "@modelcontextprotocol/sdk/server/mcp.js";
import { StdioServerTransport } from "@modelcontextprotocol/sdk/server/stdio.js";
import { loadConfig } from "./config.js";
import { FileTokenStore } from "./auth/tokenStore.js";
import { FileAccountRegistry } from "./auth/accountRegistry.js";
import { createMsalTokenProvider } from "./auth/tokenProvider.js";
import { FetchGraphClient } from "./graph/client.js";
import { listAccounts } from "./capabilities/listAccounts.js";
import { searchConversations } from "./capabilities/searchConversations.js";
import { readConversation } from "./capabilities/readConversation.js";
import { createDraft } from "./capabilities/createDraft.js";
import { sendMessage } from "./capabilities/sendMessage.js";
import type { OutgoingArgs } from "./capabilities/outgoing.js";
import { FsAttachmentReader } from "./mail/attachments.js";
import { MAX_PAGE_SIZE } from "./output/contract.js";
import type { AccountRegistry, AttachmentReader, GraphClient } from "./domain/contracts.js";
import type { ToolResult } from "./domain/types.js";

export interface ServerDeps {
  readonly registry: AccountRegistry;
  readonly graph: GraphClient;
  readonly attachments: AttachmentReader;
}

/** Tool-result envelope: dual channel (FR-OUT-1) with errors surfaced as tool errors (FR-ERR-1). */
async function toToolResult<T>(fn: () => Promise<ToolResult<T>>): Promise<{
  content: Array<{ type: "text"; text: string }>;
  structuredContent?: Record<string, unknown>;
  isError?: boolean;
}> {
  try {
    const result = await fn();
    return {
      content: [{ type: "text", text: result.summary }],
      structuredContent: result.structured as unknown as Record<string, unknown>,
    };
  } catch (e) {
    return {
      content: [{ type: "text", text: e instanceof Error ? e.message : String(e) }],
      isError: true,
    };
  }
}

const READ_ANNOTATIONS = {
  readOnlyHint: true,
  destructiveHint: false,
  idempotentHint: true,
  openWorldHint: true,
} as const;

// Write but reversible (a draft can be deleted): not destructive (FR-C4 / NFR-OPS-4).
const WRITE_ANNOTATIONS = {
  readOnlyHint: false,
  destructiveHint: false,
  idempotentHint: false,
  openWorldHint: true,
} as const;

// Sending is irreversible: MUST be destructive so the host can gate it
// (FR-C5-3 / NFR-OPS-4).
const SEND_ANNOTATIONS = {
  readOnlyHint: false,
  destructiveHint: true,
  idempotentHint: false,
  openWorldHint: true,
} as const;

/** zod schema for an attachment: exactly one of `path` / `content_base64` (validated in the reader). */
const attachmentSchema = z.object({
  filename: z.string().optional(),
  mime_type: z.string().optional(),
  path: z.string().optional(),
  content_base64: z.string().optional(),
});

/** Shared input schema for the two write tools (snake_case for the tool surface). */
const composeInputSchema = {
  account: z.string().optional(),
  to: z.array(z.string()).min(1),
  cc: z.array(z.string()).optional(),
  bcc: z.array(z.string()).optional(),
  subject: z.string().optional(),
  body: z.string(),
  is_html: z.boolean().optional(),
  reply_to_conversation_id: z.string().optional(),
  attachments: z.array(attachmentSchema).optional(),
} as const;

type ComposeToolArgs = {
  account?: string;
  to: string[];
  cc?: string[];
  bcc?: string[];
  subject?: string;
  body: string;
  is_html?: boolean;
  reply_to_conversation_id?: string;
  attachments?: Array<z.infer<typeof attachmentSchema>>;
};

/** Map the snake_case tool input to the camelCase {@link OutgoingArgs}. */
function toOutgoingArgs(args: ComposeToolArgs): OutgoingArgs {
  return {
    account: args.account,
    to: args.to,
    cc: args.cc,
    bcc: args.bcc,
    subject: args.subject,
    body: args.body,
    isHtml: args.is_html,
    replyToConversationId: args.reply_to_conversation_id,
    attachments: args.attachments?.map((a) => ({
      filename: a.filename,
      mimeType: a.mime_type,
      path: a.path,
      contentBase64: a.content_base64,
    })),
  };
}

export function createServer(deps: ServerDeps): McpServer {
  const server = new McpServer({
    name: "mcp-server-for-multiple-outlook-accounts",
    version: "0.2.0",
  });

  // C1 — list_accounts (no account selector; FR-ID-1).
  server.registerTool(
    "list_accounts",
    {
      title: "List connected accounts",
      description:
        "List the Outlook / Microsoft 365 mailboxes connected to this server. " +
        "Use the returned identities as the 'account' selector for other tools.",
      annotations: { ...READ_ANNOTATIONS, openWorldHint: false }, // closed-world local set.
    },
    () => toToolResult(() => listAccounts(deps.registry)),
  );

  // C2 — search_conversations (FR-C2-*).
  server.registerTool(
    "search_conversations",
    {
      title: "Search conversations",
      description:
        "Search a mailbox for conversations. Query supports free text plus operators: " +
        "from:, to:, subject:, is:read/is:unread, has:attachment, after:YYYY-MM-DD, before:YYYY-MM-DD. " +
        "Free-text/subject search cannot be combined with structured filters in one query.",
      inputSchema: {
        account: z.string().optional(),
        query: z.string(),
        page_size: z.number().int().min(1).max(MAX_PAGE_SIZE).optional(),
        page_token: z.string().optional(),
      },
      annotations: READ_ANNOTATIONS,
    },
    (args) =>
      toToolResult(() =>
        searchConversations(deps, {
          account: args.account,
          query: args.query,
          pageSize: args.page_size,
          pageToken: args.page_token,
        }),
      ),
  );

  // C3 — read_conversation (FR-C3-*).
  server.registerTool(
    "read_conversation",
    {
      title: "Read a conversation",
      description:
        "Read every message in a conversation: headers, plain-text body, and applied labels. " +
        "The payload is bounded; older messages and long bodies may be truncated.",
      inputSchema: {
        account: z.string().optional(),
        conversation_id: z.string(),
      },
      annotations: READ_ANNOTATIONS,
    },
    (args) =>
      toToolResult(() =>
        readConversation(deps, { account: args.account, conversationId: args.conversation_id }),
      ),
  );

  // C4 — create_draft (write, reversible; FR-C4-*).
  server.registerTool(
    "create_draft",
    {
      title: "Create a draft",
      description:
        "Compose a draft email (recipients, subject, body, optional attachments) and save it to " +
        "the account's Drafts. The draft is NOT sent. Recipients accept 'addr' or " +
        "'Display Name <addr>'. Attachments are each a local path (allow-listed) OR inline base64. " +
        "Pass reply_to_conversation_id to draft a threaded reply.",
      inputSchema: composeInputSchema,
      annotations: WRITE_ANNOTATIONS,
    },
    (args) => toToolResult(() => createDraft(deps, toOutgoingArgs(args))),
  );

  // C5 — send_message (destructive: irreversible delivery; FR-C5-*).
  server.registerTool(
    "send_message",
    {
      title: "Send a message",
      description:
        "Send an email immediately. This is irreversible. Same inputs as create_draft; pass " +
        "reply_to_conversation_id to send a threaded reply. Delivery is guaranteed not to " +
        "duplicate under retry.",
      inputSchema: composeInputSchema,
      annotations: SEND_ANNOTATIONS,
    },
    (args) => toToolResult(() => sendMessage(deps, toOutgoingArgs(args))),
  );

  // TODO(phase 4): C6–C8. organize_mail MUST be destructiveHint:true (NFR-OPS-4).

  return server;
}

async function main(): Promise<void> {
  const config = loadConfig();
  const store = new FileTokenStore({
    dataDir: config.dataDir,
    lockTimeoutMs: config.lockTimeoutMs,
  });
  const registry = new FileAccountRegistry(store);
  const graph = new FetchGraphClient({
    requestTimeoutMs: config.requestTimeoutMs,
    getToken: createMsalTokenProvider({ config, tokenStore: store }),
  });
  const attachments = new FsAttachmentReader(config.attachmentsAllowList);

  const server = createServer({ registry, graph, attachments });
  await server.connect(new StdioServerTransport());

  // Report connected accounts to stderr (NFR-OPS-2) — never stdout (JSON-RPC) or secrets (NFR-SEC-6).
  const accounts = await registry.list();
  process.stderr.write(
    `[outlook-mcp] ready on stdio — connected accounts: ${
      accounts.length ? accounts.map((a) => a.displayId).join(", ") : "none"
    }\n`,
  );
}

if (process.argv[1] && fileURLToPath(import.meta.url) === process.argv[1]) {
  main().catch((err: unknown) => {
    process.stderr.write(
      `[outlook-mcp] fatal: ${err instanceof Error ? err.message : String(err)}\n`,
    );
    process.exit(1);
  });
}
