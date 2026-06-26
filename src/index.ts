#!/usr/bin/env node
/**
 * MCP server entry point (default transport: stdio, launched by the MCP host —
 * NFR-OPS-2).
 *
 * Phase 1: C1 list_accounts. Phase 2: C2 search_conversations, C3
 * read_conversation. Phase 3: C4 create_draft, C5 send_message. Phase 4: C6
 * list_labels, C7 create_label, C8 organize_mail (see doc/architecture.md §13).
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
import { listLabels } from "./capabilities/listLabels.js";
import { createLabel } from "./capabilities/createLabel.js";
import { organizeMail } from "./capabilities/organizeMail.js";
import type { OutgoingArgs } from "./capabilities/outgoing.js";
import { FsAttachmentReader } from "./mail/attachments.js";
import { BoundedConcurrency } from "./util/bounded.js";
import { MAX_PAGE_SIZE } from "./output/contract.js";
import type {
  AccountRegistry,
  AttachmentReader,
  ConcurrencyLimiter,
  GraphClient,
} from "./domain/contracts.js";
import type { ToolResult } from "./domain/types.js";

export interface ServerDeps {
  readonly registry: AccountRegistry;
  readonly graph: GraphClient;
  readonly attachments: AttachmentReader;
  readonly limiter: ConcurrencyLimiter;
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

// Organise: destructive (removals / moves to trash/junk are non-additive) but
// idempotent — re-applying the same change set converges (FR-C8-5 / NFR-OPS-4).
const ORGANISE_ANNOTATIONS = {
  readOnlyHint: false,
  destructiveHint: true,
  idempotentHint: true,
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

  // C6 — list_labels (read-only; FR-C6-*).
  server.registerTool(
    "list_labels",
    {
      title: "List organisation labels",
      description:
        "List the account's categories (tags) and mail folders (locations), each with a stable " +
        "id, display name, kind, and whether it is a system label. Use these ids with organize_mail.",
      inputSchema: { account: z.string().optional() },
      annotations: READ_ANNOTATIONS,
    },
    (args) => toToolResult(() => listLabels(deps, { account: args.account })),
  );

  // C7 — create_label (write, additive; FR-C7-*).
  server.registerTool(
    "create_label",
    {
      title: "Create a label",
      description:
        "Create a new category (tag) or mail folder (location). kind='category' makes a tag " +
        "(optionally with a colour preset); kind='folder' makes a folder, nested under " +
        "parent_folder_id when given.",
      inputSchema: {
        account: z.string().optional(),
        name: z.string(),
        kind: z.enum(["category", "folder"]).optional(),
        color: z.string().optional(),
        parent_folder_id: z.string().optional(),
      },
      annotations: WRITE_ANNOTATIONS,
    },
    (args) =>
      toToolResult(() =>
        createLabel(deps, {
          account: args.account,
          name: args.name,
          kind: args.kind,
          color: args.color,
          parentFolderId: args.parent_folder_id,
        }),
      ),
  );

  // C8 — organize_mail (destructive; FR-C8-*).
  server.registerTool(
    "organize_mail",
    {
      title: "Organise mail",
      description:
        "Apply organisation changes to exactly one target — a conversation OR a single message. " +
        "Add/remove category labels, mark read/unread, and/or archive (remove from Inbox). " +
        "At least one change is required. Applied to a conversation, it affects every message.",
      inputSchema: {
        account: z.string().optional(),
        conversation_id: z.string().optional(),
        message_id: z.string().optional(),
        add_labels: z.array(z.string()).optional(),
        remove_labels: z.array(z.string()).optional(),
        mark_read: z.boolean().optional(),
        archive: z.boolean().optional(),
      },
      annotations: ORGANISE_ANNOTATIONS,
    },
    (args) =>
      toToolResult(() =>
        organizeMail(deps, {
          account: args.account,
          conversationId: args.conversation_id,
          messageId: args.message_id,
          addLabels: args.add_labels,
          removeLabels: args.remove_labels,
          markRead: args.mark_read,
          archive: args.archive,
        }),
      ),
  );

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
  const limiter = new BoundedConcurrency();

  const server = createServer({ registry, graph, attachments, limiter });
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
