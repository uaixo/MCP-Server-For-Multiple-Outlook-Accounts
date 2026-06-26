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
import { MAX_PAGE_SIZE } from "./output/contract.js";
import type { AccountRegistry, GraphClient } from "./domain/contracts.js";
import type { ToolResult } from "./domain/types.js";

export interface ServerDeps {
  readonly registry: AccountRegistry;
  readonly graph: GraphClient;
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

  // TODO(phase 3+): C4–C8. send_message and organize_mail MUST be destructiveHint:true (NFR-OPS-4).

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

  const server = createServer({ registry, graph });
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
