#!/usr/bin/env node
/**
 * MCP server entry point (default transport: stdio, launched by the MCP host —
 * NFR-OPS-2).
 *
 * Phase 1: registers C1 `list_accounts`. Remaining capabilities (C2–C8) are
 * added in later phases (see doc/architecture.md §13).
 */

import { fileURLToPath } from "node:url";
import { McpServer } from "@modelcontextprotocol/sdk/server/mcp.js";
import { StdioServerTransport } from "@modelcontextprotocol/sdk/server/stdio.js";
import { loadConfig } from "./config.js";
import { FileTokenStore } from "./auth/tokenStore.js";
import { FileAccountRegistry } from "./auth/accountRegistry.js";
import { listAccounts } from "./capabilities/listAccounts.js";
import type { AccountRegistry } from "./domain/contracts.js";

export interface ServerDeps {
  readonly registry: AccountRegistry;
}

export function createServer(deps: ServerDeps): McpServer {
  const server = new McpServer({
    name: "mcp-server-for-multiple-outlook-accounts",
    version: "0.1.0",
  });

  // C1 — list_accounts. No account selector (FR-ID-1: every tool EXCEPT this).
  server.registerTool(
    "list_accounts",
    {
      title: "List connected accounts",
      description:
        "List the Outlook / Microsoft 365 mailboxes connected to this server. " +
        "Use the returned identities as the 'account' selector for other tools.",
      annotations: {
        readOnlyHint: true,
        destructiveHint: false,
        idempotentHint: true,
        openWorldHint: false, // closed-world: enumerates a fixed local set.
      },
    },
    async () => {
      const result = await listAccounts(deps.registry);
      return {
        content: [{ type: "text", text: result.summary }],
        structuredContent: result.structured as unknown as Record<string, unknown>,
      };
    },
  );

  // TODO(phase 2+): register C2–C8. send_message and organize_mail MUST be
  // destructiveHint:true (NFR-OPS-4).

  return server;
}

async function main(): Promise<void> {
  const config = loadConfig();
  const store = new FileTokenStore({
    dataDir: config.dataDir,
    lockTimeoutMs: config.lockTimeoutMs,
  });
  const registry = new FileAccountRegistry(store);

  const server = createServer({ registry });
  await server.connect(new StdioServerTransport());

  // Report connected accounts to stderr (NFR-OPS-2) — never stdout (it carries
  // the JSON-RPC channel) and never secrets (NFR-SEC-6).
  const accounts = await registry.list();
  process.stderr.write(
    `[outlook-mcp] ready on stdio — connected accounts: ${
      accounts.length ? accounts.map((a) => a.displayId).join(", ") : "none"
    }\n`,
  );
}

// Run only when executed directly (not when imported by tests).
if (process.argv[1] && fileURLToPath(import.meta.url) === process.argv[1]) {
  main().catch((err: unknown) => {
    process.stderr.write(
      `[outlook-mcp] fatal: ${err instanceof Error ? err.message : String(err)}\n`,
    );
    process.exit(1);
  });
}
