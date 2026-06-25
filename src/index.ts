#!/usr/bin/env node
/**
 * MCP server entry point (default transport: stdio, launched by the MCP host —
 * NFR-OPS-2).
 *
 * STATUS: scaffold. The server starts and reports state to stderr, but NO
 * capability tools (C1–C8, spec §5/§6) are registered yet — they are added in
 * the build phase. This file establishes the transport wiring and the startup
 * contract (report connected accounts to stderr) only.
 */

import { fileURLToPath } from "node:url";
import { McpServer } from "@modelcontextprotocol/sdk/server/mcp.js";
import { StdioServerTransport } from "@modelcontextprotocol/sdk/server/stdio.js";
import { loadConfig } from "./config.js";

export function createServer(): McpServer {
  const server = new McpServer({
    name: "mcp-server-for-multiple-outlook-accounts",
    version: "0.1.0",
  });

  // TODO(build): register C1–C8 tools here, each with MCP behavioural
  // annotations (readOnlyHint / destructiveHint / idempotentHint /
  // openWorldHint). send_message and organize_mail MUST be destructiveHint:true
  // (NFR-OPS-4). See doc/architecture.md §"Tool registration".

  return server;
}

async function main(): Promise<void> {
  // Validate configuration early so a misconfigured env fails fast and clearly.
  loadConfig();

  const server = createServer();
  const transport = new StdioServerTransport();
  await server.connect(transport);

  // Startup status goes to stderr so it never corrupts the stdio JSON-RPC
  // channel (NFR-OPS-2). Connected-account enumeration is wired in the build
  // phase once the token store / account registry land.
  process.stderr.write(
    "[outlook-mcp] scaffold running on stdio — no capability tools registered yet.\n",
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
