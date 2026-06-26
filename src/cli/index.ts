#!/usr/bin/env node
/**
 * Out-of-band account-management CLI (spec §8). The MCP server itself never
 * initiates interactive consent; this CLI does, before the server is used.
 *
 *   connect [--credentials <path>] [--source <id>]   FR-AUTH-1..7,10
 *   list                                             FR-AUTH-8
 *   remove <account>                                 FR-AUTH-8
 */

import { fileURLToPath } from "node:url";
import { runConnect } from "./connect.js";
import { runList } from "./list.js";
import { runRemove } from "./remove.js";
import { redactError } from "../util/redact.js";

const USAGE = `outlook-mcp-auth — manage Outlook / Microsoft 365 accounts for the MCP server

Usage:
  outlook-mcp-auth connect [--credentials <path>] [--source <id>]   Connect a new mailbox (browser consent)
  outlook-mcp-auth list                                             List connected mailboxes
  outlook-mcp-auth remove <account>                                 Remove a connected mailbox

All data is stored locally under OUTLOOK_MCP_DATA_DIR (default ~/.outlook-mcp).
`;

export async function run(argv: string[]): Promise<number> {
  const [, , sub, ...rest] = argv;

  switch (sub) {
    case "connect":
      return runConnect(rest);
    case "list":
      return runList();
    case "remove":
      return runRemove(rest);
    case undefined:
      process.stdout.write(USAGE);
      return 1;
    case "--help":
    case "-h":
      process.stdout.write(USAGE);
      return 0;
    default:
      process.stderr.write(`Unknown command: ${sub}\n\n${USAGE}`);
      return 2;
  }
}

// Run only when executed directly (not when imported by tests).
if (process.argv[1] && fileURLToPath(import.meta.url) === process.argv[1]) {
  run(process.argv)
    .then((code) => process.exit(code))
    .catch((err: unknown) => {
      // Redact before logging: a thrown error may carry token/credential material (NFR-SEC-6).
      process.stderr.write(`[outlook-mcp-auth] fatal: ${redactError(err)}\n`);
      process.exit(1);
    });
}
