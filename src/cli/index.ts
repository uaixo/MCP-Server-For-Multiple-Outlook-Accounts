#!/usr/bin/env node
/**
 * Out-of-band account-management CLI (spec §8). The MCP server itself never
 * initiates interactive consent; this CLI does, before the server is used.
 *
 * Subcommands (all FR-AUTH-*):
 *   connect   Connect a new account via PKCE + state loopback consent
 *             (FR-AUTH-1..7,10), store its refresh token, identify it via
 *             GET /me (FR-AUTH-6), and bind it to the issuing app registration
 *             (FR-ID-5).
 *   list      List connected accounts with the credential source each uses
 *             (FR-AUTH-8).
 *   remove    Remove a connected account (FR-AUTH-8).
 *
 * STATUS: scaffold. Argument parsing and usage are present; the consent /
 * MSAL / token-store logic is implemented in the build phase.
 */

type Subcommand = "connect" | "list" | "remove";

const USAGE = `outlook-mcp-auth — manage Outlook / Microsoft 365 accounts for the MCP server

Usage:
  outlook-mcp-auth connect [--credentials <path>]   Connect a new mailbox (browser consent)
  outlook-mcp-auth list                             List connected mailboxes
  outlook-mcp-auth remove <account>                 Remove a connected mailbox

All data is stored locally under OUTLOOK_MCP_DATA_DIR (default ~/.outlook-mcp).
`;

function isSubcommand(value: string | undefined): value is Subcommand {
  return value === "connect" || value === "list" || value === "remove";
}

function main(argv: string[]): number {
  const [, , sub] = argv;

  if (sub === undefined || sub === "--help" || sub === "-h") {
    process.stdout.write(USAGE);
    return sub === undefined ? 1 : 0;
  }

  if (!isSubcommand(sub)) {
    process.stderr.write(`Unknown command: ${sub}\n\n${USAGE}`);
    return 2;
  }

  // TODO(build): dispatch to connect/list/remove implementations.
  process.stderr.write(
    `[outlook-mcp-auth] '${sub}' is not implemented yet (scaffold). See doc/architecture.md §"Authentication & onboarding".\n`,
  );
  return 64; // EX_USAGE-ish: recognised but not yet available.
}

process.exit(main(process.argv));
