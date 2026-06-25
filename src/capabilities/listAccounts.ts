/**
 * C1 — List connected accounts (FR-C1-1..3).
 *
 * The discovery entry point: returns the identities of every account connected
 * to this server (local token-store enumeration; provider-independent —
 * provider-mapping §3). When none are connected it returns a NON-error result
 * telling the user how to connect one (FR-C1-2). Its output supplies the valid
 * values for every other tool's account selector (FR-C1-3).
 *
 * Annotations (NFR-OPS-4): read-only, non-destructive, idempotent, closed-world.
 */

import type { AccountRegistry } from "../domain/contracts.js";
import type { ToolResult } from "../domain/types.js";

export interface ListedAccount {
  /** Account identity as authenticated (FR-C1-1). */
  readonly account: string;
  /** The app registration that authorised it (FR-AUTH-8, FR-ID-5). */
  readonly credential_source: string;
}

export interface ListAccountsStructured {
  readonly account_count: number;
  readonly accounts: ListedAccount[];
}

const CONNECT_HINT = 'No accounts are connected. Run "outlook-mcp-auth connect" to add a mailbox.';

export async function listAccounts(
  registry: AccountRegistry,
): Promise<ToolResult<ListAccountsStructured>> {
  const accounts = await registry.list();

  if (accounts.length === 0) {
    // FR-C1-2: empty is a normal, non-error outcome with guidance.
    return { summary: CONNECT_HINT, structured: { account_count: 0, accounts: [] } };
  }

  const rows: ListedAccount[] = accounts.map((a) => ({
    account: a.displayId,
    credential_source: a.credentialSourceId,
  }));
  const summary =
    `${rows.length} account(s) connected:\n` +
    rows.map((r) => `- ${r.account} (via ${r.credential_source})`).join("\n");

  return { summary, structured: { account_count: rows.length, accounts: rows } };
}
