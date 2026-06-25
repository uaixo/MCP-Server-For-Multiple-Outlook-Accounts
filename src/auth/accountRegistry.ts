/**
 * Account selection — the single chokepoint enforcing the spec §7 selection
 * rule (FR-ID-1..4). Every tool except `list_accounts` resolves its optional
 * account selector here before doing any work.
 *
 * Identity is matched case-insensitively (FR-ID-4). Errors are actionable
 * (FR-ERR-1/3) and carry the list of connected accounts so the caller can
 * recover without guessing.
 */

import type { Account } from "../domain/types.js";
import type { AccountRegistry, TokenStore } from "../domain/contracts.js";

/** Thrown when a selector cannot be resolved; surfaced to the host as a tool error. */
export class AccountSelectionError extends Error {
  constructor(message: string) {
    super(message);
    this.name = "AccountSelectionError";
  }
}

function connectedList(accounts: Account[]): string {
  return accounts.map((a) => a.displayId).join(", ");
}

export class FileAccountRegistry implements AccountRegistry {
  constructor(private readonly store: TokenStore) {}

  list(): Promise<Account[]> {
    return this.store.list();
  }

  async resolve(selector?: string): Promise<Account> {
    const accounts = await this.store.list();

    // Explicit selector: must match a connected account (case-insensitive).
    if (selector !== undefined && selector.trim() !== "") {
      const match = accounts.find((a) => a.id === selector.trim().toLowerCase());
      if (!match) {
        throw new AccountSelectionError(
          accounts.length === 0
            ? `Unknown account "${selector}". No accounts are connected — run "outlook-mcp-auth connect".`
            : `Unknown account "${selector}". Connected accounts: ${connectedList(accounts)}.`,
        );
      }
      return match;
    }

    // Default rule (FR-ID-2).
    if (accounts.length === 0) {
      throw new AccountSelectionError(
        `No accounts are connected. Run "outlook-mcp-auth connect" to add a mailbox.`,
      );
    }
    if (accounts.length === 1) {
      return accounts[0]!;
    }
    throw new AccountSelectionError(
      `Several accounts are connected; specify one via "account". ` +
        `Connected accounts: ${connectedList(accounts)}.`,
    );
  }
}
