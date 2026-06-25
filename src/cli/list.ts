/**
 * CLI `list` — list connected accounts with the credential source each uses
 * (FR-AUTH-8).
 */

import { loadConfig, type Config } from "../config.js";
import { FileTokenStore } from "../auth/tokenStore.js";

export async function runList(config: Config = loadConfig()): Promise<number> {
  const store = new FileTokenStore({
    dataDir: config.dataDir,
    lockTimeoutMs: config.lockTimeoutMs,
  });
  const accounts = await store.list();

  if (accounts.length === 0) {
    process.stdout.write('No accounts connected. Run "outlook-mcp-auth connect" to add one.\n');
    return 0;
  }

  process.stdout.write(`Connected accounts (${accounts.length}):\n`);
  for (const a of accounts) {
    process.stdout.write(`  ${a.displayId}  [source: ${a.credentialSourceId}]\n`);
  }
  return 0;
}
