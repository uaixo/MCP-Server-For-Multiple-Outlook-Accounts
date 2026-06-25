/**
 * CLI `remove` — remove a connected account (FR-AUTH-8). Matching is
 * case-insensitive (FR-ID-4).
 */

import { loadConfig, type Config } from "../config.js";
import { FileTokenStore } from "../auth/tokenStore.js";

export async function runRemove(argv: string[], config: Config = loadConfig()): Promise<number> {
  const target = argv[0];
  if (!target || target.trim() === "") {
    process.stderr.write("Usage: outlook-mcp-auth remove <account>\n");
    return 2;
  }

  const store = new FileTokenStore({
    dataDir: config.dataDir,
    lockTimeoutMs: config.lockTimeoutMs,
  });
  const accounts = await store.list();
  const match = accounts.find((a) => a.id === target.trim().toLowerCase());

  if (!match) {
    process.stderr.write(
      accounts.length === 0
        ? `No accounts are connected.\n`
        : `No connected account "${target}". Connected: ${accounts
            .map((a) => a.displayId)
            .join(", ")}.\n`,
    );
    return 1;
  }

  await store.remove(match.id);
  process.stdout.write(`Removed ${match.displayId}.\n`);
  return 0;
}
