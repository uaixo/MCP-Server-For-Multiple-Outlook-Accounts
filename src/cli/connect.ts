/**
 * CLI `connect` — onboard a new mailbox via interactive browser consent
 * (FR-AUTH-1..7,10). The MCP server never initiates consent; this CLI does.
 *
 * Flow (see architecture.md §5.2):
 *   pick app registration → MSAL interactive (PKCE + state + loopback) →
 *   identify account → persist cache bound to the issuing registration.
 *
 * Re-running `connect` for an existing account repairs a revoked refresh token;
 * a running server picks up the rewritten cache on its next call (FR-AUTH-9),
 * since the token store is re-read per access.
 */

import { loadConfig, type Config } from "../config.js";
import { loadCredentialSources, getCredentialSource } from "../auth/credentialSources.js";
import { interactiveConnect, CONSENT_TIMEOUT_MS } from "../auth/msalClient.js";
import { FileTokenStore } from "../auth/tokenStore.js";
import type { CredentialSource } from "../domain/contracts.js";

interface ConnectArgs {
  /** Pin a specific app-registration config file path. */
  credentials?: string;
  /** Select a discovered source by id when several exist. */
  source?: string;
}

function parseArgs(argv: string[]): ConnectArgs {
  const args: ConnectArgs = {};
  for (let i = 0; i < argv.length; i++) {
    const a = argv[i];
    if (a === "--credentials" || a === "-c") args.credentials = argv[++i];
    else if (a === "--source" || a === "-s") args.source = argv[++i];
  }
  return args;
}

/** Choose the app registration to authorise under (FR-ID-5/6). */
async function selectSource(config: Config, args: ConnectArgs): Promise<CredentialSource> {
  if (args.credentials) {
    const pinned: Config = { ...config, oauthCredentialsPath: args.credentials };
    const [source] = await loadCredentialSources(pinned);
    if (!source) throw new Error(`No usable credential config at ${args.credentials}.`);
    return source;
  }

  const sources = await loadCredentialSources(config);
  if (sources.length === 0) {
    throw new Error(
      `No app-registration config found in ${config.dataDir}. ` +
        `Add a "credentials*.json" with { "clientId": "...", "tenant": "..." } ` +
        `or pass --credentials <path>.`,
    );
  }
  if (args.source) {
    const chosen = await getCredentialSource(config, args.source);
    if (!chosen) {
      throw new Error(
        `No credential source "${args.source}". Available: ${sources.map((s) => s.id).join(", ")}.`,
      );
    }
    return chosen;
  }
  if (sources.length > 1) {
    throw new Error(
      `Several app registrations found; choose one with --source <id>. ` +
        `Available: ${sources.map((s) => s.id).join(", ")}.`,
    );
  }
  return sources[0]!;
}

export async function runConnect(argv: string[], config: Config = loadConfig()): Promise<number> {
  const args = parseArgs(argv);
  let source: CredentialSource;
  try {
    source = await selectSource(config, args);
  } catch (e) {
    process.stderr.write(`${e instanceof Error ? e.message : String(e)}\n`);
    return 1;
  }

  process.stderr.write(
    `Opening browser to connect a mailbox via app registration "${source.id}"...\n`,
  );

  try {
    const { identity, serializedCache } = await interactiveConnect(source, {
      timeoutMs: CONSENT_TIMEOUT_MS,
    });
    const store = new FileTokenStore({
      dataDir: config.dataDir,
      lockTimeoutMs: config.lockTimeoutMs,
    });
    await store.upsert(
      { id: identity.toLowerCase(), displayId: identity, credentialSourceId: source.id },
      serializedCache,
    );
    process.stdout.write(`Connected ${identity} (via ${source.id}).\n`);
    return 0;
  } catch (e) {
    process.stderr.write(`Failed to connect: ${e instanceof Error ? e.message : String(e)}\n`);
    return 1;
  }
}
