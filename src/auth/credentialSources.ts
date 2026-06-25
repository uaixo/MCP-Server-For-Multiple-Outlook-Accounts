/**
 * Entra app-registration credential sources (FR-ID-5/6, FR-AUTH-10).
 *
 * Supports one or more app registrations:
 * - If `OUTLOOK_OAUTH_CREDENTIALS` is set, that single config file is pinned and
 *   auto-discovery is disabled.
 * - Otherwise every `credentials*.json` in the data dir is auto-discovered, so
 *   accounts under different app registrations (e.g. different organisations)
 *   each refresh with the client that authorised them (FR-ID-5).
 *
 * Scopes are NOT taken from the config file; they are fixed to the
 * least-privilege set required for the supported capabilities (FR-AUTH-10).
 */

import { readFile, readdir } from "node:fs/promises";
import { basename, join } from "node:path";
import type { Config } from "../config.js";
import type { CredentialSource } from "../domain/contracts.js";

/**
 * Least-privilege scopes (FR-AUTH-10, provider-mapping §4). `offline_access` is
 * listed for traceability; MSAL adds the reserved OIDC scopes itself, so it is
 * stripped before the consent request — see {@link consentScopes}.
 */
export const OUTLOOK_SCOPES: readonly string[] = [
  "Mail.ReadWrite",
  "Mail.Send",
  "User.Read",
  "offline_access",
];

const RESERVED_SCOPES = new Set(["openid", "profile", "offline_access"]);

/** Resource scopes to pass to MSAL (reserved OIDC scopes removed; MSAL adds them). */
export function consentScopes(scopes: readonly string[]): string[] {
  return scopes.filter((s) => !RESERVED_SCOPES.has(s));
}

const CREDENTIALS_FILE_RE = /^credentials.*\.json$/i;

interface CredentialConfigFile {
  clientId: string;
  tenant?: string;
  id?: string;
}

function parseConfig(path: string, raw: string): CredentialSource {
  let parsed: unknown;
  try {
    parsed = JSON.parse(raw);
  } catch {
    throw new Error(`Credential config ${path} is not valid JSON.`);
  }
  const cfg = parsed as Partial<CredentialConfigFile>;
  if (typeof cfg.clientId !== "string" || cfg.clientId.trim() === "") {
    throw new Error(`Credential config ${path} is missing a non-empty "clientId".`);
  }
  return {
    id:
      typeof cfg.id === "string" && cfg.id.trim() !== ""
        ? cfg.id
        : basename(path).replace(/\.json$/i, ""),
    clientId: cfg.clientId,
    tenant: typeof cfg.tenant === "string" && cfg.tenant.trim() !== "" ? cfg.tenant : "common",
    scopes: [...OUTLOOK_SCOPES],
  };
}

async function loadOne(path: string): Promise<CredentialSource> {
  return parseConfig(path, await readFile(path, "utf8"));
}

/** Discover all configured app registrations per the pinned-or-auto-discover rule (FR-ID-6). */
export async function loadCredentialSources(config: Config): Promise<CredentialSource[]> {
  if (config.oauthCredentialsPath) {
    return [await loadOne(config.oauthCredentialsPath)];
  }
  let entries: string[];
  try {
    entries = await readdir(config.dataDir);
  } catch {
    return []; // data dir not created yet — no sources.
  }
  const files = entries.filter((f) => CREDENTIALS_FILE_RE.test(f)).sort();
  return Promise.all(files.map((f) => loadOne(join(config.dataDir, f))));
}

/** Resolve a single source by id, binding an account's refresh to its issuing client (FR-ID-5). */
export async function getCredentialSource(
  config: Config,
  id: string,
): Promise<CredentialSource | undefined> {
  const sources = await loadCredentialSources(config);
  return sources.find((s) => s.id === id);
}
