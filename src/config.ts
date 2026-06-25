/**
 * Operational configuration loaded from environment variables.
 *
 * Maps the neutral knobs (spec §12) to the OUTLOOK_* prefix (provider-mapping
 * §5), preserving each knob's meaning. Defaults follow the reference
 * implementation values cited in the spec.
 *
 * This module is intentionally pure (env in → config out) so it is unit-
 * testable without touching the filesystem or network.
 */

import { homedir } from "node:os";
import { join } from "node:path";

export interface Config {
  /** Where tokens + app-registration configs live. Default: ~/.outlook-mcp (NFR-SEC-1). */
  readonly dataDir: string;
  /** Pinned single app-registration config path; undefined = auto-discover (FR-ID-6). */
  readonly oauthCredentialsPath: string | undefined;
  /**
   * Directories that `path` attachments may be read from. Empty = path
   * attachments DISABLED; inline base64 still works (NFR-SEC-3).
   */
  readonly attachmentsAllowList: readonly string[];
  /** Max wait (ms) for the token-store lock before failing a write (NFR-SEC-2). */
  readonly lockTimeoutMs: number;
  /** Per-request timeout (ms) bounding every Graph call (NFR-REL-1). */
  readonly requestTimeoutMs: number;
}

/** Reference defaults cited in the spec / provider-mapping §5. */
export const DEFAULTS = {
  lockTimeoutMs: 12_000,
  requestTimeoutMs: 30_000,
} as const;

export const ENV_KEYS = {
  dataDir: "OUTLOOK_MCP_DATA_DIR",
  oauthCredentials: "OUTLOOK_OAUTH_CREDENTIALS",
  attachmentsDir: "OUTLOOK_MCP_ATTACHMENTS_DIR",
  lockTimeoutMs: "OUTLOOK_MCP_LOCK_TIMEOUT_MS",
  requestTimeoutMs: "OUTLOOK_MCP_REQUEST_TIMEOUT_MS",
} as const;

function parsePositiveInt(raw: string | undefined, fallback: number, key: string): number {
  if (raw === undefined || raw.trim() === "") return fallback;
  const n = Number(raw);
  if (!Number.isInteger(n) || n <= 0) {
    throw new Error(`${key} must be a positive integer (got ${JSON.stringify(raw)})`);
  }
  return n;
}

function parseAllowList(raw: string | undefined, sep: string): string[] {
  if (raw === undefined || raw.trim() === "") return [];
  return raw
    .split(sep)
    .map((p) => p.trim())
    .filter((p) => p.length > 0);
}

/**
 * Build the {@link Config} from an environment map (defaults to `process.env`)
 * and the OS path-list separator (defaults to the platform's). Both are
 * injectable so tests stay deterministic and cross-platform (NFR-OPS-1).
 */
export function loadConfig(
  env: NodeJS.ProcessEnv = process.env,
  pathListSeparator: string = process.platform === "win32" ? ";" : ":",
): Config {
  const dataDir = env[ENV_KEYS.dataDir]?.trim() || join(homedir(), ".outlook-mcp");
  const oauthCredentialsPath = env[ENV_KEYS.oauthCredentials]?.trim() || undefined;

  return {
    dataDir,
    oauthCredentialsPath,
    attachmentsAllowList: parseAllowList(env[ENV_KEYS.attachmentsDir], pathListSeparator),
    lockTimeoutMs: parsePositiveInt(
      env[ENV_KEYS.lockTimeoutMs],
      DEFAULTS.lockTimeoutMs,
      ENV_KEYS.lockTimeoutMs,
    ),
    requestTimeoutMs: parsePositiveInt(
      env[ENV_KEYS.requestTimeoutMs],
      DEFAULTS.requestTimeoutMs,
      ENV_KEYS.requestTimeoutMs,
    ),
  };
}
