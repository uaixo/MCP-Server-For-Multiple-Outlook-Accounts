/**
 * MSAL public-client wrapper (provider-mapping §4) — replaces google-auth-library
 * for both the consent flow and silent token refresh.
 *
 * Invariants preserved:
 * - PKCE (S256) and CSRF `state` are handled internally by MSAL's interactive
 *   loopback flow (FR-AUTH-3/4); a forged/unrelated callback is answered without
 *   aborting the genuine flow.
 * - The loopback redirect binds to 127.0.0.1 and MSAL manages the port
 *   (FR-AUTH-2, NFR-SEC-7).
 * - `offline_access` yields a long-lived refresh token held in the MSAL cache
 *   (FR-AUTH-5); the serialized cache is persisted by the token store.
 * - Refresh always uses the SAME app registration that issued the token
 *   (FR-ID-5): the caller supplies the bound {@link CredentialSource}.
 * - The consent wait is bounded (FR-AUTH-7).
 */

import { spawn } from "node:child_process";
import { PublicClientApplication } from "@azure/msal-node";
import type { AccountInfo } from "@azure/msal-node";
import type { Account, CredentialSource } from "../domain/contracts.js";
import { consentScopes } from "./credentialSources.js";

/** Reference consent timeout (FR-AUTH-7): 5 minutes. */
export const CONSENT_TIMEOUT_MS = 5 * 60_000;

function authority(tenant: string): string {
  return `https://login.microsoftonline.com/${tenant}`;
}

function buildApp(source: CredentialSource): PublicClientApplication {
  return new PublicClientApplication({
    auth: { clientId: source.clientId, authority: authority(source.tenant) },
  });
}

/** Open the system browser for the consent URL, cross-platform (NFR-OPS-1). Exported for testing. */
export function defaultOpenBrowser(url: string): Promise<void> {
  const [cmd, args] =
    process.platform === "win32"
      ? ["cmd", ["/c", "start", "", url]]
      : process.platform === "darwin"
        ? ["open", [url]]
        : ["xdg-open", [url]];
  const child = spawn(cmd, args, { detached: true, stdio: "ignore" });
  // The opener binary may be missing (e.g. headless Linux without xdg-open).
  // An unhandled 'error' event would crash the process, so swallow it and print
  // the URL so the user can open it manually; consent then completes normally.
  child.on("error", () => {
    process.stderr.write(
      `Could not launch a browser automatically. Open this URL to continue:\n${url}\n`,
    );
  });
  child.unref();
  return Promise.resolve();
}

async function withTimeout<T>(promise: Promise<T>, ms: number, what: string): Promise<T> {
  let timer: ReturnType<typeof setTimeout>;
  const timeout = new Promise<never>((_, reject) => {
    timer = setTimeout(
      () => reject(new Error(`Timed out after ${ms}ms waiting for ${what}. Please try again.`)),
      ms,
    );
  });
  try {
    return await Promise.race([promise, timeout]);
  } finally {
    clearTimeout(timer!);
  }
}

export interface ConnectResult {
  /** The authenticated identity (UPN / username), keyed lower-cased by the store (FR-AUTH-6, FR-ID-4). */
  readonly identity: string;
  /** Serialized MSAL token cache containing the new refresh token. */
  readonly serializedCache: string;
}

export interface InteractiveConnectOptions {
  readonly timeoutMs?: number;
  /** Injectable for tests; defaults to spawning the system browser. */
  readonly openBrowser?: (url: string) => Promise<void>;
}

/**
 * Run the interactive consent flow for one app registration and return the
 * authenticated identity plus the serialized token cache to persist.
 */
export async function interactiveConnect(
  source: CredentialSource,
  opts: InteractiveConnectOptions = {},
): Promise<ConnectResult> {
  const app = buildApp(source);
  const result = await withTimeout(
    app.acquireTokenInteractive({
      scopes: consentScopes(source.scopes),
      openBrowser: opts.openBrowser ?? defaultOpenBrowser,
      successTemplate:
        "<h2>Account connected.</h2><p>You can close this tab and return to the terminal.</p>",
      errorTemplate:
        "<h2>Connection failed.</h2><p>Please return to the terminal and try again.</p>",
    }),
    opts.timeoutMs ?? CONSENT_TIMEOUT_MS,
    "browser consent",
  );

  const identity = result.account?.username;
  if (!identity) {
    // FR-AUTH-6: abort if the authenticated identity can't be determined.
    throw new Error("Could not determine the authenticated account; aborting.");
  }
  return { identity, serializedCache: app.getTokenCache().serialize() };
}

export interface AccessTokenResult {
  readonly accessToken: string;
  /** Possibly-updated cache (e.g. a rotated refresh token) to persist back. */
  readonly serializedCache: string;
}

/**
 * Silently acquire an access token for an existing account, refreshing via the
 * issuing app registration (FR-ID-5). Used by the Graph layer in later phases.
 * Throws an actionable "re-connect" error when the cache can no longer refresh.
 */
export async function acquireToken(
  account: Account,
  source: CredentialSource,
  serializedCache: string,
  opts: { timeoutMs: number },
): Promise<AccessTokenResult> {
  const app = buildApp(source);
  const cache = app.getTokenCache();
  cache.deserialize(serializedCache);

  const accounts: AccountInfo[] = await cache.getAllAccounts();
  const match = accounts.find((a) => a.username.toLowerCase() === account.id.toLowerCase());
  if (!match) {
    throw new Error(`Account ${account.displayId} is not in the token cache; re-connect it.`);
  }

  const result = await withTimeout(
    app.acquireTokenSilent({ account: match, scopes: consentScopes(source.scopes) }),
    opts.timeoutMs,
    "token refresh",
  );
  return { accessToken: result.accessToken, serializedCache: cache.serialize() };
}
