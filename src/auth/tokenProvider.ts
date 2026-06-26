/**
 * Bridges the token store + credential sources + MSAL into the {@link TokenProvider}
 * the Graph client needs (FR-ID-5, FR-AUTH-9).
 *
 * For each request it loads the account's bound app registration, reads its
 * serialized MSAL cache from the store, silently acquires an access token, and
 * persists the cache back when it changed (e.g. a rotated refresh token) so a
 * re-consent is picked up without a server restart.
 */

import type { Config } from "../config.js";
import type { TokenStore } from "../domain/contracts.js";
import type { TokenProvider } from "../graph/client.js";
import { getCredentialSource } from "./credentialSources.js";
import { acquireToken } from "./msalClient.js";

export interface TokenProviderDeps {
  readonly config: Config;
  readonly tokenStore: TokenStore;
}

export function createMsalTokenProvider(deps: TokenProviderDeps): TokenProvider {
  return async (account) => {
    const source = await getCredentialSource(deps.config, account.credentialSourceId);
    if (!source) {
      throw new Error(
        `Credential source "${account.credentialSourceId}" for ${account.displayId} is missing; ` +
          `re-connect this account with "outlook-mcp-auth connect".`,
      );
    }
    const cache = await deps.tokenStore.readCache(account.id);
    if (cache === undefined) {
      throw new Error(
        `No stored credentials for ${account.displayId}; re-connect with "outlook-mcp-auth connect".`,
      );
    }

    const { accessToken, serializedCache } = await acquireToken(account, source, cache, {
      timeoutMs: deps.config.requestTimeoutMs,
    });
    if (serializedCache !== cache) {
      await deps.tokenStore.upsert(account, serializedCache);
    }
    return accessToken;
  };
}
