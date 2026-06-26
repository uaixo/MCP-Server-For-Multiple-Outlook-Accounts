/**
 * Thin Microsoft Graph client over `fetch` (provider-mapping §6, §7-item-4).
 *
 * The single outbound HTTP path. Applies:
 * - a per-request timeout via `AbortSignal` (NFR-REL-1),
 * - the bounded/jittered retry policy with the no-duplicate-send guarantee
 *   (NFR-REL-2/3, via {@link withRetry}),
 * - actionable error mapping (FR-ERR-1, via errors.ts).
 *
 * Access tokens are obtained through an injected {@link TokenProvider} so the
 * HTTP layer is independent of MSAL and unit-testable without it.
 */

import type { Account, GraphClient, GraphRequest } from "../domain/contracts.js";
import { errorFromResponse, errorFromThrown, GraphError } from "./errors.js";
import { withRetry, type RetryOptions } from "./retry.js";

export const GRAPH_BASE = "https://graph.microsoft.com/v1.0";

/** The only origin we will ever send an access token to (see {@link buildUrl}). */
const GRAPH_ORIGIN = new URL(GRAPH_BASE).origin;

/** Returns a valid access token for an account (refreshing as needed). */
export type TokenProvider = (account: Account) => Promise<string>;

export interface GraphClientDeps {
  readonly requestTimeoutMs: number;
  readonly getToken: TokenProvider;
  /** Injectable for tests; defaults to the global `fetch`. */
  readonly fetchImpl?: typeof fetch;
  readonly retry?: RetryOptions;
}

/**
 * Resolve a request to an absolute URL, **pinned to the Microsoft Graph origin**.
 *
 * Relative paths are joined to {@link GRAPH_BASE}. An absolute path (e.g. an
 * `@odata.nextLink` pagination cursor, or a caller-supplied `page_token`) is only
 * accepted when its origin is Graph's. Without this, a crafted absolute URL would
 * be fetched with the `Authorization: Bearer …` header attached — leaking the
 * access token to an attacker-controlled host (SSRF / token exfiltration).
 */
function buildUrl(req: GraphRequest): string {
  let baseUrl: string;
  if (req.path.startsWith("http")) {
    let parsed: URL;
    try {
      parsed = new URL(req.path);
    } catch {
      throw new Error("Invalid Microsoft Graph request URL.");
    }
    if (parsed.origin !== GRAPH_ORIGIN) {
      throw new Error(
        `Refusing to send a request to a non-Microsoft-Graph host (${parsed.origin}).`,
      );
    }
    baseUrl = req.path;
  } else {
    baseUrl = `${GRAPH_BASE}${req.path}`;
  }
  if (!req.query) return baseUrl;
  const url = new URL(baseUrl);
  for (const [key, value] of Object.entries(req.query)) {
    if (value !== undefined) url.searchParams.set(key, String(value));
  }
  return url.toString();
}

export class FetchGraphClient implements GraphClient {
  constructor(private readonly deps: GraphClientDeps) {}

  request<T>(account: Account, req: GraphRequest): Promise<T> {
    return withRetry(req.retryClass, () => this.attempt<T>(account, req), this.deps.retry);
  }

  private async attempt<T>(account: Account, req: GraphRequest): Promise<T> {
    // Resolve + validate the URL BEFORE fetching a token or hitting the network,
    // so an off-origin URL is rejected outright (not mapped to a retryable
    // transport error, and without ever minting a token for it).
    const url = buildUrl(req);
    const token = await this.deps.getToken(account);
    const doFetch = this.deps.fetchImpl ?? fetch;

    const headers: Record<string, string> = {
      Authorization: `Bearer ${token}`,
      Accept: "application/json",
      ...req.headers,
    };
    if (req.body !== undefined) headers["Content-Type"] = "application/json";

    let res: Response;
    try {
      res = await doFetch(url, {
        method: req.method,
        headers,
        body: req.body !== undefined ? JSON.stringify(req.body) : undefined,
        signal: AbortSignal.timeout(this.deps.requestTimeoutMs), // NFR-REL-1
      });
    } catch (e) {
      throw errorFromThrown(e);
    }

    if (!res.ok) throw await errorFromResponse(res);
    // Empty-body successes: 204 No Content (e.g. PATCH) and 202 Accepted
    // (`POST /me/sendMail`). Parse JSON only when there is a body to parse.
    if (res.status === 204) return undefined as T;
    const text = await res.text();
    if (!text) return undefined as T;
    try {
      return JSON.parse(text) as T;
    } catch {
      // A 2xx with a non-JSON body is unexpected; surface it as a mapped,
      // non-retryable error rather than a raw SyntaxError.
      throw new GraphError("Microsoft Graph returned a malformed response.", "unknown", res.status);
    }
  }
}
