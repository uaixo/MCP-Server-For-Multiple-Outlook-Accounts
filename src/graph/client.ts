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
import { errorFromResponse, errorFromThrown } from "./errors.js";
import { withRetry, type RetryOptions } from "./retry.js";

export const GRAPH_BASE = "https://graph.microsoft.com/v1.0";

/** Returns a valid access token for an account (refreshing as needed). */
export type TokenProvider = (account: Account) => Promise<string>;

export interface GraphClientDeps {
  readonly requestTimeoutMs: number;
  readonly getToken: TokenProvider;
  /** Injectable for tests; defaults to the global `fetch`. */
  readonly fetchImpl?: typeof fetch;
  readonly retry?: RetryOptions;
}

function buildUrl(req: GraphRequest): string {
  const baseUrl = req.path.startsWith("http") ? req.path : `${GRAPH_BASE}${req.path}`;
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
      res = await doFetch(buildUrl(req), {
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
    return (text ? JSON.parse(text) : undefined) as T;
  }
}
