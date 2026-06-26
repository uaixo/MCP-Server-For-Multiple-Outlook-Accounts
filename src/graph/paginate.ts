/**
 * Follow Graph's `@odata.nextLink` to collect a list across pages, bounded by a
 * hard item cap so a pathologically large collection can't produce unbounded
 * work or output. The cap is reported via `truncated`.
 */

import type { Account, GraphClient, GraphRequest } from "../domain/contracts.js";
import type { GraphListResponse } from "./types.js";

export async function collectPaged<T>(
  graph: GraphClient,
  account: Account,
  first: GraphRequest,
  maxItems: number,
): Promise<{ items: T[]; truncated: boolean }> {
  const items: T[] = [];
  let req: GraphRequest | undefined = first;
  let truncated = false;

  while (req) {
    const page: GraphListResponse<T> = await graph.request<GraphListResponse<T>>(account, req);
    for (const item of page.value) {
      if (items.length >= maxItems) {
        truncated = true;
        break;
      }
      items.push(item);
    }
    const next: string | undefined = page["@odata.nextLink"];
    // A nextLink is an absolute URL; replay it as-is with the same retry class.
    req =
      !truncated && next ? { method: "GET", path: next, retryClass: first.retryClass } : undefined;
  }

  return { items, truncated };
}
