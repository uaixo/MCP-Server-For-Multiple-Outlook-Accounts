/**
 * Microsoft Graph error mapping (FR-ERR-1).
 *
 * Maps HTTP responses and transport failures to a small set of actionable
 * categories with human-readable messages, never leaking secrets (NFR-SEC-6).
 * The {@link GraphError.category} drives the retry policy (see retry.ts).
 */

export type GraphErrorCategory =
  | "auth" // 401/403 — token expired/revoked or insufficient consent
  | "rateLimited" // 429 — throttled before processing (safe to retry)
  | "timeout" // request aborted by the per-request timeout
  | "transport" // network/DNS/connection failure
  | "server" // 5xx — transient server error
  | "client" // other 4xx — caller error, not retryable
  | "unknown";

export class GraphError extends Error {
  constructor(
    message: string,
    readonly category: GraphErrorCategory,
    readonly status?: number,
    /** Honoured backoff from a `Retry-After` header, in milliseconds. */
    readonly retryAfterMs?: number,
  ) {
    super(message);
    this.name = "GraphError";
  }
}

export function categorizeStatus(status: number): GraphErrorCategory {
  if (status === 401 || status === 403) return "auth";
  if (status === 429) return "rateLimited";
  if (status === 408) return "timeout";
  if (status >= 500) return "server";
  if (status >= 400) return "client";
  return "unknown";
}

/** Parse a `Retry-After` header (delta-seconds or HTTP-date) into milliseconds. */
export function parseRetryAfter(
  value: string | null,
  now: number = Date.now(),
): number | undefined {
  if (!value) return undefined;
  const secs = Number(value);
  if (Number.isFinite(secs)) return Math.max(0, secs * 1000);
  const when = Date.parse(value);
  return Number.isNaN(when) ? undefined : Math.max(0, when - now);
}

function actionableMessage(category: GraphErrorCategory, status: number, detail?: string): string {
  switch (category) {
    case "auth":
      return `Authentication failed (HTTP ${status}). Re-connect this account with "outlook-mcp-auth connect".`;
    case "rateLimited":
      return `Microsoft Graph is rate limiting requests (HTTP 429). Try again shortly.`;
    case "server":
      return `Microsoft Graph had a transient error (HTTP ${status}). Try again shortly.`;
    case "client":
      return `Microsoft Graph rejected the request (HTTP ${status})${detail ? `: ${detail}` : ""}.`;
    default:
      return `Microsoft Graph request failed (HTTP ${status})${detail ? `: ${detail}` : ""}.`;
  }
}

/** Extract the `error.message` from a Graph error body without throwing. */
function extractGraphMessage(bodyText: string): string | undefined {
  try {
    const parsed = JSON.parse(bodyText) as { error?: { message?: unknown } };
    const msg = parsed.error?.message;
    return typeof msg === "string" && msg.trim() !== "" ? msg.trim() : undefined;
  } catch {
    return undefined;
  }
}

/** Build a {@link GraphError} from a non-OK `Response`. */
export async function errorFromResponse(res: Response): Promise<GraphError> {
  const category = categorizeStatus(res.status);
  const retryAfterMs =
    category === "rateLimited" || category === "server"
      ? parseRetryAfter(res.headers.get("retry-after"))
      : undefined;
  let detail: string | undefined;
  try {
    detail = extractGraphMessage(await res.text());
  } catch {
    detail = undefined;
  }
  return new GraphError(
    actionableMessage(category, res.status, detail),
    category,
    res.status,
    retryAfterMs,
  );
}

/** Map a thrown fetch/transport error to a {@link GraphError}. */
export function errorFromThrown(err: unknown): GraphError {
  if (err instanceof GraphError) return err;
  const name =
    typeof err === "object" && err !== null && "name" in err
      ? String((err as { name?: unknown }).name)
      : "";
  if (name === "TimeoutError" || name === "AbortError") {
    return new GraphError("The request to Microsoft Graph timed out.", "timeout");
  }
  return new GraphError("Could not reach Microsoft Graph (network error).", "transport");
}
