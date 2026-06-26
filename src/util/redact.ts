/**
 * Secret redaction for log/error boundaries (NFR-SEC-6).
 *
 * The server is careful never to log tokens, credentials, or message content —
 * but error objects bubbling up from MSAL, `fetch`, or the token cache can carry
 * such material in their messages. Rather than trust every call site, we scrub
 * known secret shapes from any string just before it is written to stderr. This
 * makes "never log secrets" an enforced boundary, not only a convention.
 *
 * `redact` is conservative: it targets the specific shapes that actually leak
 * (Bearer headers, JWT access tokens, and `*_token` / `client_secret` /
 * `password` assignments) so ordinary diagnostic text is left readable.
 */

const REDACTED = "[redacted]";

/** A 3-part JWT (Microsoft access/id tokens are JWTs, beginning `eyJ`). */
const JWT = /eyJ[A-Za-z0-9_-]+\.[A-Za-z0-9_-]+\.[A-Za-z0-9_-]+/g;

/** An `Authorization: Bearer <token>` value. */
const BEARER = /Bearer\s+[A-Za-z0-9._~+/=-]+/gi;

/**
 * A secret carried as `key: "value"`, `key=value`, or `"key":"value"` in JSON,
 * query strings, or free text. Keeps the key, redacts the value.
 */
const KEYED_SECRET =
  /(["']?(?:access_token|refresh_token|id_token|client_secret|client-secret|password|authorization)["']?\s*[:=]\s*)(["']?)[^"'\s,&}]+\2/gi;

/** Scrub known secret shapes from a string so it is safe to log (NFR-SEC-6). */
export function redact(text: string): string {
  return text
    .replace(BEARER, `Bearer ${REDACTED}`)
    .replace(
      KEYED_SECRET,
      (_m, prefix: string, quote: string) => `${prefix}${quote}${REDACTED}${quote}`,
    )
    .replace(JWT, REDACTED);
}

/** Redact an unknown thrown value's message for logging (NFR-SEC-6). */
export function redactError(err: unknown): string {
  return redact(err instanceof Error ? err.message : String(err));
}
