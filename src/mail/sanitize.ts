/**
 * Header-injection defence for outgoing mail (NFR-SEC-5).
 *
 * Values that end up in mail headers — recipient display names/addresses,
 * subject, attachment filenames — must never carry CR/LF (or other control
 * characters), or a crafted display name / subject could inject additional
 * headers ("header splitting"). These helpers strip those characters before a
 * value is placed into the outgoing Graph message.
 *
 * We do NOT depend on Graph's structured JSON happening to be injection-safe:
 * stripping at the boundary keeps the guarantee local and testable, and holds
 * even if a future code path emits MIME headers directly.
 */

/** True for a line break: CR, LF, or the Unicode line/paragraph separators. */
function isLineBreak(code: number): boolean {
  return code === 0x0a || code === 0x0d || code === 0x2028 || code === 0x2029;
}

/** True for a C0/C1 control or DEL that must never sit in a header. Tab is allowed. */
function isStrippableControl(code: number): boolean {
  if (code === 0x09) return false; // keep tab
  return code <= 0x1f || code === 0x7f;
}

/**
 * Strip CR/LF (and other control characters) from a value bound for a mail
 * header. Line breaks collapse to a single space so adjacent tokens don't merge
 * into one; the remaining controls are removed; the result is trimmed.
 * (NFR-SEC-5)
 */
export function sanitizeHeaderValue(value: string): string {
  let out = "";
  let pendingSpace = false;
  for (const ch of value) {
    const code = ch.codePointAt(0)!;
    if (isLineBreak(code)) {
      pendingSpace = out.length > 0; // collapse a run of breaks into one space
      continue;
    }
    if (isStrippableControl(code)) continue;
    if (pendingSpace) {
      out += " ";
      pendingSpace = false;
    }
    out += ch;
  }
  return out.trim();
}

/**
 * Sanitize an attachment filename: strip header-injecting characters and drop
 * any path separators so the name can't smuggle a directory traversal or a
 * second header. (NFR-SEC-5)
 */
export function sanitizeFilename(name: string): string {
  return sanitizeHeaderValue(name).replace(/[/\\]+/g, "_");
}
