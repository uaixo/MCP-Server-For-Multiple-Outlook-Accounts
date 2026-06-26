/**
 * Search-query translation (FR-C2-1, provider-mapping §7-item-2).
 *
 * Re-expresses a documented subset of Gmail-style operators as Microsoft Graph
 * OData `$filter` clauses and/or a KQL `$search` string:
 *
 *   is:unread / is:read   -> isRead eq false / true            ($filter)
 *   from:<addr>           -> from/emailAddress/address eq '…'   ($filter)
 *   to:<addr>             -> toRecipients/any(r: … eq '…')      ($filter)
 *   has:attachment        -> hasAttachments eq true             ($filter)
 *   after: / before:<dt>  -> receivedDateTime ge / le <iso>     ($filter)
 *   subject:<text>, words -> free-text keywords                 ($search)
 *
 * Graph cannot combine `$search` and `$filter` in one request, so a query that
 * mixes free-text/subject search with structured filters is rejected with an
 * actionable message (FR-ERR-3). Unsupported operators are likewise rejected.
 */

export class SearchQueryError extends Error {
  constructor(message: string) {
    super(message);
    this.name = "SearchQueryError";
  }
}

export interface TranslatedQuery {
  /** OData `$filter` expression, when the query used structured operators. */
  readonly filter?: string;
  /** KQL `$search` value (unquoted), when the query used free-text/subject terms. */
  readonly search?: string;
}

interface Token {
  readonly key?: string;
  readonly value: string;
}

const TOKEN_RE = /(\w+):(?:"([^"]*)"|(\S+))|"([^"]*)"|(\S+)/g;

function tokenize(query: string): Token[] {
  const tokens: Token[] = [];
  for (const m of query.matchAll(TOKEN_RE)) {
    if (m[1] !== undefined) {
      tokens.push({ key: m[1].toLowerCase(), value: m[2] ?? m[3] ?? "" });
    } else {
      const value = m[4] ?? m[5] ?? "";
      if (value) tokens.push({ value });
    }
  }
  return tokens;
}

/** Escape an OData string literal: strip CR/LF and double embedded single quotes. */
function odataString(value: string): string {
  return value.replace(/[\r\n]+/g, " ").replace(/'/g, "''");
}

/** Validate and normalise a date operand to an ISO-8601 instant for OData. */
function odataDate(value: string): string {
  const v = value.trim();
  if (/^\d{4}-\d{2}-\d{2}$/.test(v)) {
    const iso = `${v}T00:00:00Z`;
    if (!Number.isNaN(Date.parse(iso))) return iso;
  } else {
    const t = Date.parse(v);
    if (!Number.isNaN(t)) return new Date(t).toISOString();
  }
  throw new SearchQueryError(`Invalid date "${value}". Use YYYY-MM-DD or an ISO-8601 timestamp.`);
}

/** Sanitise a term for the KQL `$search` value (which the caller wraps in double quotes). */
function kqlTerm(value: string): string {
  return value.replace(/["\r\n]+/g, " ").trim();
}

export function translate(query: string): TranslatedQuery {
  const filterClauses: string[] = [];
  const searchTerms: string[] = [];

  for (const t of tokenize(query)) {
    if (t.key === undefined) {
      const term = kqlTerm(t.value);
      if (term) searchTerms.push(term);
      continue;
    }
    switch (t.key) {
      case "is": {
        const v = t.value.toLowerCase();
        if (v === "unread") filterClauses.push("isRead eq false");
        else if (v === "read") filterClauses.push("isRead eq true");
        else
          throw new SearchQueryError(
            `Unsupported "is:" value "${t.value}". Use is:read or is:unread.`,
          );
        break;
      }
      case "from":
        filterClauses.push(`from/emailAddress/address eq '${odataString(t.value)}'`);
        break;
      case "to":
        filterClauses.push(
          `toRecipients/any(r:r/emailAddress/address eq '${odataString(t.value)}')`,
        );
        break;
      case "has": {
        const v = t.value.toLowerCase();
        if (v === "attachment" || v === "attachments") filterClauses.push("hasAttachments eq true");
        else
          throw new SearchQueryError(`Unsupported "has:" value "${t.value}". Use has:attachment.`);
        break;
      }
      case "after":
        filterClauses.push(`receivedDateTime ge ${odataDate(t.value)}`);
        break;
      case "before":
        filterClauses.push(`receivedDateTime le ${odataDate(t.value)}`);
        break;
      case "subject": {
        const term = kqlTerm(t.value);
        if (term) searchTerms.push(`subject:${term}`);
        break;
      }
      default:
        throw new SearchQueryError(
          `Unsupported search operator "${t.key}:". Supported: from, to, subject, is, has, after, before.`,
        );
    }
  }

  const filter = filterClauses.length ? filterClauses.join(" and ") : undefined;
  const search = searchTerms.length ? searchTerms.join(" ") : undefined;

  if (filter && search) {
    throw new SearchQueryError(
      `Microsoft Graph can't combine free-text/subject search with structured filters ` +
        `(is/from/to/has/after/before) in one query. Use one or the other.`,
    );
  }
  return { filter, search };
}
