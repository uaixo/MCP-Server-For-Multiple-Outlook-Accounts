/**
 * Subsystem contracts (interfaces) — the architectural seams of the server.
 *
 * These declare WHAT each subsystem must do and cite the requirement IDs they
 * satisfy. They contain no logic; the build phase implements them. Keeping the
 * seams explicit lets the capability layer be unit-tested against mocked Graph
 * and MSAL (the chosen test strategy).
 *
 * STATUS: scaffold — contracts only.
 */

import type {
  Account,
  AttachmentInput,
  ComposeInput,
  ConversationDetail,
  ConversationSummary,
  OrganisationLabel,
  ResultCursor,
} from "./types.js";

/** A discovered Entra app-registration credential source (FR-ID-5/6, provider-mapping §4). */
export interface CredentialSource {
  /** Stable id (e.g. config filename) recorded against each account it authorises. */
  readonly id: string;
  readonly clientId: string;
  /** Entra tenant ("common", "organizations", or a specific tenant id). */
  readonly tenant: string;
  /** Least-privilege scopes (Mail.ReadWrite, Mail.Send, User.Read, offline_access) (FR-AUTH-10). */
  readonly scopes: string[];
}

/**
 * Persistent, secure store of per-account refresh tokens / MSAL caches.
 * Writes are atomic (temp + rename) and guarded by a cross-process lock with
 * stale-lock recovery (NFR-SEC-2). The file is 600 inside a 700 dir (NFR-SEC-1).
 * A corrupt store is treated as "no accounts" with a one-time warning (FR-ERR-2).
 */
export interface TokenStore {
  /** List accounts known to the store (keyed lower-cased) (FR-C1-1, FR-ID-4). */
  list(): Promise<Account[]>;
  /** Persist/replace one account's serialized MSAL cache, atomically + locked. */
  upsert(account: Account, serializedCache: string): Promise<void>;
  /** Remove an account (CLI `remove`, FR-AUTH-8). */
  remove(accountId: string): Promise<void>;
  /** Read a serialized MSAL cache for refresh; re-read on each call so re-consent is picked up without restart (FR-AUTH-9). */
  readCache(accountId: string): Promise<string | undefined>;
}

/**
 * Resolves the optional account selector to a concrete account per the
 * selection rule (spec §7): default when exactly one; disambiguation error when
 * several; connect-guidance error when none; validation error listing accounts
 * when an unknown account is named (FR-ID-2/3). Matching is case-insensitive
 * (FR-ID-4).
 */
export interface AccountRegistry {
  list(): Promise<Account[]>;
  resolve(selector?: string): Promise<Account>;
}

/** Retry classification for a Graph call (NFR-REL-2/3). */
export type RetryClass =
  /** Read / organise: may retry on 429, transient 5xx, transport errors. */
  | "safe"
  /** Send / draft-create: retry ONLY on pre-processing 429 — never on ambiguous failures that may have already delivered (NFR-REL-3). */
  | "nonDuplicable";

export interface GraphRequest {
  readonly method: "GET" | "POST" | "PATCH" | "DELETE";
  /** Path relative to the Graph base (e.g. `/me/messages`), OR an absolute Graph URL (e.g. an `@odata.nextLink`). */
  readonly path: string;
  readonly query?: Record<string, string | number | boolean | undefined>;
  readonly body?: unknown;
  /** Extra request headers (e.g. `ConsistencyLevel: eventual` for `$count`/`$search`). */
  readonly headers?: Record<string, string>;
  readonly retryClass: RetryClass;
}

/**
 * Thin wrapper over `fetch` to Microsoft Graph (chosen over the official SDK
 * for precise control of the no-duplicate-send retry rule). Applies the
 * per-request timeout via AbortSignal (NFR-REL-1), bounded jittered backoff
 * honouring `Retry-After` (NFR-REL-2), and the {@link RetryClass} policy
 * (NFR-REL-3). Maps Graph errors to actionable categories (FR-ERR-1).
 */
export interface GraphClient {
  request<T>(account: Account, req: GraphRequest): Promise<T>;
}

/** Caps a fan-out of per-item Graph calls to respect rate limits (NFR-REL-4, FR-C2-6). */
export interface ConcurrencyLimiter {
  run<T>(tasks: Array<() => Promise<T>>): Promise<T[]>;
}

/**
 * Translates the supported search-operator subset (Gmail-style, e.g.
 * `is:unread`, `from:x`, `has:attachment`) into Graph `$search` + OData
 * `$filter` (provider-mapping §7 item 2, FR-C2-1). Unsupported operators are
 * rejected or translated predictably.
 */
export interface SearchTranslator {
  translate(query: string): { search?: string; filter?: string };
}

/** A single Graph operation produced by decomposing a neutral organise intent. */
export interface GraphOperation {
  readonly description: string;
  readonly request: GraphRequest;
}

/**
 * Decomposes ONE neutral organise request into the correct combination of Graph
 * operations — categories[] (tag), move (folder/archive/junk), isRead
 * (read-state). This is the core porting risk (spec §4 concept-decomposition
 * rule, FR-C8-6; provider-mapping §3.1 / §7 item 1).
 */
export interface OrganiseDecomposer {
  decompose(intent: OrganiseIntent): GraphOperation[];
}

/** Add/remove labels plus derived intents; at least one change required (FR-C8-2/3). */
export interface OrganiseIntent {
  readonly addLabelIds?: string[];
  readonly removeLabelIds?: string[];
  readonly markRead?: boolean;
  readonly archive?: boolean;
}

/** Reads attachment bytes safely: path reads only within the allow-list, opened once and validated via the handle (NFR-SEC-3/4). */
export interface AttachmentReader {
  read(input: AttachmentInput): Promise<{ filename: string; mimeType: string; bytes: Uint8Array }>;
}

/**
 * Re-exports for convenience so capability modules import a single contracts
 * surface. (Capabilities C1–C8 will be added under src/capabilities/.)
 */
export type {
  Account,
  ComposeInput,
  ConversationDetail,
  ConversationSummary,
  OrganisationLabel,
  ResultCursor,
};
