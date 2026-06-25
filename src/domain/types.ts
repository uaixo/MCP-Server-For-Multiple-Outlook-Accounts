/**
 * Neutral domain model — realises the concepts in business-specification §4 on
 * top of Outlook / Microsoft Graph primitives (see provider-mapping §2).
 *
 * These types are the shared vocabulary the capability layer (§6) is written
 * against. Provider-specific shapes (raw Graph JSON) are intentionally NOT
 * exposed here; they are confined to the graph/ layer and mapped into these.
 *
 * STATUS: scaffold — types are the design contract; capability logic is not yet
 * implemented.
 */

/** One authenticated mailbox identity. (spec §4 "Account"; FR-C1-1) */
export interface Account {
  /** Primary identity string — userPrincipalName / primary SMTP (`mail`). Stored lower-cased as the key (FR-ID-4). */
  readonly id: string;
  /** Display identity exactly as returned by `GET /me` (may differ in case). */
  readonly displayId: string;
  /** Which credential source (Entra app registration) authorised this account (FR-ID-5). */
  readonly credentialSourceId: string;
}

/** A provider grouping of related messages — Outlook `conversationId`. (spec §4 "Conversation") */
export interface ConversationSummary {
  readonly conversationId: string;
  readonly subject?: string;
  readonly sender?: string;
  readonly date?: string;
  readonly snippet?: string;
  /** Set when this single entry could not be fully fetched; the search as a whole still succeeds (FR-C2-5). */
  readonly error?: string;
}

/** A single email within a conversation — Outlook message `id`. (spec §4 "Message") */
export interface Message {
  readonly id: string;
  readonly conversationId: string;
  readonly from?: string;
  readonly to?: string[];
  readonly cc?: string[];
  readonly date?: string;
  readonly subject?: string;
  /** Body rendered as readable plain text, even when the source was HTML (FR-C3-3). */
  readonly bodyText: string;
  /** Organisation labels currently applied to this message (categories + folder + read-state). */
  readonly labels?: string[];
}

/** A full conversation read result, payload-bounded per FR-C3-2 / NFR-PERF-2. */
export interface ConversationDetail {
  readonly conversationId: string;
  /** Newest-first, capped to the message limit. */
  readonly messages: Message[];
  /** True when older messages and/or body characters were dropped to respect caps. */
  readonly truncated: boolean;
  /** How many older messages were omitted when `truncated` is true (FR-C3-2). */
  readonly omittedMessageCount?: number;
}

/**
 * A tag/state/location applied to mail. On Outlook this neutral concept SPLITS
 * (provider-mapping §2) into category (tag), mailFolder (location), and the
 * `isRead` flag — see {@link OrganisationLabelKind}. (spec §4 "Organisation label")
 */
export type OrganisationLabelKind = "category" | "folder" | "readState";

export interface OrganisationLabel {
  /** Stable id used by organise_mail (C8). For categories this is the name; for folders, the folder id. */
  readonly id: string;
  readonly displayName: string;
  readonly kind: OrganisationLabelKind;
  /** System-provided (Inbox, Archive, Junk, Deleted) vs user-created. (FR-C6-1) */
  readonly system: boolean;
}

/** Recipient input: a bare address or a `Display Name <addr>` form (FR-C4-5). */
export type RecipientInput = string;

/**
 * An attachment supplied as EXACTLY ONE of a local path or inline base64
 * (FR-C4-3). Path reads are guarded by the allow-list (NFR-SEC-3/4).
 */
export interface AttachmentInput {
  readonly filename?: string;
  readonly mimeType?: string;
  /** Local filesystem path — only honoured when within the configured allow-list. */
  readonly path?: string;
  /** Inline base64 content — always available, the safe alternative. Requires `filename`. */
  readonly contentBase64?: string;
}

/** Composition inputs shared by create_draft (C4) and send_message (C5). */
export interface ComposeInput {
  readonly to: RecipientInput[];
  readonly cc?: RecipientInput[];
  readonly bcc?: RecipientInput[];
  readonly subject?: string;
  readonly body: string;
  readonly isHtml?: boolean;
  readonly attachments?: AttachmentInput[];
  /** Reply target: when set, threading is derived so the message files into this conversation (FR-C4-4 / FR-C5-1). */
  readonly replyToConversationId?: string;
}

/** Opaque next-page token for paginating search results (spec §4 "Result cursor"; FR-C2-3). */
export type ResultCursor = string;

/** The optional per-tool account selector (spec §7 "Account selector"; FR-ID-1). */
export interface AccountSelectable {
  /** Account identity string; omit to use the default-selection rule (FR-ID-2). */
  readonly account?: string;
}

/**
 * Dual-channel tool output (spec §11). Every tool returns a concise human
 * summary AND a structured, authoritative payload (FR-OUT-1/2). The structured
 * object remains complete even when the text is truncated for size.
 */
export interface ToolResult<TStructured> {
  /** Concise, possibly size-truncated, human-readable summary. */
  readonly summary: string;
  /** Authoritative machine-readable payload (within documented caps). */
  readonly structured: TStructured;
}
