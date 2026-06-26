/**
 * Minimal Microsoft Graph response shapes consumed by the read capabilities.
 * Only the fields we `$select` are modelled; everything is optional because
 * Graph omits empty values.
 */

export interface GraphEmailAddress {
  name?: string;
  address?: string;
}

export interface GraphRecipient {
  emailAddress?: GraphEmailAddress;
}

export interface GraphMessage {
  id: string;
  conversationId?: string;
  subject?: string;
  from?: GraphRecipient;
  toRecipients?: GraphRecipient[];
  ccRecipients?: GraphRecipient[];
  receivedDateTime?: string;
  bodyPreview?: string;
  body?: { contentType?: "html" | "text"; content?: string };
  categories?: string[];
  isRead?: boolean;
  /** RFC 5322 Message-ID; used to thread a reply via In-Reply-To/References. */
  internetMessageId?: string;
  /** Returned by Graph after a draft is created (a link to open it in Outlook). */
  webLink?: string;
}

/** A Graph `fileAttachment` resource (the `@odata.type` discriminator is required). */
export interface GraphFileAttachment {
  "@odata.type": "#microsoft.graph.fileAttachment";
  name: string;
  contentType: string;
  /** Base64-encoded file bytes. */
  contentBytes: string;
}

/** A custom internet (MIME) header carried on an outgoing message. */
export interface GraphInternetMessageHeader {
  name: string;
  value: string;
}

/**
 * An outgoing message resource — the body of `POST /me/messages` (draft) and the
 * `message` field of `POST /me/sendMail`. Only the fields the write path sets are
 * modelled; everything is optional because the compose layer omits empties.
 */
export interface GraphOutgoingMessage {
  subject?: string;
  body: { contentType: "Text" | "HTML"; content: string };
  toRecipients: GraphRecipient[];
  ccRecipients?: GraphRecipient[];
  bccRecipients?: GraphRecipient[];
  attachments?: GraphFileAttachment[];
  internetMessageHeaders?: GraphInternetMessageHeader[];
}

/** Build a Graph recipient resource from an address and optional display name. */
export function toGraphRecipient(address: string, name?: string): GraphRecipient {
  return { emailAddress: name ? { name, address } : { address } };
}

export interface GraphListResponse<T> {
  value: T[];
  "@odata.nextLink"?: string;
  "@odata.count"?: number;
}

/** An Outlook category (tag) from `/me/outlook/masterCategories`. Applied to messages by name. */
export interface GraphMasterCategory {
  id: string;
  displayName: string;
  /** A `categoryColor` preset ("preset0".."preset24") or "none". */
  color?: string;
}

/** An Outlook mail folder (location) from `/me/mailFolders`. */
export interface GraphMailFolder {
  id: string;
  displayName: string;
  parentFolderId?: string;
  childFolderCount?: number;
  totalItemCount?: number;
  unreadItemCount?: number;
}

/** Render a recipient as `Display Name <addr>`, falling back to whichever part exists. */
export function formatRecipient(r: GraphRecipient | undefined): string | undefined {
  const name = r?.emailAddress?.name?.trim();
  const addr = r?.emailAddress?.address?.trim();
  if (name && addr) return name === addr ? addr : `${name} <${addr}>`;
  return addr ?? name ?? undefined;
}
