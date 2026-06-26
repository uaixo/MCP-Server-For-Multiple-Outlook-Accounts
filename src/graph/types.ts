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
}

export interface GraphListResponse<T> {
  value: T[];
  "@odata.nextLink"?: string;
  "@odata.count"?: number;
}

/** Render a recipient as `Display Name <addr>`, falling back to whichever part exists. */
export function formatRecipient(r: GraphRecipient | undefined): string | undefined {
  const name = r?.emailAddress?.name?.trim();
  const addr = r?.emailAddress?.address?.trim();
  if (name && addr) return name === addr ? addr : `${name} <${addr}>`;
  return addr ?? name ?? undefined;
}
