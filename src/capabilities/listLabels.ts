/**
 * C6 — List organisation labels (FR-C6-1/2).
 *
 * On Outlook the neutral "organisation label" splits into two provider concepts
 * (provider-mapping §2), so this combines two reads:
 *   - `GET /me/outlook/masterCategories` — categories (tags), applied by name; and
 *   - `GET /me/mailFolders` — folders (locations).
 * Each is returned with a stable id, display name, kind, and a system flag so the
 * result is the id-discovery source for organize_mail (C8, FR-C6-2). Category ids
 * are their names (that is how C8 adds/removes them); folder ids are folder ids.
 *
 * Scope note: top-level folders are listed (the `GET /me/mailFolders` set);
 * nested child folders are not recursively enumerated in v1.
 *
 * Annotations (NFR-OPS-4): read-only, non-destructive, idempotent, open-world.
 */

import type { AccountRegistry, GraphClient, GraphRequest } from "../domain/contracts.js";
import type { ToolResult } from "../domain/types.js";
import type { GraphMailFolder, GraphMasterCategory } from "../graph/types.js";
import { collectPaged } from "../graph/paginate.js";
import { clampText } from "../output/contract.js";

/** Bound the number of labels returned (categories + folders). */
export const MAX_LABELS = 500;

/** Outlook's built-in folders — everything else is user-created (FR-C6-1). */
const SYSTEM_FOLDERS = new Set([
  "Inbox",
  "Drafts",
  "Sent Items",
  "Deleted Items",
  "Junk Email",
  "Junk E-mail",
  "Outbox",
  "Archive",
  "Conversation History",
  "Clutter",
  "Notes",
  "RSS Feeds",
]);

export interface LabelOut {
  readonly id: string;
  readonly display_name: string;
  readonly kind: "category" | "folder";
  readonly system: boolean;
}

export interface ListLabelsStructured {
  readonly account: string;
  readonly label_count: number;
  readonly truncated: boolean;
  readonly labels: LabelOut[];
}

export interface ListLabelsDeps {
  readonly registry: AccountRegistry;
  readonly graph: GraphClient;
}

export async function listLabels(
  deps: ListLabelsDeps,
  args: { account?: string },
): Promise<ToolResult<ListLabelsStructured>> {
  const account = await deps.registry.resolve(args.account);

  const categoriesReq: GraphRequest = {
    method: "GET",
    path: "/me/outlook/masterCategories",
    retryClass: "safe",
  };
  const foldersReq: GraphRequest = {
    method: "GET",
    path: "/me/mailFolders",
    query: { $top: 100, $select: "id,displayName,parentFolderId,childFolderCount" },
    retryClass: "safe",
  };

  const [categories, folders] = await Promise.all([
    collectPaged<GraphMasterCategory>(deps.graph, account, categoriesReq, MAX_LABELS),
    collectPaged<GraphMailFolder>(deps.graph, account, foldersReq, MAX_LABELS),
  ]);

  const labels: LabelOut[] = [
    // Categories are applied by name, so the stable id IS the display name.
    ...categories.items.map((c) => ({
      id: c.displayName,
      display_name: c.displayName,
      kind: "category" as const,
      system: false,
    })),
    ...folders.items.map((f) => ({
      id: f.id,
      display_name: f.displayName,
      kind: "folder" as const,
      system: SYSTEM_FOLDERS.has(f.displayName),
    })),
  ];

  const truncated = categories.truncated || folders.truncated;
  const structured: ListLabelsStructured = {
    account: account.displayId,
    label_count: labels.length,
    truncated,
    labels,
  };

  const fmt = (l: LabelOut) =>
    `- [${l.kind}] ${l.display_name}${l.system ? " (system)" : ""}` +
    (l.kind === "folder" ? `  (id: ${l.id})` : "");
  const header =
    `${labels.length} label(s) in ${account.displayId} ` +
    `(${categories.items.length} categor${categories.items.length === 1 ? "y" : "ies"}, ` +
    `${folders.items.length} folder(s))${truncated ? " — truncated" : ""}:`;
  const summary = clampText([header, ...labels.map(fmt)].join("\n")).text;

  return { summary, structured };
}
