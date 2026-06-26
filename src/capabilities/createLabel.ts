/**
 * C7 — Create a label (FR-C7-1).
 *
 * The neutral "label" splits on Outlook (provider-mapping §2/§3 C7), so the
 * caller chooses which to create via `kind`:
 *   - `category` — a tag: `POST /me/outlook/masterCategories` (with a colour
 *     preset; no nesting); or
 *   - `folder` — a location: `POST /me/mailFolders`, or
 *     `POST /me/mailFolders/{parentId}/childFolders` when `parentFolderId` is
 *     given (folder nesting, e.g. `Clients/Acme`).
 *
 * Creation is additive (FR-C7 annotation: write, non-destructive) but
 * non-idempotent, so it uses the `nonDuplicable` retry class — only a
 * pre-processing 429 is retried, never an ambiguous failure that may already
 * have created the label (NFR-REL-3).
 *
 * Annotations (NFR-OPS-4): write, non-destructive, non-idempotent, open-world.
 */

import type { AccountRegistry, GraphClient, GraphRequest } from "../domain/contracts.js";
import type { ToolResult } from "../domain/types.js";
import type { GraphMailFolder, GraphMasterCategory } from "../graph/types.js";

/** Default colour preset for a new category when none is supplied. */
const DEFAULT_CATEGORY_COLOR = "preset0";

export interface CreateLabelArgs {
  readonly account?: string;
  readonly name: string;
  readonly kind?: "category" | "folder";
  /** Category colour preset ("preset0".."preset24" or "none"); category only. */
  readonly color?: string;
  /** Parent folder id for a nested folder; folder only. */
  readonly parentFolderId?: string;
}

export interface CreateLabelStructured {
  readonly account: string;
  readonly created: true;
  readonly label: {
    readonly id: string;
    readonly display_name: string;
    readonly kind: "category" | "folder";
  };
}

export interface CreateLabelDeps {
  readonly registry: AccountRegistry;
  readonly graph: GraphClient;
}

export async function createLabel(
  deps: CreateLabelDeps,
  args: CreateLabelArgs,
): Promise<ToolResult<CreateLabelStructured>> {
  const account = await deps.registry.resolve(args.account);

  const name = args.name?.trim();
  if (!name) throw new Error("A label name is required."); // FR-ERR-3
  const kind = args.kind ?? "category";

  let req: GraphRequest;
  if (kind === "category") {
    req = {
      method: "POST",
      path: "/me/outlook/masterCategories",
      body: { displayName: name, color: args.color?.trim() || DEFAULT_CATEGORY_COLOR },
      retryClass: "nonDuplicable",
    };
  } else {
    const parent = args.parentFolderId?.trim();
    req = {
      method: "POST",
      path: parent
        ? `/me/mailFolders/${encodeURIComponent(parent)}/childFolders`
        : "/me/mailFolders",
      body: { displayName: name },
      retryClass: "nonDuplicable",
    };
  }

  const created = await deps.graph.request<GraphMasterCategory | GraphMailFolder>(account, req);

  const structured: CreateLabelStructured = {
    account: account.displayId,
    created: true,
    label: {
      // A category's stable id is its name (how C8 references it); a folder's is its id.
      id: kind === "category" ? created.displayName : created.id,
      display_name: created.displayName,
      kind,
    },
  };

  const summary = `Created ${kind} "${created.displayName}" in ${account.displayId}.`;
  return { summary, structured };
}
