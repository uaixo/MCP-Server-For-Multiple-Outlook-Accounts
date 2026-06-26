/**
 * Recursive mail-folder enumeration (FR-C6-1).
 *
 * `GET /me/mailFolders` returns only top-level folders; nested folders live under
 * `GET /me/mailFolders/{id}/childFolders`. This walks the tree level by level so
 * nested folders are discoverable (e.g. a `Clients/Acme` folder created under a
 * parent), building each folder's full path. It is bounded two ways: a total
 * `maxItems` cap and {@link MAX_FOLDER_DEPTH}, and child fetches at each level run
 * under the concurrency limiter (NFR-REL-4). `truncated` is set when either bound
 * elides folders.
 */

import type { Account, ConcurrencyLimiter, GraphClient, GraphRequest } from "../domain/contracts.js";
import type { GraphMailFolder } from "./types.js";
import { collectPaged } from "./paginate.js";

/** Max folder-tree depth traversed; bounds recursion on pathological nesting. */
export const MAX_FOLDER_DEPTH = 10;

const FOLDER_SELECT = "id,displayName,parentFolderId,childFolderCount";

export interface FolderNode {
  readonly id: string;
  /** Full path from the root, e.g. "Inbox/Clients/Acme". */
  readonly path: string;
  /** The folder's own (leaf) display name. */
  readonly name: string;
  /** 0 for top-level folders. */
  readonly depth: number;
}

function folderListRequest(path: string): GraphRequest {
  return { method: "GET", path, query: { $top: 100, $select: FOLDER_SELECT }, retryClass: "safe" };
}

interface Pending {
  readonly folder: GraphMailFolder;
  readonly path: string;
  readonly depth: number;
}

export async function collectFolderTree(
  graph: GraphClient,
  account: Account,
  limiter: ConcurrencyLimiter,
  maxItems: number,
): Promise<{ folders: FolderNode[]; truncated: boolean }> {
  const top = await collectPaged<GraphMailFolder>(
    graph,
    account,
    folderListRequest("/me/mailFolders"),
    maxItems,
  );
  let truncated = top.truncated;

  const folders: FolderNode[] = [];
  let frontier: Pending[] = top.items.map((f) => ({ folder: f, path: f.displayName, depth: 0 }));

  while (frontier.length > 0) {
    for (const node of frontier) {
      if (folders.length >= maxItems) break;
      folders.push({
        id: node.folder.id,
        path: node.path,
        name: node.folder.displayName,
        depth: node.depth,
      });
    }
    if (folders.length >= maxItems) {
      truncated = truncated || frontier.length > 0;
      break;
    }

    // Descend into folders that have children and are within the depth bound.
    const hasChildren = (n: Pending) => (n.folder.childFolderCount ?? 0) > 0;
    const parents = frontier.filter((n) => hasChildren(n) && n.depth < MAX_FOLDER_DEPTH);
    // Children we won't descend into because of the depth cap are elided.
    if (frontier.some((n) => hasChildren(n) && n.depth >= MAX_FOLDER_DEPTH)) truncated = true;
    if (parents.length === 0) break;

    const tasks = parents.map((parent) => async () => {
      const page = await collectPaged<GraphMailFolder>(
        graph,
        account,
        folderListRequest(`/me/mailFolders/${encodeURIComponent(parent.folder.id)}/childFolders`),
        maxItems,
      );
      return {
        truncated: page.truncated,
        children: page.items.map((cf) => ({
          folder: cf,
          path: `${parent.path}/${cf.displayName}`,
          depth: parent.depth + 1,
        })),
      };
    });
    const results = await limiter.run(tasks);
    if (results.some((r) => r.truncated)) truncated = true;
    frontier = results.flatMap((r) => r.children);
  }

  return { folders, truncated };
}
