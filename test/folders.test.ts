import { describe, expect, it, vi } from "vitest";
import { collectFolderTree, MAX_FOLDER_DEPTH } from "../src/graph/folders.js";
import { BoundedConcurrency } from "../src/util/bounded.js";
import type { Account } from "../src/domain/types.js";
import type { GraphClient, GraphRequest } from "../src/domain/contracts.js";

const account: Account = { id: "a@x.com", displayId: "a@x.com", credentialSourceId: "app1" };
const limiter = new BoundedConcurrency(5);

function graph(handler: (req: GraphRequest) => unknown): GraphClient {
  return { request: vi.fn(async (_a: Account, req: GraphRequest) => handler(req)) } as unknown as GraphClient;
}

describe("collectFolderTree (FR-C6-1)", () => {
  it("flattens a tree with full paths and depths", async () => {
    const g = graph((req) => {
      if (req.path.includes("/childFolders")) {
        if (req.path.includes("A")) return { value: [{ id: "A1", displayName: "Child" }] };
        return { value: [] };
      }
      return { value: [{ id: "A", displayName: "Parent", childFolderCount: 1 }] };
    });

    const { folders, truncated } = await collectFolderTree(g, account, limiter, 100);
    expect(truncated).toBe(false);
    expect(folders).toEqual([
      { id: "A", path: "Parent", name: "Parent", depth: 0 },
      { id: "A1", path: "Parent/Child", name: "Child", depth: 1 },
    ]);
  });

  it("truncates at the item cap", async () => {
    const g = graph(() => ({
      value: [
        { id: "1", displayName: "One" },
        { id: "2", displayName: "Two" },
        { id: "3", displayName: "Three" },
      ],
    }));
    const { folders, truncated } = await collectFolderTree(g, account, limiter, 2);
    expect(folders).toHaveLength(2);
    expect(truncated).toBe(true);
  });

  it("stops descending at the depth cap and flags truncation", async () => {
    // Every folder always claims to have a child → infinite depth without the cap.
    const g = graph((req) => {
      const n = (req.path.match(/F(\d+)/) ?? [])[1];
      const id = n ? `F${Number(n) + 1}` : "F1";
      return { value: [{ id, displayName: id, childFolderCount: 1 }] };
    });
    const { folders, truncated } = await collectFolderTree(g, account, limiter, 1000);
    expect(truncated).toBe(true);
    // depth 0..MAX_FOLDER_DEPTH inclusive are recorded, then descent stops.
    expect(folders).toHaveLength(MAX_FOLDER_DEPTH + 1);
  });
});
