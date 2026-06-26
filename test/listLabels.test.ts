import { describe, expect, it, vi } from "vitest";
import { listLabels } from "../src/capabilities/listLabels.js";
import { BoundedConcurrency } from "../src/util/bounded.js";
import type { Account } from "../src/domain/types.js";
import type { AccountRegistry, GraphClient, GraphRequest } from "../src/domain/contracts.js";

const account: Account = { id: "a@x.com", displayId: "a@x.com", credentialSourceId: "app1" };
const registry: AccountRegistry = { list: async () => [account], resolve: async () => account };
const limiter = new BoundedConcurrency(5);
const deps = (graph: GraphClient) => ({ registry, graph, limiter });

/** Mock that routes by request path and can page via @odata.nextLink. */
function graphByPath(routes: Record<string, unknown>) {
  const calls: GraphRequest[] = [];
  const request = vi.fn(async (_acct: Account, req: GraphRequest) => {
    calls.push(req);
    const key = Object.keys(routes).find((k) => req.path.includes(k));
    return key ? routes[key] : { value: [] };
  });
  return { graph: { request } as unknown as GraphClient, calls };
}

describe("listLabels (C6)", () => {
  it("merges categories (by name) and folders (by id), flagging system folders", async () => {
    const { graph } = graphByPath({
      masterCategories: {
        value: [
          { id: "c1", displayName: "Work", color: "preset0" },
          { id: "c2", displayName: "Personal", color: "preset1" },
        ],
      },
      mailFolders: {
        value: [
          { id: "AAA", displayName: "Inbox" },
          { id: "BBB", displayName: "Acme" },
        ],
      },
    });

    const result = await listLabels(deps(graph), {});
    const { labels } = result.structured;

    const work = labels.find((l) => l.display_name === "Work")!;
    expect(work).toMatchObject({ id: "Work", kind: "category", system: false });

    const inbox = labels.find((l) => l.display_name === "Inbox")!;
    expect(inbox).toMatchObject({ id: "AAA", kind: "folder", system: true });

    const acme = labels.find((l) => l.display_name === "Acme")!;
    expect(acme).toMatchObject({ id: "BBB", kind: "folder", system: false });

    expect(result.structured.label_count).toBe(4);
    expect(result.structured.truncated).toBe(false);
  });

  it("follows @odata.nextLink to gather all folders", async () => {
    const request = vi.fn(async (_acct: Account, req: GraphRequest) => {
      if (req.path.includes("masterCategories")) return { value: [] };
      if (req.path.includes("page2")) {
        return { value: [{ id: "F2", displayName: "Second" }] };
      }
      return {
        value: [{ id: "F1", displayName: "First" }],
        "@odata.nextLink": "https://graph.microsoft.com/v1.0/me/mailFolders?page2",
      };
    });
    const graph = { request } as unknown as GraphClient;

    const result = await listLabels(deps(graph), {});
    const names = result.structured.labels.map((l) => l.display_name);
    expect(names).toEqual(["First", "Second"]);
  });

  it("recurses into child folders and reports full paths (FR-C6-1)", async () => {
    const request = vi.fn(async (_acct: Account, req: GraphRequest) => {
      if (req.path.includes("masterCategories")) return { value: [] };
      if (req.path.includes("/childFolders")) {
        if (req.path.includes("INBOX"))
          return { value: [{ id: "SUB", displayName: "Clients", childFolderCount: 1 }] };
        if (req.path.includes("SUB"))
          return { value: [{ id: "ACME", displayName: "Acme", childFolderCount: 0 }] };
        return { value: [] };
      }
      // top-level mail folders
      return { value: [{ id: "INBOX", displayName: "Inbox", childFolderCount: 1 }] };
    });
    const graph = { request } as unknown as GraphClient;

    const result = await listLabels(deps(graph), {});
    const folders = result.structured.labels.filter((l) => l.kind === "folder");
    expect(folders.map((l) => l.display_name)).toEqual([
      "Inbox",
      "Inbox/Clients",
      "Inbox/Clients/Acme",
    ]);

    const acme = folders.find((l) => l.display_name === "Inbox/Clients/Acme")!;
    expect(acme).toMatchObject({ id: "ACME", system: false }); // nested → not system
    expect(folders.find((l) => l.display_name === "Inbox")!.system).toBe(true); // top-level well-known
  });
});
