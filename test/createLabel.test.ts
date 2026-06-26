import { describe, expect, it, vi } from "vitest";
import { createLabel } from "../src/capabilities/createLabel.js";
import type { Account } from "../src/domain/types.js";
import type { AccountRegistry, GraphClient, GraphRequest } from "../src/domain/contracts.js";

const account: Account = { id: "a@x.com", displayId: "a@x.com", credentialSourceId: "app1" };
const registry: AccountRegistry = { list: async () => [account], resolve: async () => account };

function graphReturning(response: unknown) {
  const calls: GraphRequest[] = [];
  const request = vi.fn(async (_acct: Account, req: GraphRequest) => {
    calls.push(req);
    return response;
  });
  return { graph: { request } as unknown as GraphClient, calls };
}

describe("createLabel (C7)", () => {
  it("creates a category by name with a default colour (id = name)", async () => {
    const { graph, calls } = graphReturning({ id: "c1", displayName: "Clients", color: "preset0" });

    const result = await createLabel({ registry, graph }, { name: "Clients" });

    const req = calls[0]!;
    expect(req.method).toBe("POST");
    expect(req.path).toBe("/me/outlook/masterCategories");
    expect(req.retryClass).toBe("nonDuplicable");
    expect(req.body).toEqual({ displayName: "Clients", color: "preset0" });

    expect(result.structured.label).toEqual({
      id: "Clients",
      display_name: "Clients",
      kind: "category",
    });
  });

  it("honours an explicit category colour", async () => {
    const { graph, calls } = graphReturning({ id: "c2", displayName: "Hot", color: "preset5" });
    await createLabel({ registry, graph }, { name: "Hot", kind: "category", color: "preset5" });
    expect((calls[0]!.body as { color: string }).color).toBe("preset5");
  });

  it("creates a top-level folder (id = folder id)", async () => {
    const { graph, calls } = graphReturning({ id: "F1", displayName: "Acme" });
    const result = await createLabel({ registry, graph }, { name: "Acme", kind: "folder" });

    expect(calls[0]!.path).toBe("/me/mailFolders");
    expect(calls[0]!.body).toEqual({ displayName: "Acme" });
    expect(result.structured.label).toEqual({ id: "F1", display_name: "Acme", kind: "folder" });
  });

  it("creates a nested folder under a parent (childFolders)", async () => {
    const { graph, calls } = graphReturning({ id: "F2", displayName: "Acme" });
    await createLabel(
      { registry, graph },
      { name: "Acme", kind: "folder", parentFolderId: "PARENT" },
    );
    expect(calls[0]!.path).toBe("/me/mailFolders/PARENT/childFolders");
  });

  it("requires a non-empty name", async () => {
    const { graph } = graphReturning({});
    await expect(createLabel({ registry, graph }, { name: "  " })).rejects.toThrow(
      /name is required/i,
    );
  });
});
