import { describe, expect, it } from "vitest";
import { decompose, mergeCategories, resultingCategories } from "../src/organise/decompose.js";
import type { OrganiseTargetMessage } from "../src/domain/contracts.js";

const msg = (over: Partial<OrganiseTargetMessage> = {}): OrganiseTargetMessage => ({
  id: "m1",
  categories: [],
  isRead: true,
  ...over,
});

describe("mergeCategories", () => {
  it("adds, removes, and de-duplicates against current", () => {
    expect(mergeCategories(["Work"], ["Hot", "Work"], [])).toEqual(["Work", "Hot"]);
    expect(mergeCategories(["Work", "Hot"], [], ["Hot"])).toEqual(["Work"]);
  });

  it("lets remove win when a name is both added and removed", () => {
    expect(mergeCategories([], ["X"], ["X"])).toEqual([]);
  });
});

describe("decompose (FR-C8-6)", () => {
  it("collapses category add/remove and read-state into a single PATCH", () => {
    const ops = decompose(msg({ categories: ["Old"] }), {
      addLabelIds: ["Work"],
      removeLabelIds: ["Old"],
      markRead: false,
    });
    expect(ops).toHaveLength(1);
    expect(ops[0]!.request).toMatchObject({
      method: "PATCH",
      path: "/me/messages/m1",
      retryClass: "safe",
      body: { categories: ["Work"], isRead: false },
    });
  });

  it("only sets categories when no read-state change is requested", () => {
    const ops = decompose(msg({ categories: [] }), { addLabelIds: ["Work"] });
    expect(ops[0]!.request.body).toEqual({ categories: ["Work"] });
  });

  it("emits a separate move op for archive", () => {
    const ops = decompose(msg(), { markRead: true, archive: true });
    expect(ops).toHaveLength(2);
    expect(ops[0]!.request).toMatchObject({ method: "PATCH", body: { isRead: true } });
    expect(ops[1]!.request).toMatchObject({
      method: "POST",
      path: "/me/messages/m1/move",
      body: { destinationId: "archive" },
    });
  });

  it("produces no PATCH when only archive is requested", () => {
    const ops = decompose(msg(), { archive: true });
    expect(ops).toHaveLength(1);
    expect(ops[0]!.request.path).toBe("/me/messages/m1/move");
  });

  it("computes the resulting categories for the union report", () => {
    expect(resultingCategories(msg({ categories: ["A"] }), { addLabelIds: ["B"] })).toEqual([
      "A",
      "B",
    ]);
  });
});
