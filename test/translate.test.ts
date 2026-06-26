import { describe, expect, it } from "vitest";
import { SearchQueryError, translate } from "../src/search/translate.js";

describe("search translation (FR-C2-1)", () => {
  it("treats bare words as a $search string", () => {
    expect(translate("quarterly report")).toEqual({ search: "quarterly report" });
  });

  it("maps is:unread / is:read to an isRead filter", () => {
    expect(translate("is:unread")).toEqual({ filter: "isRead eq false" });
    expect(translate("is:read")).toEqual({ filter: "isRead eq true" });
  });

  it("maps from:/to: to OData address filters", () => {
    expect(translate("from:bob@x.com")).toEqual({
      filter: "from/emailAddress/address eq 'bob@x.com'",
    });
    expect(translate("to:sue@y.com")).toEqual({
      filter: "toRecipients/any(r:r/emailAddress/address eq 'sue@y.com')",
    });
  });

  it("maps has:attachment and date bounds", () => {
    expect(translate("has:attachment")).toEqual({ filter: "hasAttachments eq true" });
    expect(translate("after:2026-01-01")).toEqual({
      filter: "receivedDateTime ge 2026-01-01T00:00:00Z",
    });
    expect(translate("before:2026-02-01")).toEqual({
      filter: "receivedDateTime le 2026-02-01T00:00:00Z",
    });
  });

  it("combines multiple filter operators with 'and'", () => {
    expect(translate("from:bob@x.com is:unread")).toEqual({
      filter: "from/emailAddress/address eq 'bob@x.com' and isRead eq false",
    });
  });

  it("escapes single quotes in OData operands", () => {
    expect(translate("from:o'brien@x.com")).toEqual({
      filter: "from/emailAddress/address eq 'o''brien@x.com'",
    });
  });

  it("routes subject: and free text to $search", () => {
    expect(translate('subject:"status update"')).toEqual({ search: "subject:status update" });
    expect(translate("hello subject:foo")).toEqual({ search: "hello subject:foo" });
  });

  it("rejects mixing free-text/subject search with structured filters", () => {
    expect(() => translate("report is:unread")).toThrow(SearchQueryError);
    expect(() => translate("subject:foo from:bob@x.com")).toThrow(/one or the other/i);
  });

  it("rejects unsupported operators and values", () => {
    expect(() => translate("foo:bar")).toThrow(/Unsupported search operator/i);
    expect(() => translate("is:starred")).toThrow(/is:/);
    expect(() => translate("has:label")).toThrow(/has:/);
    expect(() => translate("after:notadate")).toThrow(/Invalid date/i);
  });
});
