import { describe, expect, it } from "vitest";
import {
  clampPageSize,
  clampText,
  DEFAULT_PAGE_SIZE,
  MAX_PAGE_SIZE,
} from "../src/output/contract.js";
import { htmlToText } from "../src/util/html.js";

describe("output bounds (NFR-PERF-1, FR-C2-2)", () => {
  it("clampText truncates with an ellipsis past the budget", () => {
    const r = clampText("abcdef", 4);
    expect(r.truncated).toBe(true);
    expect(r.text.endsWith("…")).toBe(true);
    expect(r.text.length).toBe(4);
  });

  it("clampText keeps text within budget", () => {
    const r = clampText("abc", 10);
    expect(r).toEqual({ text: "abc", truncated: false });
  });

  it("clampPageSize defaults, floors, and caps", () => {
    expect(clampPageSize(undefined)).toBe(DEFAULT_PAGE_SIZE);
    expect(clampPageSize(0)).toBe(1);
    expect(clampPageSize(5)).toBe(5);
    expect(clampPageSize(1000)).toBe(MAX_PAGE_SIZE);
  });
});

describe("htmlToText (FR-C3-3)", () => {
  it("strips tags, converts breaks, and decodes entities", () => {
    const text = htmlToText("<p>Hello&nbsp;<b>world</b></p><br>Line&amp;two<script>x()</script>");
    expect(text).toContain("Hello world");
    expect(text).toContain("Line&two");
    expect(text).not.toContain("<");
    expect(text).not.toContain("x()");
  });
});
