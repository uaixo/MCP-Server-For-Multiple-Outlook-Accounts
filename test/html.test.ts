import { describe, expect, it } from "vitest";
import { htmlToText } from "../src/util/html.js";

describe("htmlToText (FR-C3-3)", () => {
  it("strips tags and turns block elements into line breaks", () => {
    expect(htmlToText("<p>Hello <b>there</b></p><div>next</div>")).toBe("Hello there\nnext");
  });

  it("decodes named, decimal, and hex entities", () => {
    expect(htmlToText("a &amp; b")).toBe("a & b");
    expect(htmlToText("&#65;&#66;")).toBe("AB"); // decimal
    expect(htmlToText("&#x41;&#x42;")).toBe("AB"); // hex
    expect(htmlToText("it&#x27;s")).toBe("it's");
  });

  it("ignores an out-of-range numeric entity instead of throwing", () => {
    expect(htmlToText("x&#9999999999;y")).toBe("xy");
    expect(() => htmlToText("&#xFFFFFFFF;")).not.toThrow();
  });

  it("drops script/style content", () => {
    expect(htmlToText("<style>a{}</style>hi<script>evil()</script>")).toBe("hi");
  });
});
