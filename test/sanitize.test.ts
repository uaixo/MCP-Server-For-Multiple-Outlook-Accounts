import { describe, expect, it } from "vitest";
import { sanitizeFilename, sanitizeHeaderValue } from "../src/mail/sanitize.js";

const NUL = String.fromCharCode(0x00);
const BEL = String.fromCharCode(0x07);
const LINE_SEP = String.fromCharCode(0x2028);
const PARA_SEP = String.fromCharCode(0x2029);

describe("sanitizeHeaderValue (NFR-SEC-5)", () => {
  it("removes CR and LF so headers cannot be split", () => {
    const injected = "Real Name\r\nBcc: victim@evil.com";
    const cleaned = sanitizeHeaderValue(injected);
    expect(cleaned).not.toMatch(/[\r\n]/);
    expect(cleaned).toBe("Real Name Bcc: victim@evil.com");
  });

  it("collapses a run of line breaks into a single space", () => {
    expect(sanitizeHeaderValue("a\r\n\n\rb")).toBe("a b");
  });

  it("strips other control characters but keeps tabs and ordinary text", () => {
    expect(sanitizeHeaderValue(`Quarterly${NUL} report${BEL}`)).toBe("Quarterly report");
    expect(sanitizeHeaderValue("col1\tcol2")).toBe("col1\tcol2");
  });

  it("strips Unicode line/paragraph separators", () => {
    expect(sanitizeHeaderValue(`a${LINE_SEP}b${PARA_SEP}c`)).toBe("a b c");
  });

  it("trims surrounding whitespace introduced by leading/trailing breaks", () => {
    expect(sanitizeHeaderValue("\r\nhello\r\n")).toBe("hello");
  });
});

describe("sanitizeFilename (NFR-SEC-5)", () => {
  it("strips header-injecting characters", () => {
    expect(sanitizeFilename("report\r\n.pdf")).toBe("report .pdf");
  });

  it("replaces path separators so a filename cannot traverse", () => {
    expect(sanitizeFilename("../../etc/passwd")).toBe(".._.._etc_passwd");
    expect(sanitizeFilename("a\\b\\c.txt")).toBe("a_b_c.txt");
  });
});
