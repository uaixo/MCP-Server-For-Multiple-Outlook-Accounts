import { afterAll, beforeAll, describe, expect, it } from "vitest";
import { mkdtemp, rm, writeFile, symlink, mkdir } from "node:fs/promises";
import { tmpdir } from "node:os";
import { join } from "node:path";
import { FsAttachmentReader } from "../src/mail/attachments.js";
import { MAX_INLINE_ATTACHMENT_BYTES } from "../src/output/contract.js";

let root: string;
let allowedDir: string;
let outsideDir: string;

beforeAll(async () => {
  root = await mkdtemp(join(tmpdir(), "att-test-"));
  allowedDir = join(root, "allowed");
  outsideDir = join(root, "outside");
  await mkdir(allowedDir);
  await mkdir(outsideDir);
  await writeFile(join(allowedDir, "report.pdf"), "PDFDATA");
  await writeFile(join(outsideDir, "secret.env"), "off-limits file contents");
});

afterAll(async () => {
  await rm(root, { recursive: true, force: true });
});

const toText = (bytes: Uint8Array) => Buffer.from(bytes).toString("utf8");
const b64 = (s: string) => Buffer.from(s, "utf8").toString("base64");

describe("FsAttachmentReader — input validation (FR-C4-3)", () => {
  const reader = new FsAttachmentReader([]);

  it("rejects an attachment with neither path nor inline content", async () => {
    await expect(reader.read({})).rejects.toThrow(/either a local path or inline/i);
  });

  it("rejects an attachment with both path and inline content", async () => {
    await expect(
      reader.read({ path: "/x", contentBase64: b64("hi"), filename: "f" }),
    ).rejects.toThrow(/not both/i);
  });

  it("requires a filename for inline content", async () => {
    await expect(reader.read({ contentBase64: b64("hi") })).rejects.toThrow(/require a filename/i);
  });

  it("rejects inline content that is not valid base64", async () => {
    await expect(reader.read({ contentBase64: "@@@@", filename: "x.bin" })).rejects.toThrow(
      /not valid base64/i,
    );
  });
});

describe("FsAttachmentReader — inline base64 (the safe path)", () => {
  const reader = new FsAttachmentReader([]); // allow-list empty: inline still works

  it("decodes bytes and infers MIME from the filename", async () => {
    const out = await reader.read({ contentBase64: b64("hello"), filename: "note.txt" });
    expect(toText(out.bytes)).toBe("hello");
    expect(out.filename).toBe("note.txt");
    expect(out.mimeType).toBe("text/plain");
  });

  it("honours an explicit MIME type and sanitizes the filename", async () => {
    const out = await reader.read({
      contentBase64: b64("x"),
      filename: "a\r\nb.bin",
      mimeType: "application/custom",
    });
    expect(out.filename).toBe("a b.bin");
    expect(out.mimeType).toBe("application/custom");
  });
});

describe("FsAttachmentReader — path guard (NFR-SEC-3/4)", () => {
  it("refuses path reads when the allow-list is empty", async () => {
    const reader = new FsAttachmentReader([]);
    await expect(reader.read({ path: join(allowedDir, "report.pdf") })).rejects.toThrow(
      /disabled/i,
    );
  });

  it("reads a file inside an allow-listed directory and infers filename + MIME", async () => {
    const reader = new FsAttachmentReader([allowedDir]);
    const out = await reader.read({ path: join(allowedDir, "report.pdf") });
    expect(toText(out.bytes)).toBe("PDFDATA");
    expect(out.filename).toBe("report.pdf");
    expect(out.mimeType).toBe("application/pdf");
  });

  it("refuses a path outside the allow-list", async () => {
    const reader = new FsAttachmentReader([allowedDir]);
    await expect(reader.read({ path: join(outsideDir, "secret.env") })).rejects.toThrow(
      /outside the allowed directory/i,
    );
  });

  it("refuses a `..` traversal that escapes the allow-list", async () => {
    const reader = new FsAttachmentReader([allowedDir]);
    const escaping = join(allowedDir, "..", "outside", "secret.env");
    await expect(reader.read({ path: escaping })).rejects.toThrow(/outside the allowed directory/i);
  });

  it("refuses a symlink inside the allow-list that points outside it", async () => {
    const link = join(allowedDir, "link-to-secret.env");
    await symlink(join(outsideDir, "secret.env"), link);
    const reader = new FsAttachmentReader([allowedDir]);
    await expect(reader.read({ path: link })).rejects.toThrow(/outside the allowed directory/i);
  });

  it("still reads through a symlink that resolves to a real file inside the allow-list", async () => {
    // realpath resolves the link to the real target before O_NOFOLLOW opens it,
    // so legitimate in-allow-list symlinks keep working.
    const link = join(allowedDir, "link-to-report.pdf");
    await symlink(join(allowedDir, "report.pdf"), link);
    const reader = new FsAttachmentReader([allowedDir]);
    const out = await reader.read({ path: link });
    expect(toText(out.bytes)).toBe("PDFDATA");
  });
});

describe("FsAttachmentReader — size guard (NFR-PERF-3)", () => {
  it("rejects an oversize file by stat.size before reading it", async () => {
    const reader = new FsAttachmentReader([allowedDir], 4); // "PDFDATA" is 7 bytes
    await expect(reader.read({ path: join(allowedDir, "report.pdf") })).rejects.toThrow(
      /over the .* MB limit/i,
    );
  });

  it("rejects oversize inline base64 content", async () => {
    const reader = new FsAttachmentReader([], 4);
    await expect(reader.read({ contentBase64: b64("hello"), filename: "n.txt" })).rejects.toThrow(
      /over the .* MB limit/i,
    );
  });

  it("accepts content within the limit", async () => {
    const reader = new FsAttachmentReader([allowedDir], 1024);
    const out = await reader.read({ path: join(allowedDir, "report.pdf") });
    expect(toText(out.bytes)).toBe("PDFDATA");
  });

  it("defaults to the inline (~3 MB) attachment limit", async () => {
    const reader = new FsAttachmentReader([]); // default maxBytes
    const oversize = Buffer.alloc(MAX_INLINE_ATTACHMENT_BYTES + 1).toString("base64");
    await expect(reader.read({ contentBase64: oversize, filename: "big.bin" })).rejects.toThrow(
      /over the .* MB limit/i,
    );
  });
});
