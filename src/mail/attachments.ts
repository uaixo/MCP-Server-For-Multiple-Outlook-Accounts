/**
 * Safe attachment resolution for the write path (NFR-SEC-3/4, FR-C4-3).
 *
 * Each attachment is supplied as **exactly one of**:
 *  - inline base64 `contentBase64` (always available — the safe default), or
 *  - a local file `path` (only honoured when it resolves inside the configured
 *    allow-list `OUTLOOK_MCP_ATTACHMENTS_DIR`).
 *
 * Path reads are the dangerous case: without a guard the server could be coaxed
 * into emailing arbitrary local files (keys, `.env`). So we:
 *  1. refuse path reads entirely when the allow-list is empty (NFR-SEC-3);
 *  2. fully resolve the path (symlinks, `..`) with `realpath` and confirm it
 *     sits inside an allow-listed directory **before** opening (NFR-SEC-3);
 *  3. open the resolved file **once** and read through that handle, validating
 *     it is a regular file via the handle — no check-then-reopen window
 *     (TOCTOU-safe, NFR-SEC-4).
 *
 * Filenames/MIME types are inferred when omitted; inline content must name its
 * file. Filenames are sanitized so they cannot inject mail headers (NFR-SEC-5).
 */

import { open, realpath } from "node:fs/promises";
import { basename, extname, resolve, sep } from "node:path";
import type { AttachmentInput, AttachmentReader, ResolvedAttachment } from "../domain/contracts.js";
import { sanitizeFilename } from "./sanitize.js";

/** Minimal extension → MIME map; everything else is the generic binary type. */
const MIME_BY_EXT: Record<string, string> = {
  ".txt": "text/plain",
  ".csv": "text/csv",
  ".html": "text/html",
  ".htm": "text/html",
  ".md": "text/markdown",
  ".json": "application/json",
  ".pdf": "application/pdf",
  ".png": "image/png",
  ".jpg": "image/jpeg",
  ".jpeg": "image/jpeg",
  ".gif": "image/gif",
  ".svg": "image/svg+xml",
  ".zip": "application/zip",
  ".doc": "application/msword",
  ".docx": "application/vnd.openxmlformats-officedocument.wordprocessingml.document",
  ".xls": "application/vnd.ms-excel",
  ".xlsx": "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
  ".ppt": "application/vnd.ms-powerpoint",
  ".pptx": "application/vnd.openxmlformats-officedocument.presentationml.presentation",
};
const DEFAULT_MIME = "application/octet-stream";

function inferMime(filename: string, provided?: string): string {
  const fromInput = provided?.trim();
  if (fromInput) return fromInput;
  return MIME_BY_EXT[extname(filename).toLowerCase()] ?? DEFAULT_MIME;
}

/** Reads attachments from the filesystem, guarded by an allow-list (NFR-SEC-3/4). */
export class FsAttachmentReader implements AttachmentReader {
  private readonly allowList: readonly string[];

  constructor(allowList: readonly string[]) {
    this.allowList = allowList;
  }

  async read(input: AttachmentInput): Promise<ResolvedAttachment> {
    const hasPath = typeof input.path === "string" && input.path.trim() !== "";
    const hasInline = typeof input.contentBase64 === "string" && input.contentBase64.trim() !== "";

    // Exactly one source must be supplied (FR-C4-3).
    if (hasPath && hasInline) {
      throw new Error("An attachment must supply either a path or inline content, not both.");
    }
    if (!hasPath && !hasInline) {
      throw new Error("An attachment must supply either a local path or inline base64 content.");
    }

    return hasInline ? this.readInline(input) : this.readPath(input);
  }

  private readInline(input: AttachmentInput): ResolvedAttachment {
    const name = input.filename?.trim();
    if (!name) {
      throw new Error("Inline (base64) attachments require a filename."); // FR-C4-3
    }
    const bytes = decodeBase64(input.contentBase64!);
    const filename = sanitizeFilename(name);
    return { filename, mimeType: inferMime(filename, input.mimeType), bytes };
  }

  private async readPath(input: AttachmentInput): Promise<ResolvedAttachment> {
    if (this.allowList.length === 0) {
      // Path reads disabled by default; inline base64 is the safe alternative (NFR-SEC-3).
      throw new Error(
        "Reading attachments by path is disabled. Set OUTLOOK_MCP_ATTACHMENTS_DIR to an " +
          "allowed directory, or supply the attachment as inline base64 content.",
      );
    }

    // Resolve symlinks and `..` up front, then confirm the real path is inside
    // an allow-listed directory before opening anything (NFR-SEC-3).
    const realPath = await realpath(resolve(input.path!));
    if (!(await this.isInsideAllowList(realPath))) {
      throw new Error(
        "Refusing to read an attachment outside the allowed directory " +
          "(OUTLOOK_MCP_ATTACHMENTS_DIR).",
      );
    }

    // Open the resolved path ONCE and read through the handle so there is no
    // check-then-reopen window (TOCTOU-safe, NFR-SEC-4).
    const handle = await open(realPath, "r");
    try {
      const stat = await handle.stat();
      if (!stat.isFile()) {
        throw new Error("Attachment path does not point to a regular file.");
      }
      const buffer = await handle.readFile();
      const filename = sanitizeFilename(input.filename?.trim() || basename(realPath));
      return {
        filename,
        mimeType: inferMime(filename, input.mimeType),
        bytes: new Uint8Array(buffer.buffer, buffer.byteOffset, buffer.byteLength),
      };
    } finally {
      await handle.close();
    }
  }

  private async isInsideAllowList(realPath: string): Promise<boolean> {
    for (const dir of this.allowList) {
      let realDir: string;
      try {
        realDir = await realpath(resolve(dir));
      } catch {
        continue; // a configured dir that doesn't exist can't contain anything
      }
      if (realPath === realDir || realPath.startsWith(realDir + sep)) return true;
    }
    return false;
  }
}

function decodeBase64(value: string): Uint8Array {
  const buffer = Buffer.from(value, "base64");
  // Buffer.from silently drops invalid characters; an all-invalid string yields
  // an empty buffer. Treat a non-empty input that decodes to nothing as invalid.
  if (buffer.length === 0 && value.trim() !== "") {
    throw new Error("Inline attachment content is not valid base64.");
  }
  return new Uint8Array(buffer.buffer, buffer.byteOffset, buffer.byteLength);
}
