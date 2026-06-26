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
 *
 * Size: each attachment is bounded by `MAX_INLINE_ATTACHMENT_BYTES` (~3 MB),
 * the limit for a `fileAttachment` sent inline in one Graph request. Larger
 * files need an upload session, which v1 does not implement.
 */

import { open, realpath, type FileHandle } from "node:fs/promises";
import { constants as FS } from "node:fs";
import { basename, extname, resolve, sep } from "node:path";
import type { AttachmentInput, AttachmentReader, ResolvedAttachment } from "../domain/contracts.js";
import { MAX_INLINE_ATTACHMENT_BYTES } from "../output/contract.js";
import { sanitizeFilename } from "./sanitize.js";

/**
 * Open read-only and (where the OS supports it) refuse to follow a symlink at
 * the FINAL path segment. We open the already-`realpath`'d path, so its last
 * component is a real file under normal conditions; `O_NOFOLLOW` closes the
 * narrow TOCTOU window where an attacker swaps that component for a symlink
 * between resolution and open (NFR-SEC-4). `O_NOFOLLOW` is absent on Windows
 * (where creating symlinks needs privilege anyway), so it degrades to `O_RDONLY`.
 */
const OPEN_READ_NOFOLLOW = FS.O_RDONLY | (FS.O_NOFOLLOW ?? 0);

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

function tooLarge(filename: string, size: number, maxBytes: number): Error {
  const mb = (n: number) => (n / (1024 * 1024)).toFixed(1);
  return new Error(
    `Attachment "${filename}" is ${mb(size)} MB, over the ${mb(maxBytes)} MB limit.`,
  );
}

/** Reads attachments from the filesystem, guarded by an allow-list (NFR-SEC-3/4). */
export class FsAttachmentReader implements AttachmentReader {
  private readonly allowList: readonly string[];
  private readonly maxBytes: number;

  constructor(allowList: readonly string[], maxBytes: number = MAX_INLINE_ATTACHMENT_BYTES) {
    this.allowList = allowList;
    this.maxBytes = maxBytes;
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
    if (bytes.byteLength > this.maxBytes) {
      throw tooLarge(filename, bytes.byteLength, this.maxBytes); // NFR-PERF-3
    }
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
    // check-then-reopen window (TOCTOU-safe, NFR-SEC-4). O_NOFOLLOW rejects a
    // final-segment symlink swapped in after resolution.
    let handle: FileHandle;
    try {
      handle = await open(realPath, OPEN_READ_NOFOLLOW);
    } catch (e) {
      if (e instanceof Error && "code" in e && (e as { code?: string }).code === "ELOOP") {
        throw new Error("Refusing to follow a symlink to the attachment file.");
      }
      throw e;
    }
    try {
      const stat = await handle.stat();
      if (!stat.isFile()) {
        throw new Error("Attachment path does not point to a regular file.");
      }
      // Reject oversize files BEFORE reading them into memory, so a giant file
      // inside the allow-list can't exhaust memory (NFR-PERF-3).
      if (stat.size > this.maxBytes) {
        throw tooLarge(basename(realPath), stat.size, this.maxBytes);
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
