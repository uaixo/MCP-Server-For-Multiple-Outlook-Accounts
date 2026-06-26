import { describe, expect, it } from "vitest";
import { composeMessage, parseRecipient } from "../src/mail/compose.js";
import { MAX_INLINE_ATTACHMENT_BYTES } from "../src/output/contract.js";
import type { ComposeInput, ResolvedAttachment } from "../src/domain/contracts.js";

const bytes = (s: string): ResolvedAttachment => ({
  filename: "f.bin",
  mimeType: "application/octet-stream",
  bytes: new Uint8Array(Buffer.from(s, "utf8")),
});

const base: ComposeInput = { to: ["a@x.com"], body: "hi" };

describe("parseRecipient (FR-C4-5)", () => {
  it("parses a bare address", () => {
    expect(parseRecipient("a@x.com")).toEqual({ emailAddress: { address: "a@x.com" } });
  });

  it("parses a Display Name <addr> form", () => {
    expect(parseRecipient("Alice Example <alice@x.com>")).toEqual({
      emailAddress: { name: "Alice Example", address: "alice@x.com" },
    });
  });

  it("strips header-injection attempts from name and address (NFR-SEC-5)", () => {
    const r = parseRecipient("Ann\r\nBcc: evil@x.com <ann@x.com>");
    expect(r.emailAddress?.name).toBe("Ann Bcc: evil@x.com");
    expect(r.emailAddress?.address).toBe("ann@x.com");
  });

  it("rejects an invalid address", () => {
    expect(() => parseRecipient("not-an-email")).toThrow(/invalid recipient/i);
    expect(() => parseRecipient("  ")).toThrow(/cannot be empty/i);
  });
});

describe("composeMessage (FR-C4-1)", () => {
  it("builds to/cc/bcc, a text body, and requires at least one recipient", () => {
    const out = composeMessage(
      { to: ["a@x.com"], cc: ["c@x.com"], bcc: ["b@x.com"], subject: "Hi", body: "yo" },
      [],
    );
    expect(out.message.toRecipients).toEqual([{ emailAddress: { address: "a@x.com" } }]);
    expect(out.message.ccRecipients).toEqual([{ emailAddress: { address: "c@x.com" } }]);
    expect(out.message.bccRecipients).toEqual([{ emailAddress: { address: "b@x.com" } }]);
    expect(out.message.subject).toBe("Hi");
    expect(out.message.body).toEqual({ contentType: "Text", content: "yo" });
    expect(out.recipients.to).toEqual(["a@x.com"]);
  });

  it("throws when there are no recipients", () => {
    expect(() => composeMessage({ to: [], body: "x" }, [])).toThrow(/at least one recipient/i);
  });

  it("marks the body HTML when is_html is set", () => {
    const out = composeMessage({ ...base, isHtml: true, body: "<p>hi</p>" }, []);
    expect(out.message.body.contentType).toBe("HTML");
  });

  it("encodes attachments as fileAttachment resources (FR-C4-3)", () => {
    const out = composeMessage(base, [
      { filename: "a.txt", mimeType: "text/plain", bytes: new Uint8Array(Buffer.from("hello")) },
    ]);
    expect(out.message.attachments).toEqual([
      {
        "@odata.type": "#microsoft.graph.fileAttachment",
        name: "a.txt",
        contentType: "text/plain",
        contentBytes: Buffer.from("hello").toString("base64"),
      },
    ]);
  });
});

describe("composeMessage — reply threading (FR-C4-4)", () => {
  it("defaults an omitted subject to the original prefixed with Re:", () => {
    const out = composeMessage(base, [], { reply: { subject: "Project plan" } });
    expect(out.message.subject).toBe("Re: Project plan");
  });

  it("does not double-prefix an already-Re: subject", () => {
    const out = composeMessage(base, [], { reply: { subject: "RE: Project plan" } });
    expect(out.message.subject).toBe("RE: Project plan");
  });

  it("lets an explicit subject win over the reply default", () => {
    const out = composeMessage({ ...base, subject: "Custom" }, [], {
      reply: { subject: "Original" },
    });
    expect(out.message.subject).toBe("Custom");
  });

  it("sets In-Reply-To/References from the original internetMessageId", () => {
    const out = composeMessage(base, [], {
      reply: { internetMessageId: "<abc@mail>" },
    });
    expect(out.message.internetMessageHeaders).toEqual([
      { name: "In-Reply-To", value: "<abc@mail>" },
      { name: "References", value: "<abc@mail>" },
    ]);
  });
});

describe("composeMessage — outgoing size guard (NFR-PERF-3)", () => {
  it("rejects a message over the size limit before any send", () => {
    expect(() =>
      composeMessage({ ...base, body: "x".repeat(50) }, [bytes("y".repeat(50))], {
        maxBytes: 10,
      }),
    ).toThrow(/over the .* MB limit/i);
  });

  it("reports the computed raw size when within the limit", () => {
    const out = composeMessage({ ...base, body: "12345" }, [bytes("678")]);
    expect(out.sizeBytes).toBe(8);
  });
});

describe("composeMessage — inline vs upload split", () => {
  it("inlines small attachments and defers large ones to uploadAttachments", () => {
    const big: ResolvedAttachment = {
      filename: "big.bin",
      mimeType: "application/octet-stream",
      bytes: new Uint8Array(MAX_INLINE_ATTACHMENT_BYTES + 1),
    };
    const small: ResolvedAttachment = {
      filename: "s.txt",
      mimeType: "text/plain",
      bytes: new Uint8Array(Buffer.from("hi")),
    };

    const out = composeMessage(base, [big, small]);

    // Small one is inlined; the large one is held back for the upload session.
    expect(out.message.attachments).toHaveLength(1);
    expect(out.message.attachments![0]!.name).toBe("s.txt");
    expect(out.uploadAttachments.map((a) => a.filename)).toEqual(["big.bin"]);
    // The size guard counts the body ("hi") + every attachment, inline or not.
    expect(out.sizeBytes).toBe(2 + 2 + (MAX_INLINE_ATTACHMENT_BYTES + 1));
  });
});
