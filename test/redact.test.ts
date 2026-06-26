import { describe, expect, it } from "vitest";
import { redact, redactError } from "../src/util/redact.js";

// A JWT-shaped sample (not a real token) used to prove access tokens are scrubbed.
const JWT_SAMPLE = "eyJhbGciOi.eyJzdWIiOiJ123.sIgNaTuRe-x";

describe("redact (NFR-SEC-6)", () => {
  it("scrubs an Authorization Bearer token", () => {
    const out = redact(`Authorization: Bearer ${JWT_SAMPLE}`);
    expect(out).not.toContain(JWT_SAMPLE);
    expect(out).toContain("[redacted]");
  });

  it("scrubs a bare Bearer token (no Authorization key)", () => {
    const out = redact(`sent header Bearer ${JWT_SAMPLE} upstream`);
    expect(out).not.toContain(JWT_SAMPLE);
    expect(out).toContain("Bearer [redacted]");
  });

  it("scrubs a standalone JWT (access/id token)", () => {
    const out = redact(`unexpected token ${JWT_SAMPLE} in cache`);
    expect(out).not.toContain(JWT_SAMPLE);
    expect(out).toContain("[redacted]");
  });

  it("scrubs an access_token in a query string", () => {
    const out = redact("GET /token?access_token=abc123def456&scope=mail");
    expect(out).not.toContain("abc123def456");
    expect(out).toContain("scope=mail"); // non-secret params survive
  });

  it("scrubs refresh_token and client_secret in JSON", () => {
    const out = redact('{"refresh_token":"0.ARoAr3","client_secret":"sup3r-s3cret"}');
    expect(out).not.toContain("0.ARoAr3");
    expect(out).not.toContain("sup3r-s3cret");
  });

  it("scrubs a password assignment", () => {
    expect(redact("password: hunter2")).not.toContain("hunter2");
  });

  it("leaves ordinary diagnostic text (including HTTP status codes) intact", () => {
    const msg = "Microsoft Graph had a transient error (HTTP 503). Try again shortly.";
    expect(redact(msg)).toBe(msg);
  });
});

describe("redactError (NFR-SEC-6)", () => {
  it("redacts an Error message", () => {
    const err = new Error(`MSAL failed with Bearer ${JWT_SAMPLE}`);
    expect(redactError(err)).not.toContain(JWT_SAMPLE);
  });

  it("stringifies and redacts a non-Error value", () => {
    expect(redactError(`refresh_token=${JWT_SAMPLE}`)).not.toContain(JWT_SAMPLE);
  });
});
