import { mkdtemp, rm, stat, writeFile } from "node:fs/promises";
import { tmpdir } from "node:os";
import { join } from "node:path";
import { afterEach, beforeEach, describe, expect, it, vi } from "vitest";
import { FileTokenStore } from "../src/auth/tokenStore.js";

describe("FileTokenStore (NFR-SEC-1/2, FR-ERR-2)", () => {
  let dir: string;
  const newStore = (warn?: (m: string) => void) =>
    new FileTokenStore({ dataDir: dir, lockTimeoutMs: 2000, warn });

  beforeEach(async () => {
    dir = await mkdtemp(join(tmpdir(), "store-"));
  });
  afterEach(async () => {
    await rm(dir, { recursive: true, force: true });
  });

  it("round-trips an account and its cache, keyed case-insensitively (FR-ID-4)", async () => {
    const store = newStore();
    await store.upsert(
      { id: "User@Example.com", displayId: "User@Example.com", credentialSourceId: "app1" },
      '{"cache":1}',
    );

    const accounts = await store.list();
    expect(accounts).toHaveLength(1);
    expect(accounts[0]).toMatchObject({
      id: "user@example.com", // lower-cased key
      displayId: "User@Example.com", // original case preserved
      credentialSourceId: "app1",
    });
    expect(await store.readCache("USER@EXAMPLE.COM")).toBe('{"cache":1}');
  });

  // POSIX file modes only; Windows enforces owner-only access via ACLs, where
  // chmod is a no-op (NFR-SEC-1). Skip the bit-mode assertion there.
  it.skipIf(process.platform === "win32")(
    "writes the store file 0600 inside a 0700 data dir (NFR-SEC-1)",
    async () => {
      const store = newStore();
      await store.upsert({ id: "a@x.com", displayId: "a@x.com", credentialSourceId: "app1" }, "{}");

      const dirMode = (await stat(dir)).mode & 0o777;
      const fileMode = (await stat(join(dir, "tokens.json"))).mode & 0o777;
      expect(dirMode).toBe(0o700);
      expect(fileMode).toBe(0o600);
    },
  );

  it("removes an account", async () => {
    const store = newStore();
    await store.upsert({ id: "a@x.com", displayId: "a@x.com", credentialSourceId: "app1" }, "{}");
    await store.remove("A@X.com");
    expect(await store.list()).toEqual([]);
  });

  it("treats a corrupt store as empty and warns once (FR-ERR-2)", async () => {
    await writeFile(join(dir, "tokens.json"), "{ this is not valid json");
    const warn = vi.fn();
    const store = newStore(warn);

    expect(await store.list()).toEqual([]);
    expect(await store.list()).toEqual([]);
    expect(warn).toHaveBeenCalledTimes(1);
    expect(warn.mock.calls[0]?.[0]).toMatch(/unreadable/i);
  });

  it("does not leave a temp file behind after an atomic write", async () => {
    const store = newStore();
    await store.upsert({ id: "a@x.com", displayId: "a@x.com", credentialSourceId: "app1" }, "{}");
    const { readdir } = await import("node:fs/promises");
    const entries = await readdir(dir);
    expect(entries.some((e) => e.includes(".tmp-"))).toBe(false);
  });
});
