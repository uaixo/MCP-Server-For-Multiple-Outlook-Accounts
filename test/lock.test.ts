import { mkdtemp, rm, stat, utimes, writeFile } from "node:fs/promises";
import { tmpdir } from "node:os";
import { join } from "node:path";
import { afterEach, beforeEach, describe, expect, it } from "vitest";
import { withLock } from "../src/util/lock.js";

const sleep = (ms: number): Promise<void> => new Promise((r) => setTimeout(r, ms));

describe("withLock (NFR-SEC-2)", () => {
  let dir: string;
  beforeEach(async () => {
    dir = await mkdtemp(join(tmpdir(), "lock-"));
  });
  afterEach(async () => {
    await rm(dir, { recursive: true, force: true });
  });

  it("runs the critical section and releases the lock afterward", async () => {
    const lock = join(dir, "a.lock");
    await withLock(lock, 1000, async () => undefined);
    await expect(stat(lock)).rejects.toMatchObject({ code: "ENOENT" });
  });

  it("serializes concurrent holders (no interleaving)", async () => {
    const lock = join(dir, "b.lock");
    const order: string[] = [];
    await Promise.all([
      withLock(lock, 2000, async () => {
        order.push("a-start");
        await sleep(40);
        order.push("a-end");
      }),
      withLock(lock, 2000, async () => {
        order.push("b-start");
        await sleep(10);
        order.push("b-end");
      }),
    ]);
    const interleaved =
      order.indexOf("a-end") > order.indexOf("b-start") &&
      order.indexOf("b-end") > order.indexOf("a-start");
    expect(interleaved).toBe(false);
  });

  it("recovers a stale lock older than the timeout", async () => {
    const lock = join(dir, "c.lock");
    await writeFile(lock, "stale");
    const past = new Date(Date.now() - 10_000);
    await utimes(lock, past, past);

    let ran = false;
    await withLock(lock, 1000, async () => {
      ran = true;
    });
    expect(ran).toBe(true);
  });
});
