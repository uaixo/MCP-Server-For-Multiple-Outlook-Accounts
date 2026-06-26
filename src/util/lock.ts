/**
 * Cross-process advisory file lock with stale-lock recovery (NFR-SEC-2).
 *
 * Used to guard token-store writes so concurrent writers (a server token
 * refresh and a CLI `connect`) cannot lose updates or expose a partial file.
 * The lock is an exclusively-created lockfile.
 *
 * Staleness is decoupled from the acquire timeout: a lockfile is only stolen
 * once it is older than `staleMs` (presumed-dead holder), which defaults to a
 * value well above any legitimate hold. This prevents stealing the lock from a
 * slow-but-live writer just because a waiter's short `timeoutMs` elapsed.
 */

import { open, stat, unlink } from "node:fs/promises";

/** Default age before a held lockfile is presumed dead and may be stolen. */
export const DEFAULT_STALE_LOCK_MS = 60_000;

function errno(e: unknown): string | undefined {
  return typeof e === "object" && e !== null && "code" in e
    ? String((e as { code?: unknown }).code)
    : undefined;
}

const sleep = (ms: number): Promise<void> => new Promise((resolve) => setTimeout(resolve, ms));

async function acquire(lockPath: string, timeoutMs: number, staleMs: number): Promise<void> {
  const deadline = Date.now() + timeoutMs;
  const payload = JSON.stringify({ pid: process.pid, time: Date.now() });

  for (;;) {
    try {
      const handle = await open(lockPath, "wx"); // O_CREAT | O_EXCL — atomic create.
      try {
        await handle.writeFile(payload);
      } finally {
        await handle.close();
      }
      return;
    } catch (e) {
      if (errno(e) !== "EEXIST") throw e;

      // Lockfile exists — recover it only once it is older than `staleMs`, so a
      // live writer holding the lock briefly is never stolen from.
      try {
        const st = await stat(lockPath);
        if (Date.now() - st.mtimeMs > staleMs) {
          await unlink(lockPath).catch(() => undefined);
          continue; // retry immediately after stealing the stale lock.
        }
      } catch {
        continue; // lock vanished between EEXIST and stat — retry.
      }

      if (Date.now() >= deadline) {
        throw new Error(
          `Timed out after ${timeoutMs}ms waiting for lock ${lockPath}. ` +
            `If no other process is running, delete the stale lockfile.`,
        );
      }
      await sleep(40 + Math.floor(Math.random() * 40)); // jittered poll.
    }
  }
}

async function release(lockPath: string): Promise<void> {
  await unlink(lockPath).catch(() => undefined);
}

/**
 * Run `fn` while holding the lock at `lockPath`, releasing it on completion or
 * error. `timeoutMs` bounds how long to WAIT for the lock; `staleMs` (default
 * {@link DEFAULT_STALE_LOCK_MS}, never below `timeoutMs`) is the age at which a
 * held lock is presumed dead and may be stolen.
 */
export async function withLock<T>(
  lockPath: string,
  timeoutMs: number,
  fn: () => Promise<T>,
  staleMs: number = Math.max(timeoutMs, DEFAULT_STALE_LOCK_MS),
): Promise<T> {
  await acquire(lockPath, timeoutMs, staleMs);
  try {
    return await fn();
  } finally {
    await release(lockPath);
  }
}
