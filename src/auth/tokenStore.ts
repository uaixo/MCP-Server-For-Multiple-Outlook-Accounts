/**
 * Secure, local-only token store (NFR-SEC-1/2, FR-ERR-2).
 *
 * - Stored only on the user's machine: the data dir is created owner-only
 *   (0700) and the store file is written owner-only (0600). (NFR-SEC-1)
 * - Writes are atomic (temp file + rename) and guarded by a cross-process lock
 *   with stale-lock recovery, so a server refresh and a CLI connect cannot
 *   clobber each other or expose a partial file. (NFR-SEC-2)
 * - A corrupt/malformed store never crashes the server: it is treated as "no
 *   accounts" with a one-time warning explaining how to repair it. (FR-ERR-2)
 *
 * Each account owns its serialized MSAL token cache, bound (via
 * `credentialSourceId`) to the app registration that authorised it, so refresh
 * always uses the issuing client (FR-ID-5). The cache is re-read from disk on
 * each access so a re-consent is picked up without a server restart (FR-AUTH-9).
 */

import { chmod, mkdir, readFile, rename, writeFile } from "node:fs/promises";
import { randomBytes } from "node:crypto";
import { join } from "node:path";
import type { Account } from "../domain/types.js";
import type { TokenStore } from "../domain/contracts.js";
import { withLock } from "../util/lock.js";
import { redact } from "../util/redact.js";

const STORE_VERSION = 1 as const;
const FILE_MODE = 0o600;
const DIR_MODE = 0o700;

interface StoredAccount {
  id: string; // lower-cased identity key (FR-ID-4)
  displayId: string; // identity as authenticated (original case)
  credentialSourceId: string; // issuing app registration (FR-ID-5)
  cache: string; // serialized MSAL token cache for this account
}

interface StoreFile {
  version: typeof STORE_VERSION;
  accounts: Record<string, StoredAccount>;
}

function emptyStore(): StoreFile {
  return { version: STORE_VERSION, accounts: {} };
}

function isStoreFile(value: unknown): value is StoreFile {
  if (typeof value !== "object" || value === null) return false;
  const v = value as Record<string, unknown>;
  return v.version === STORE_VERSION && typeof v.accounts === "object" && v.accounts !== null;
}

export interface TokenStoreOptions {
  readonly dataDir: string;
  readonly lockTimeoutMs: number;
  /** Sink for the one-time corrupt-store warning (defaults to stderr). Never receives secrets (NFR-SEC-6). */
  readonly warn?: (message: string) => void;
}

export class FileTokenStore implements TokenStore {
  private readonly filePath: string;
  private readonly lockPath: string;
  private readonly warn: (message: string) => void;
  private corruptWarned = false;

  constructor(private readonly opts: TokenStoreOptions) {
    this.filePath = join(opts.dataDir, "tokens.json");
    this.lockPath = `${this.filePath}.lock`;
    // The default sink redacts before writing so no secret can reach stderr (NFR-SEC-6).
    this.warn = opts.warn ?? ((m) => process.stderr.write(`${redact(m)}\n`));
  }

  async list(): Promise<Account[]> {
    const store = await this.read();
    return Object.values(store.accounts).map((a) => ({
      id: a.id,
      displayId: a.displayId,
      credentialSourceId: a.credentialSourceId,
    }));
  }

  async readCache(accountId: string): Promise<string | undefined> {
    const store = await this.read();
    return store.accounts[accountId.toLowerCase()]?.cache;
  }

  async upsert(account: Account, serializedCache: string): Promise<void> {
    const key = account.id.toLowerCase();
    await this.mutate((store) => {
      store.accounts[key] = {
        id: key,
        displayId: account.displayId,
        credentialSourceId: account.credentialSourceId,
        cache: serializedCache,
      };
    });
  }

  async remove(accountId: string): Promise<void> {
    const key = accountId.toLowerCase();
    await this.mutate((store) => {
      delete store.accounts[key];
    });
  }

  /** Read-modify-write under the cross-process lock with an atomic replace (NFR-SEC-2). */
  private async mutate(change: (store: StoreFile) => void): Promise<void> {
    await this.ensureDir();
    await withLock(this.lockPath, this.opts.lockTimeoutMs, async () => {
      const store = await this.read();
      change(store);
      await this.atomicWrite(store);
    });
  }

  private async ensureDir(): Promise<void> {
    await mkdir(this.opts.dataDir, { recursive: true, mode: DIR_MODE });
    // mkdir's mode is masked by umask; enforce owner-only explicitly (NFR-SEC-1).
    await chmod(this.opts.dataDir, DIR_MODE).catch(() => undefined);
  }

  private async read(): Promise<StoreFile> {
    let raw: string;
    try {
      raw = await readFile(this.filePath, "utf8");
    } catch (e) {
      if (
        typeof e === "object" &&
        e !== null &&
        "code" in e &&
        (e as { code?: unknown }).code === "ENOENT"
      ) {
        return emptyStore();
      }
      throw e;
    }
    try {
      const parsed: unknown = JSON.parse(raw);
      if (!isStoreFile(parsed)) throw new Error("unexpected shape");
      return parsed;
    } catch {
      this.warnCorruptOnce();
      return emptyStore(); // FR-ERR-2: never crash on a corrupt store.
    }
  }

  private async atomicWrite(store: StoreFile): Promise<void> {
    const tmp = `${this.filePath}.tmp-${randomBytes(6).toString("hex")}`;
    await writeFile(tmp, JSON.stringify(store, null, 2), { mode: FILE_MODE });
    await chmod(tmp, FILE_MODE).catch(() => undefined);
    await rename(tmp, this.filePath); // atomic replace on the same filesystem.
    await chmod(this.filePath, FILE_MODE).catch(() => undefined);
  }

  private warnCorruptOnce(): void {
    if (this.corruptWarned) return;
    this.corruptWarned = true;
    this.warn(
      `[outlook-mcp] Token store at ${this.filePath} is unreadable; treating as no accounts. ` +
        `Re-run "outlook-mcp-auth connect", or delete the file to reset.`,
    );
  }
}
