# Architecture & Design — Multi-Account Outlook / Microsoft 365 MCP Server

> **Status:** Design v0.1 · **Phase:** Scaffold (no capability code yet) · **For review before the full build.**
> **Provider variant of** the provider-neutral [`business-specification.md`](./business-specification.md),
> bound to Microsoft Graph per [`provider-mapping.md`](./provider-mapping.md).
> Requirement traceability (every FR/NFR/CON → module → test) lives in
> [`traceability-matrix.md`](./traceability-matrix.md).

This document describes **how** the Outlook variant is built. The specification is the contract;
this design explains the chosen realisation. Where a decision implements a requirement, the
requirement ID is cited inline.

---

## 1. Scope of this deliverable

This is the **scaffold + design** phase. Delivered now:

- A buildable, green TypeScript project (`tsc` build, `vitest` tests, Prettier) — see §3.
- The neutral **domain types** (`src/domain/types.ts`) and **subsystem contracts**
  (`src/domain/contracts.ts`) — the architectural seams.
- A real, tested **config loader** (`src/config.ts`).
- Minimal **server** and **CLI** entry points that start and report status but register **no
  capability tools** yet.
- This design and the traceability matrix.

**Not** in this deliverable: the capability implementations (C1–C8), the MSAL consent flow, the
token store, the Graph client, and their tests. Those land in the build phases (§13) after this
design is approved.

---

## 2. Technology decisions

| Decision | Choice | Rationale (spec link) |
| --- | --- | --- |
| Language | **TypeScript 6.0.3** (pinned) | Requested; strict typing for a safety-critical tool surface. |
| Runtime | **Node.js ≥ 18**, Windows + macOS | NFR-OPS-1. Verified on Node 22. |
| Module system | **ESM** (`"type": "module"`, `nodenext`) | MCP SDK and MSAL ship ESM; aligns with modern Node. |
| MCP framework | **`@modelcontextprotocol/sdk`** | Official SDK; provides stdio transport (NFR-OPS-2) and per-tool annotations (NFR-OPS-4). |
| Identity | **`@azure/msal-node`** (public client) | provider-mapping §4: MSAL replaces `google-auth-library`; built-in PKCE + token cache. |
| Graph access | **Thin wrapper over `fetch`** (no Graph SDK) | Chosen so the **no-duplicate-send** retry rule (NFR-REL-3) and per-request timeout (NFR-REL-1) are under our exact control; the SDK's retry middleware would otherwise retry ambiguous failures. provider-mapping §6/§7-item-4. |
| Validation | **`zod`** | Tool input schemas + early validation errors before any provider call (FR-ERR-3). |
| Tests | **`vitest`**, **Graph/MSAL mocked** | Offline unit + integration coverage of the §13 criteria that don't need live Azure. Live sandbox acceptance is run by the operator locally. |

> **Why no official Graph SDK?** NFR-REL-3 forbids retrying `send`/`draft-create` on ambiguous
> failures (a 5xx or timeout that may already have delivered). The SDK's retry middleware does not
> distinguish "pre-processing 429" from "ambiguous failure" the way the spec requires, so we own
> the HTTP layer. See §8.

---

## 3. Project layout

Implemented now (scaffold) is marked ✓; planned modules are marked ◻.

```
src/
  index.ts                 ✓ MCP server entry (stdio; registers tools; reports accounts to stderr) — NFR-OPS-2
  config.ts                ✓ env-var config loader (§12 / mapping §5)
  domain/
    types.ts               ✓ neutral domain model (§4)
    contracts.ts           ✓ subsystem interfaces (the seams)
  cli/
    index.ts               ✓ account-management CLI dispatcher (§8)
    connect.ts             ◻ PKCE+state loopback consent, GET /me, token store  (FR-AUTH-1..7,10)
    list.ts                ◻ list accounts + credential source                  (FR-AUTH-8)
    remove.ts              ◻ remove account                                      (FR-AUTH-8)
  auth/
    credentialSources.ts   ◻ discover one/many app-registration configs         (FR-ID-5/6)
    msalClient.ts          ◻ MSAL public client per credential source; silent refresh (FR-AUTH-5/9)
    tokenStore.ts          ◻ secure store: 600/700, atomic temp+rename, locked  (NFR-SEC-1/2, FR-ERR-2)
    accountRegistry.ts     ◻ selector resolution + default rule                 (FR-ID-1..4)
  graph/
    client.ts              ◻ fetch wrapper: timeout + retry + error mapping      (NFR-REL-1/2, FR-ERR-1)
    retry.ts               ◻ retry classes (safe vs nonDuplicable), backoff      (NFR-REL-2/3)
    errors.ts              ◻ Graph error code → actionable category              (FR-ERR-1)
  search/
    translate.ts           ◻ operator subset → $search + $filter                (FR-C2-1, mapping §7-2)
  organise/
    decompose.ts           ◻ neutral intent → categories/move/isRead fan-out     (FR-C8-6, mapping §3.1)
  mail/
    compose.ts             ◻ message build + recipient parsing                   (FR-C4-1/5)
    attachments.ts         ◻ allow-list path guard + TOCTOU-safe read            (NFR-SEC-3/4)
    sanitize.ts            ◻ CR/LF header-injection stripping                     (NFR-SEC-5)
  output/
    contract.ts           ◻ dual-channel result + char budget + truncation      (FR-OUT-*, NFR-PERF-1/2)
  util/
    lock.ts               ◻ cross-process file lock w/ stale recovery            (NFR-SEC-2)
    bounded.ts            ◻ concurrency limiter                                   (NFR-REL-4, FR-C2-6)
  capabilities/
    listAccounts.ts (C1) ◻   searchConversations.ts (C2) ◻   readConversation.ts (C3) ◻
    createDraft.ts (C4)  ◻   sendMessage.ts (C5)         ◻   listLabels.ts (C6)       ◻
    createLabel.ts (C7)  ◻   organizeMail.ts (C8)        ◻
test/
  config.test.ts           ✓ scaffold smoke test (config loader)
  ...                      ◻ mocked unit + integration suites per module
```

---

## 4. Layered architecture

```
            ┌──────────────────────────────────────────────┐
  MCP host  │  index.ts  — registers 8 tools w/ annotations  │
  (stdio)──▶│            — dual-channel results (§11)         │
            └───────────────┬──────────────────────────────┘
                            │ calls
            ┌───────────────▼──────────────────────────────┐
            │  capabilities/  (C1–C8)                        │  ← thin orchestration;
            │  - resolve account selector (AccountRegistry)  │    one file per tool
            │  - validate inputs (zod) before any call       │
            └───┬───────────────┬───────────────┬───────────┘
                │               │               │
      ┌─────────▼───┐  ┌────────▼────────┐  ┌───▼─────────────┐
      │ search/      │  │ organise/        │  │ mail/            │  ← pure translation /
      │ translate    │  │ decompose        │  │ compose/attach   │    composition (no I/O)
      └─────────┬───┘  └────────┬────────┘  └───┬─────────────┘
                └───────────────┼───────────────┘
                      ┌─────────▼─────────┐
                      │ graph/client       │  ← the ONLY outbound HTTP path
                      │ timeout+retry+errs │
                      └─────────┬─────────┘
                                │ uses tokens from
                      ┌─────────▼─────────┐
                      │ auth/ (MSAL +      │  ← refresh tokens; per-account
                      │ tokenStore + reg.) │    client binding (FR-ID-5)
                      └───────────────────┘
```

**Design rules**
- The **pure** layers (`search`, `organise`, `mail/compose`, `mail/sanitize`, `output`) do no I/O,
  so they are unit-tested directly without mocks — they carry the highest-risk logic.
- `graph/client` is the **single egress point**; timeout, retry, and error mapping live there once.
- Capabilities are **thin**: resolve account → build request(s) via pure layers → call Graph →
  shape the dual-channel result. This keeps each FR traceable to a small surface.

---

## 5. Identity, account selection & onboarding

### 5.1 Account selection rule (spec §7)
`AccountRegistry.resolve(selector?)` is the single chokepoint enforcing FR-ID-2/3/4:

| Connected accounts | selector | Result |
| --- | --- | --- |
| 0 | any | actionable error: "no accounts — run `outlook-mcp-auth connect`" (FR-ID-2, FR-C1-2) |
| 1 | omitted | use the one account (FR-ID-2) |
| ≥2 | omitted | disambiguation error listing connected accounts (FR-ID-2) |
| any | named, known (case-insensitive) | use it (FR-ID-4) |
| any | named, unknown | error listing connected accounts (FR-ID-3) |

Every tool except `list_accounts` calls `resolve` first.

### 5.2 Onboarding flow (CLI `connect`, spec §8 / mapping §4)
```
connect
  → pick credential source (app registration): --credentials path, or auto-discovered  (FR-ID-5/6)
  → MSAL public client; build authorize URL with PKCE(S256) + random state             (FR-AUTH-3/4)
  → start short-lived loopback server on 127.0.0.1 (http://localhost, MSAL port)        (FR-AUTH-2, NFR-SEC-7)
  → open browser; user consents (scopes: Mail.ReadWrite, Mail.Send, User.Read, offline_access) (FR-AUTH-10)
  → callback: verify state (forged/unrelated callbacks answered neutrally, genuine flow survives) (FR-AUTH-4)
  → exchange code (PKCE verifier) → access + refresh token (offline_access)             (FR-AUTH-5)
  → GET /me → userPrincipalName / mail → account identity                               (FR-AUTH-6)
  → persist MSAL cache to token store, keyed lower-cased, bound to this app registration (FR-ID-4/5)
  → bounded wait (reference 5 min) else fail with guidance                              (FR-AUTH-7)
```
Re-running `connect` after a refresh-token revocation **repairs** the account; a running server
re-reads the cache on its next call, so it picks up the new credentials **without restart**
(FR-AUTH-9 — `TokenStore.readCache` is called per request, never cached in memory across the boundary).

---

## 6. The label-decomposition fan-out (C8) — highest porting risk

Per the concept-decomposition rule (spec §4) and FR-C8-6, **one** neutral `organize_mail` request
maps to a **combination** of Graph operations. `organise/decompose.ts` owns this table
(provider-mapping §3.1):

| Neutral intent | Graph operation(s) |
| --- | --- |
| Add user tag | `PATCH /me/messages/{id}` → add to `categories[]` |
| Remove user tag | `PATCH /me/messages/{id}` → remove from `categories[]` |
| Mark read / unread | `PATCH /me/messages/{id}` → `isRead = true/false` |
| Archive (remove from inbox) | `POST /me/messages/{id}/move` → destination = Archive folder |
| Trash | `POST /me/messages/{id}/move` → Deleted Items (or `DELETE`) |
| Junk/spam | `POST /me/messages/{id}/move` → Junk Email |
| Apply to a whole **conversation** | enumerate the conversation's message ids, apply per message, then report the **union** of resulting labels (FR-C8-4) — Graph has no single conversation-modify |

Validation (FR-C8-1/2): exactly one target (conversation **or** message), and at least one change,
both enforced **before** any Graph call (FR-ERR-3). The tool is annotated `destructiveHint: true`
(FR-C8-5 / NFR-OPS-4). Idempotent (FR-C8 annotation): re-applying the same add/remove set converges.

---

## 7. Search translation (C2) — second porting risk

`search/translate.ts` maps a documented **subset** of Gmail-style operators to Graph `$search` +
OData `$filter` (provider-mapping §7-item-2):

| Input operator | Graph translation |
| --- | --- |
| free text | `$search="..."` |
| `is:unread` / `is:read` | `$filter=isRead eq false` / `eq true` |
| `from:x` | `$filter=from/emailAddress/address eq 'x'` |
| `to:x` | `$filter=toRecipients/any(r:r/emailAddress/address eq 'x')` |
| `subject:x` | `$search="subject:x"` |
| `has:attachment` | `$filter=hasAttachments eq true` |
| `after:`/`before:` | `$filter=receivedDateTime ge/le <iso>` |

Unsupported operators are **rejected predictably** with an actionable message (FR-ERR-1/3) rather
than silently ignored. Results are grouped by `conversationId`; pagination returns Graph's
`@odata.nextLink` **opaquely** as the neutral cursor (FR-C2-3, mapping §6) — we never parse
`$skiptoken`. Per-conversation detail fetches are bounded by the concurrency limiter (FR-C2-6,
NFR-REL-4); a conversation whose summary fails is **degraded** with an `error` field rather than
failing the whole search (FR-C2-5).

---

## 8. Graph client: timeout, retry & no-duplicate-send

`graph/client.ts` + `graph/retry.ts` are the single egress point.

- **Timeout (NFR-REL-1):** every request runs under an `AbortSignal` bounded by
  `requestTimeoutMs` (default 30s). A stalled socket fails fast as a timeout (FR-ERR-1).
- **Retry classes (NFR-REL-2/3):** each request declares a `RetryClass`:
  - `safe` (read / organise): retry on HTTP 429, transient 5xx, and transport errors, with
    **bounded jittered backoff**, honouring `Retry-After` (mapping §6).
  - `nonDuplicable` (`send_message`, `create_draft`): retry **only** on a *pre-processing* 429
    (request rejected before any side effect). **Never** retry a 5xx or timeout that may have
    already delivered. This is the mechanism behind acceptance criterion §13.4 (no duplicate sends).
- **Error mapping (FR-ERR-1):** `graph/errors.ts` maps Graph `code`/`message` to actionable
  categories: auth expired → "re-connect this account"; 429 → "try again shortly"; timeout →
  reported as timeout; everything else → a safe generic message. Never raised as a crash.

**Send strategy (mapping §7-item-4):** v1 uses `POST /me/sendMail` (single call) so there is no
ambiguous two-step window; combined with the `nonDuplicable` policy this gives the no-duplicate
guarantee (FR-C5-4).

---

## 9. Security model

| Control | Design | Req |
| --- | --- | --- |
| Local secrets | Token store file `chmod 600` inside a `chmod 700` data dir; nothing transits a third party. | NFR-SEC-1 |
| Atomic, locked writes | Write temp file → `fsync` (via an open handle) → `rename`; guard with a cross-process lock (`util/lock.ts`). Lock **staleness is decoupled** from the acquire timeout (`staleMs`, default 60s) so a slow-but-live writer is never stolen from. | NFR-SEC-2 |
| Attachment path guard | `path` attachments **disabled** unless `OUTLOOK_MCP_ATTACHMENTS_DIR` is set; the resolved real path (symlinks, `..`) must fall inside an allowed dir **before** reading. Inline base64 always works. | NFR-SEC-3 |
| TOCTOU-safe read | Open the resolved file **once** with `O_NOFOLLOW` (where supported) so the final segment can't be swapped for a symlink after resolution; validate + read via that handle — no check-then-reopen window. | NFR-SEC-4 |
| Header-injection safe | Strip CR/LF from recipients, subject, filenames (`mail/sanitize.ts`) so a display name / subject can't inject headers. | NFR-SEC-5 |
| No secret logging | Never log tokens, credentials, or message bodies; `util/redact` scrubs secret shapes at every stderr boundary; tool-error text is redacted too. | NFR-SEC-6 |
| Token-egress pinning | `graph/client` attaches the access token only to requests whose origin is `graph.microsoft.com`; an absolute URL (e.g. a `page_token` cursor) pointing elsewhere is refused, so the token can't be exfiltrated via a crafted URL. | NFR-SEC-6 |
| Consent hardening | Loopback callback binds loopback-only, is short-lived, enforces PKCE + `state`. | NFR-SEC-7 |

---

## 10. Output contract & resource bounds

`output/contract.ts` produces the dual channel for every tool (FR-OUT-1): a concise human summary
plus the authoritative structured object (FR-OUT-2). Stable field names: `account`, `count`,
`conversation_id`, `message_id`, `next_page_token`, `truncated`, `omitted_message_count`
(FR-OUT-3). Bounds: total response text ≤ ~25,000 chars (NFR-PERF-1); conversation reads cap at
~100 messages and ~20,000 body chars, keeping the **newest**, setting `truncated` and
`omitted_message_count` (NFR-PERF-2 / FR-C3-2). Outgoing size validated locally against the
effective mailbox limit before the API call (NFR-PERF-3).

---

## 11. Tool registration & annotations (NFR-OPS-4)

Registered in `index.ts` during the build phase. Annotations gate destructive actions in the host:

| Tool | readOnly | destructive | idempotent | openWorld |
| --- | --- | --- | --- | --- |
| `list_accounts` | ✓ | ✗ | ✓ | ✗ (closed-world) |
| `search_conversations` | ✓ | ✗ | ✓ | ✓ |
| `read_conversation` | ✓ | ✗ | ✓ | ✓ |
| `create_draft` | ✗ | ✗ (reversible) | ✗ | ✓ |
| `send_message` | ✗ | **✓** | ✗ | ✓ |
| `list_labels` | ✓ | ✗ | ✓ | ✓ |
| `create_label` | ✗ | ✗ (additive) | ✗ | ✓ |
| `organize_mail` | ✗ | **✓** | ✓ | ✓ |

`send_message` and `organize_mail` are `destructiveHint: true` (spec §6.5/§6.8, NFR-OPS-4).

---

## 12. Configuration (spec §12 / mapping §5)

Loaded by `config.ts` (implemented + tested now):

| Env var | Meaning | Default |
| --- | --- | --- |
| `OUTLOOK_MCP_DATA_DIR` | tokens + app-registration configs | `~/.outlook-mcp` |
| `OUTLOOK_OAUTH_CREDENTIALS` | pin one app registration; disables discovery | unset (auto-discover) |
| `OUTLOOK_MCP_ATTACHMENTS_DIR` | allow-list for `path` attachments | unset (`path` disabled) |
| `OUTLOOK_MCP_LOCK_TIMEOUT_MS` | token-store lock wait | `12000` |
| `OUTLOOK_MCP_REQUEST_TIMEOUT_MS` | per-Graph-call timeout | `30000` |

---

## 13. Test strategy & build roadmap

**Test strategy (this build):** mocked unit + integration with `vitest`. Graph is mocked at the
`fetch` boundary; MSAL is mocked at the client boundary. Offline-coverable §13 criteria: account
selection (§13.1), capability behaviour against mocked Graph (§13.2), no-duplicate-send under
forced transient failure (§13.4), attachment guard (§13.5), annotations (§13.6), timeout/retry
(§13.7), token file modes (§13.8), bounded output (§13.9). **Live** §13.2/§13.3 against a real
Outlook sandbox + Entra app registration are run by the operator locally (requires real
credentials this environment cannot hold).

**Build phases:**
1. ✅ **Auth core (done)** — credential sources, MSAL client, secure token store (+lock), account
   registry, CLI `connect`/`list`/`remove`, and **C1** `list_accounts`. (FR-AUTH-*, FR-ID-*,
   NFR-SEC-1/2) — 29 tests, Graph/MSAL mocked.
2. ✅ **Read path (done)** — Graph client (timeout/retry/errors), search translate, **C2/C3**,
   output contract + bounds. (FR-C2/C3, NFR-REL-1/2, NFR-PERF-*) — 71 tests, Graph/MSAL mocked.
3. ✅ **Write path (done)** — compose/attachments/sanitize, **C4** `create_draft` + **C5**
   `send_message` with a single-call `sendMail` under the `nonDuplicable` retry class (no duplicate
   sends), allow-listed/TOCTOU-safe attachments, header-injection stripping, and local
   outgoing-size validation. (FR-C4/C5, NFR-SEC-3/4/5, NFR-PERF-3, NFR-REL-3) — 113 tests,
   Graph/MSAL mocked.
4. ✅ **Organise path (done)** — **C6** `list_labels` (categories + folders), **C7** `create_label`
   (category or folder, with nesting), and **C8** `organize_mail` with the label-decomposition
   fan-out (category PATCH / `move` / read-state) applied per message under a bounded concurrency
   limit, reporting the conversation-wide union. (FR-C6/C7/C8, NFR-REL-4) — 139 tests, Graph/MSAL
   mocked.
5. ✅ **Hardening & docs (done)** — a secret-redaction boundary (`util/redact`) enforcing
   NFR-SEC-6 across every stderr log site, the operator onboarding guide (`doc/ONBOARDING.md`,
   CON-3 + ASM-1), and the cross-platform check carried by the CI matrix (ubuntu/macOS/windows ×
   Node 20 & 22 + a Node-18 smoke). — 148 tests, Graph/MSAL mocked. The offline build is
   feature-complete; only the operator's **live** acceptance runs remain.

---

## 14. Constraints & assumptions carried (spec §14)

- **CON-1:** this build is the Outlook provider only; one server = one provider.
- **CON-2:** the host acts on destructive annotations; the server only declares them.
- **CON-3:** unverified Entra apps may need admin consent / listed test users — covered in
  onboarding docs (build phase 5).
- **CON-4:** large binaries use the allow-listed `path` mechanism, not inline base64.
- **ASM-1/2:** operator can register an Entra app and run the one-time CLI; egress to Graph/Entra
  endpoints is available.
```
