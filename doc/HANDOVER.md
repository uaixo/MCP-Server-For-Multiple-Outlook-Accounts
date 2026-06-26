# Handover

> **Read this first if you're picking the project up.** It's a point-in-time snapshot: what has
> been built, the current state of the tree, and exactly what is left to do. For *how we work*
> (workflow, commands, decisions) see [`CONTRIBUTING.md`](./CONTRIBUTING.md); for *what to build*
> see the spec/design docs in §6 below. This file is the bridge between the two.

---

## 1. TL;DR

- **The offline build is feature-complete.** All eight capabilities (C1–C8) are implemented,
  spec-traced, and tested; two rounds of code review and the §4.3 optional enhancements are done.
- **179 tests pass** (Graph + MSAL mocked). `typecheck` (now covering `src` **and** `test`),
  `build`, and `format:check` are clean; CI is green across ubuntu/macOS/windows × Node 20 & 22,
  a Node-18 runtime smoke, and gitleaks.
- **The only remaining work is the LIVE acceptance run** — real browser consent + real Microsoft
  Graph calls against an Entra app registration and an Outlook mailbox. It can't be done in this
  build environment (no real credentials); see §4.1 for the exact checklist.
- A handful of **documented, by-design limitations** remain (§4.2) and some **optional
  enhancements** are listed (§4.3). None are blockers.

---

## 2. What has been delivered

Built spec-driven, one phase per PR, then two review rounds. All merged to `main`.

| Phase / PR | Scope | Status |
| --- | --- | --- |
| 1 | Auth core: Entra credential discovery, MSAL public-client (consent + silent refresh), secure token store (`0600`/`0700`, atomic + locked), account registry, CLI `connect`/`list`/`remove`, **C1 `list_accounts`** | ✅ merged |
| 2 | Microsoft Graph client (timeout / retry / error mapping), search translation, **C2 `search_conversations`**, **C3 `read_conversation`**, output bounds | ✅ merged |
| 3 (PR #6) | Write path: compose / attachments / sanitize, **C4 `create_draft`**, **C5 `send_message`** with the no-duplicate-send guarantee | ✅ merged |
| 4 (PR #7) | Organise path: **C6 `list_labels`**, **C7 `create_label`**, **C8 `organize_mail`** (the label-decomposition fan-out) + bounded concurrency | ✅ merged |
| 5 (PR #8) | Hardening (secret-redaction boundary, NFR-SEC-6) + operator onboarding docs (`ONBOARDING.md`, CON-3/ASM-1) | ✅ merged |
| Review 1 (PR #9, #10) | Token-egress pinning (SSRF), bounded attachment reads, tool-error redaction, malformed-response mapping; then `O_NOFOLLOW`, token-store `fsync`, lock-staleness decoupling | ✅ merged |
| Review 2 (PR #11, #12) | Browser-spawn crash fix, bounded-concurrency stop-on-error, deterministic organise order, hex HTML entities, **test type-checking in CI**; then per-attachment inline size limit, organise partial-failure docs, dedup | ✅ merged |

Per-requirement status (every FR/NFR/CON → module → test) lives in
[`traceability-matrix.md`](./traceability-matrix.md). It is the authoritative checklist.

---

## 3. Code orientation (where things live)

```
src/
  index.ts                 MCP server: registers all 8 tools, dual-channel result, redaction
  cli/                     out-of-band account CLI (connect/list/remove)
  auth/                    credentialSources, msalClient, tokenProvider, tokenStore, accountRegistry
  graph/                   client (single egress, retry, SSRF-pinned), errors, paginate, types
  search/translate.ts      Gmail-style operators → $search/$filter
  mail/                    sanitize, compose, attachments, replyLookup   (compose+sanitize are PURE)
  organise/decompose.ts    C8 fan-out: intent → Graph ops               (PURE)
  capabilities/            one file per tool (C1–C8) + outgoing.ts (shared write plumbing)
  output/contract.ts       response bounds + size limits
  util/                    bounded (concurrency), lock, html, redact
  domain/                  types + contracts (the architectural seams)
```

**Design rules to preserve** (architecture.md §4): `graph/client` is the *single* egress point;
the **pure** layers (`compose`, `sanitize`, `decompose`, `search`, `output`) do no I/O and carry the
highest-risk logic, so they are unit-tested directly; capabilities stay thin (resolve account →
build request via pure layers → call Graph → shape the dual-channel result).

---

## 4. What is still outstanding

### 4.1 Live acceptance run (the only required remaining work)

Everything testable without real credentials is done and green. What a real Entra app + Outlook
mailbox is needed to confirm (spec §13.2/§13.3; set up via [`ONBOARDING.md`](./ONBOARDING.md)):

- [ ] **Onboarding end-to-end** — `outlook-mcp-auth connect` opens the browser, PKCE + `state`
      round-trip, the loopback callback completes, the account is identified and persisted; `list`
      and `remove` behave; the unverified-app consent path works for your account type (§4 of
      ONBOARDING).
- [ ] **Silent refresh + re-consent without restart** (FR-AUTH-9) — a running server picks up a
      re-`connect`'d cache on its next call.
- [ ] **C2/C3 search/read** — the `$search`/`$filter` translations, `$count` + `ConsistencyLevel`,
      and HTML→text rendering against real mail.
- [ ] **C4/C5 write** — draft creation; `sendMail` actually delivers; **no duplicate** under an
      induced transient failure; **reply threading** — confirm Graph honours the
      `In-Reply-To`/`References` internet headers (see §4.2); the effective outgoing-size limit;
      **a > 3 MB attachment** uploads via the session (createUploadSession + chunked PUTs) and the
      create-draft → upload → send path delivers exactly once.
- [ ] **C6/C7/C8 organise** — category vs. folder listing **including nested folders (full paths)**;
      category/folder creation (incl. nested folder); the category-PATCH / `move` fan-out and the
      conversation-wide **union** report; the `archive` / `trash` (`deleteditems`) / `junk`
      (`junkemail`) well-known moves.
- [ ] **Host gating** — the MCP host actually prompts on the `destructiveHint: true` tools
      (`send_message`, `organize_mail`) (CON-2).

As each passes, flip the corresponding §11 row in `traceability-matrix.md` and note it.

### 4.2 Known limitations (by design / live-confirmed — documented, not bugs)

- **Attachments are capped at the message size (~25 MB).** Files at/under the inline limit
  (`MAX_INLINE_ATTACHMENT_BYTES`, ~3 MB) ride inline as `fileAttachment`; larger ones are uploaded
  to the draft via a **Graph upload session** (`mail/uploadSession.ts`). Per-attachment and whole-
  message size are still bounded locally (`MAX_OUTGOING_MESSAGE_BYTES`). *(Live-confirm the
  effective mailbox limit and that chunked PUTs to the upload URL succeed — §4.1.)*
- **Reply threading uses RFC-5322 `In-Reply-To`/`References` headers** set in `compose.ts`. This is
  the pure, testable approach; whether Graph honours those headers on a structured `POST` is a
  live-confirmed item (§4.1). If Graph drops them, switch to `POST /me/messages/{id}/createReply`
  then patch — see provider-mapping §2.
- **Organise is not transactional.** A conversation fan-out that fails midway leaves earlier
  messages changed and rejects; because C8 is idempotent, **re-running** safely finishes it. The
  tool description says so.
- **Residual attachment TOCTOU on intermediate path segments.** `O_NOFOLLOW` protects the final
  segment; fully closing the window needs per-segment `openat`, which Node doesn't expose portably.
  Low risk (local attacker + write access to the allow-listed dir).
- **CON-2** (the host honours destructive annotations) is the host's responsibility — the server
  only declares the annotations.

### 4.3 Optional enhancements

All three structural enhancements below are now **implemented** (`mail/uploadSession.ts`,
`graph/folders.ts`, `organise/decompose.ts`). What remains is genuinely optional polish.

- ✅ **Large-attachment upload sessions** — `createUploadSession` + chunked PUTs lift the ~3 MB
  inline ceiling toward the message cap (`mail/uploadSession.ts`; C4/C5).
- ✅ **More organise intents** — `archive`, **`trash`** (Deleted Items), and **`junk`** (Junk Email)
  moves are all wired in `decompose.ts` / `organize_mail` (mutually exclusive).
- ✅ **Recursive folder listing in C6** — `graph/folders.ts` enumerates child folders; `list_labels`
  reports each folder's full path (e.g. `Inbox/Clients/Acme`).
- ◻ **Still optional:** richer HTML→text rendering if real mail surfaces gaps; further organise
  destinations beyond archive/trash/junk; per-segment `openat` to close the residual attachment
  TOCTOU window (§4.2).

---

## 5. How to pick up & verify

```bash
npm install
npm run typecheck     # tsc on src AND test (tsconfig.test.json) — type errors in tests now fail CI
npm test              # vitest (Graph + MSAL mocked; no live creds)
npm run build         # compile to dist/
npm run format:check  # prettier
```

- **Node ≥ 20** to run the tests; the *published* server targets Node ≥ 18 (the CI Node-18 job
  smoke-tests only build + CLI).
- **Workflow** (CONTRIBUTING §3): branch from `main` (`claude/<...>` by convention), one PR per
  slice, open as **draft**, let CI go green, mark ready, merge. Two CI gates must pass:
  `.github/workflows/ci.yml` (build/test matrix + typecheck + format) and `secret-scan.yml`
  (gitleaks). Never commit secrets/tokens.
- **Stay spec-driven:** every new module/test cites the FR/NFR it satisfies, and the traceability
  matrix is updated in the same change.

---

## 6. Document map

| Doc | What it's for |
| --- | --- |
| [`business-specification.md`](./business-specification.md) | The provider-neutral **contract** (FR/NFR/CON IDs). What the build must satisfy. |
| [`provider-mapping.md`](./provider-mapping.md) | How each neutral requirement binds to **Microsoft Graph**; §7 watch-items. |
| [`architecture.md`](./architecture.md) | The design: module layout (§3/§4), organise fan-out (§6), search translate (§7), retry/no-dup-send (§8), security model (§9), build roadmap (§13). |
| [`traceability-matrix.md`](./traceability-matrix.md) | Live checklist: every requirement → module → test → status. **Update as you go.** |
| [`ONBOARDING.md`](./ONBOARDING.md) | Operator setup for the live run: Entra app registration, consent, the CLI. |
| [`CONTRIBUTING.md`](./CONTRIBUTING.md) | How we work + the per-phase record of what shipped. |
| **`HANDOVER.md`** (this file) | Point-in-time status + the outstanding-work checklist. |

---

*Bottom line: the build is done and green; the next person's job is the operator live-acceptance
pass (§4.1), optionally followed by the enhancements in §4.3. There are no known open bugs.*
