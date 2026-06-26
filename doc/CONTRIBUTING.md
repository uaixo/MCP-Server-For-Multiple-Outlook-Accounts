# Contributing & Handoff Guide

> Practical guide for anyone picking up this project. The **what to build** lives in the specs and
> design docs; this file captures the **how we work** and the **current pick-up point** so a new
> contributor can continue without prior context.

## 1. Where we are right now

| Phase | Scope | Status |
| --- | --- | --- |
| 1 | Auth core + secure token store + CLI + `list_accounts` (C1) | ✅ merged |
| 2 | Graph client + `search_conversations` (C2) + `read_conversation` (C3) | ✅ merged |
| 3 | Write path — `create_draft` (C4) + `send_message` (C5) | ✅ done |
| 4 | Organise — `list_labels`/`create_label`/`organize_mail` (C6–C8) | ✅ done |
| 5 | Hardening (secret-redaction boundary) + onboarding docs | ✅ done |

**All five build phases are complete.** The offline build is feature-complete (C1–C8 + hardening +
docs); the remaining work is the **live acceptance** runs the operator performs against a real
Entra app registration + Outlook mailbox (spec §13.2/§13.3 — they need real credentials this
environment can't hold). See [`ONBOARDING.md`](./ONBOARDING.md) to set that up.

The authoritative roadmap is [`architecture.md` §13](./architecture.md). Per-requirement status
(done / partial / planned, with the module + test that satisfies each) is in
[`traceability-matrix.md`](./traceability-matrix.md).

## 2. The documents and how they relate

- [`business-specification.md`](./business-specification.md) — provider-neutral **contract**
  (FR/NFR/CON IDs). This is what the build must satisfy.
- [`provider-mapping.md`](./provider-mapping.md) — how each neutral requirement binds to **Microsoft
  Graph** (the build target). Start here for the C4/C5 Graph calls and the §7 watch-items.
- [`architecture.md`](./architecture.md) — the design: module layout (§3), retry classes (§8),
  security model (§9), the organise fan-out table (§6), search translation (§7), build roadmap (§13).
- [`traceability-matrix.md`](./traceability-matrix.md) — the live checklist: every requirement →
  module → test → status.

## 3. How we work (process)

- **Spec-driven:** every module/test cites the FR/NFR ID it satisfies; update the traceability
  matrix as you complete items.
- **One PR per phase** (or smaller slice). Open it as a **draft**, let CI run, mark it **ready**, and
  **merge when green**.
- **Branching:** cut a feature branch from `main` (any name; the automation used
  `claude/<...>`). Never commit straight to `main`.
- **Commits:** clear, descriptive messages. Keep secrets/tokens out of the repo (the gitleaks CI
  will fail the build otherwise).
- After pushing, open a draft PR; address CI/review feedback; merge once all checks pass.

## 4. Local setup & commands

Requires **Node ≥ 20** for development (the test runner, vitest, needs ≥ 20; the *published server*
targets Node ≥ 18).

```bash
npm install
npm run typecheck      # tsc --noEmit
npm test               # vitest (Graph + MSAL are mocked; no live creds)
npm run build          # compile to dist/
npm run format:check   # prettier (run `npm run format` to fix)
```

**CI gates** (both must be green before merge):

- `.github/workflows/ci.yml` — build/test matrix (ubuntu/macOS/windows × Node 20 & 22) + a Node-18
  build/CLI smoke.
- `.github/workflows/secret-scan.yml` — gitleaks.

## 5. Test strategy

Mocked unit + integration. Microsoft Graph is mocked at the `fetch` boundary; MSAL at the client
boundary — so **no live credentials are needed** in CI. **Live** acceptance (real browser consent +
real Graph calls, spec §13) is run **locally by the operator** with a real Entra app registration +
Outlook mailbox; this is where the `$search`/`$filter`, `$count`/`ConsistencyLevel`, and send
behaviour get confirmed against the real API.

## 6. Phase 3 — the write path (C4 `create_draft`, C5 `send_message`) ✅ done

References: `provider-mapping.md` §3 (C4/C5 rows) and §7 item 4 (send semantics / no-duplicate);
`architecture.md` §8 (retry classes) and §9 (attachment guard, header-injection safety). The write
capabilities mirror the existing capability shape in `src/capabilities/readConversation.ts` and the
tool wiring + `toToolResult` envelope in `src/index.ts`.

What shipped (each module + its test, with the requirements it satisfies):

- [x] `src/mail/sanitize.ts` — strip CR/LF (and other controls) from header-bound values
      (recipients, subject, filenames) so a display name/subject can't inject headers. **NFR-SEC-5.**
      → `test/sanitize.test.ts`
- [x] `src/mail/compose.ts` — build the Graph message JSON: parse recipients (`addr` or
      `Display Name <addr>`), `to`/`cc`/`bcc`, subject, body + `is_html`; reply threading via
      `In-Reply-To`/`References`, default `Re:` subject; validate outgoing size locally before the
      call. **FR-C4-1/4/5, NFR-PERF-3.** Uses `sanitize`. → `test/compose.test.ts`
- [x] `src/mail/attachments.ts` — attachment input is **exactly one of** a local `path` **or** inline
      base64. `path` reads are **disabled unless** within `OUTLOOK_MCP_ATTACHMENTS_DIR`; fully resolve
      symlinks/`..` and validate the real path is inside an allowed dir **before** reading; open the
      file **once** and read via the handle (TOCTOU-safe). Infer filename/MIME; require filename for
      inline. **NFR-SEC-3/4, FR-C4-3.** → `test/attachments.test.ts`
- [x] `src/mail/replyLookup.ts` — fetch the conversation's latest message (subject +
      `internetMessageId`) to drive the `Re:` default and threading headers. **FR-C4-4 / FR-C5-1.**
- [x] `src/capabilities/createDraft.ts` (C4) — `POST /me/messages` (creates with `isDraft`);
      attachments as `fileAttachment` resources; **retryClass `nonDuplicable`**; do not send.
      **FR-C4-1..4.** → `test/createDraft.test.ts`
- [x] `src/capabilities/sendMessage.ts` (C5) — `POST /me/sendMail` (single call to avoid an ambiguous
      two-step window); **retryClass `nonDuplicable`** so only a pre-processing 429 is retried —
      never an ambiguous 5xx/timeout → **no duplicate sends**. **FR-C5-1/2/4, NFR-REL-3.** A
      forced-transient-failure test asserts the send is attempted exactly once.
      → `test/sendMessage.test.ts`
- [x] `src/capabilities/outgoing.ts` — shared plumbing (resolve attachments → reply context →
      compose) used by both write tools.
- [x] `src/index.ts` — registers `create_draft` (write, non-destructive) and `send_message`
      (**`destructiveHint: true`**) with zod input schemas + the account selector. **NFR-OPS-4.**
      → `test/toolsAnnotations.test.ts`
- [x] Docs — Phase-3 rows in `traceability-matrix.md` flipped to ✅, README capabilities table +
      status updated, `architecture.md` §13 phase 3 marked done.

> **Note (live-confirmed):** the `FetchGraphClient` was extended to treat 202/empty responses as a
> no-body success (`sendMail` returns 202). Graph's acceptance of the `In-Reply-To`/`References`
> internet headers and the effective outgoing-size limit are confirmed by the operator's live run.

## 6a. Phase 4 — the organise path (C6/C7/C8) ✅ done

References: `provider-mapping.md` §3 (C6/C7 rows) + §3.1 (the C8 fan-out) + §7 item 1;
`architecture.md` §6 (the decomposition table). The decomposition is the core porting risk.

What shipped:

- [x] `src/util/bounded.ts` — bounded-concurrency runner for the per-message fan-out, preserving
      order. **NFR-REL-4.** → `test/bounded.test.ts`
- [x] `src/graph/paginate.ts` — follow `@odata.nextLink` to a hard item cap (shared by C6 + C8).
- [x] `src/capabilities/listLabels.ts` (C6) — combine `GET /me/outlook/masterCategories` (tags) +
      `GET /me/mailFolders` (folders); category id = name, folder id = id; flag system folders.
      **FR-C6-1/2.** → `test/listLabels.test.ts`
- [x] `src/capabilities/createLabel.ts` (C7) — `POST masterCategories` (category) or
      `POST mailFolders` / `…/childFolders` (folder, nestable); `nonDuplicable`. **FR-C7-1.**
      → `test/createLabel.test.ts`
- [x] `src/organise/decompose.ts` — pure fan-out: merge categories[] + isRead into one PATCH, plus
      a `move` op for archive. **FR-C8-6.** → `test/decompose.test.ts`
- [x] `src/capabilities/organizeMail.ts` (C8) — validate exactly-one-target + at-least-one-change;
      resolve a single message or enumerate a conversation; fan out via `decompose` under the
      limiter; report the union of resulting labels. **FR-C8-1..6.** → `test/organizeMail.test.ts`
- [x] `src/index.ts` — register `list_labels` (read), `create_label` (write), `organize_mail`
      (**`destructiveHint: true`, `idempotentHint: true`**). **NFR-OPS-4.**
- [x] Docs — C6/C7/C8 + NFR-REL-4 rows flipped to ✅; README + architecture §13 + this guide updated.

> **Scope note (live-confirmed):** category PATCH replaces the whole `categories[]`, so C8 fetches
> each target's current categories and merges. Top-level folders are listed; nested child folders
> are not recursively enumerated in v1. The `move`/PATCH fan-out and well-known `destinationId:
> "archive"` are confirmed against the real API by the operator.

## 6b. Phase 5 — hardening + onboarding ✅ done

References: spec §10.1 (NFR-SEC-6), §8 + §14 (CON-3 / ASM-1); `architecture.md` §9 (security model)
and §13 (phase 5).

What shipped:

- [x] `src/util/redact.ts` — pure `redact()` / `redactError()` that scrub Bearer tokens, JWTs, and
      `*_token` / `client_secret` / `password` assignments. **NFR-SEC-6.** → `test/redact.test.ts`
- [x] Redaction wired into every stderr boundary that prints an arbitrary error: `index.ts` fatal
      handler, `cli/index.ts` + `cli/connect.ts`, and the `tokenStore` default warn sink. The
      startup banner was extracted to a tested `formatReadyBanner` (identities only).
      → `test/readyBanner.test.ts`
- [x] `doc/ONBOARDING.md` — operator guide: Entra public-client registration, least-privilege
      scopes, `credentials*.json`, the **unverified-app consent policy** (personal vs.
      work/school + admin consent), the connect/list/remove CLI, host wiring, and troubleshooting.
      **CON-3, ASM-1.** Linked from README + traceability matrix.
- [x] Cross-platform check — already enforced by the CI matrix (ubuntu/macOS/windows × Node 20 & 22
      + a Node-18 runtime smoke); no code change needed (NFR-OPS-1).
- [x] Docs — NFR-SEC-6 / CON-3 / ASM-1 rows flipped to ✅; README + architecture §13 + this guide
      updated to "phase 5 done / offline build feature-complete".

> **What's left:** only the operator's **live** acceptance (real browser consent + real Graph
> calls). Everything testable without real credentials is done and green (148 tests).

## 7. Decisions already made (don't relitigate without reason)

- **Graph layer** is a thin `fetch` wrapper (not the official SDK) so the no-duplicate-send retry
  rule is under our control. Use `GraphRequest.retryClass: "nonDuplicable"` for send/draft.
- **Search:** Graph can't combine `$search` and `$filter` in one request — see `search/translate.ts`.
- **Identity** is keyed by the MSAL `account.username` (UPN), lower-cased.
- **Windows CI:** the POSIX file-mode test is skipped on `win32` (Windows uses ACLs); `.gitattributes`
  forces LF so Prettier is stable.
- **Tooling:** TypeScript 6.0.3, ESM (`nodenext`), vitest, Prettier. Node ≥ 20 to run tests.
