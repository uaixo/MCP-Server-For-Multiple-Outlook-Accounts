# Requirements Traceability Matrix — Outlook / Microsoft 365 MCP Server

> **Companion to** [`business-specification.md`](./business-specification.md),
> [`provider-mapping.md`](./provider-mapping.md), and [`architecture.md`](./architecture.md).
> Every requirement ID in the spec maps here to the module(s) that will satisfy it and the
> test(s) that will prove it. This is the spec-driven control: each delivered behaviour traces
> back to a requirement, and each requirement forward to code + test.

**Status legend**

- ✅ **Done** — implemented and tested in this repo.
- 🟡 **Partial** — implemented for the current slice / type or contract exists; more pending.
- ◻ **Planned** — design fixed (see architecture.md); not yet coded.

**Current phase:** **Phase 2 complete** — the read path: Microsoft Graph client (timeout / retry /
error-mapping), search translation, and **C2 `search_conversations`** + **C3 `read_conversation`**,
on top of phase 1 (auth core + **C1 `list_accounts`**). **71 tests** pass (Graph/MSAL mocked).
Write (C4/C5) and organise (C6–C8) capabilities remain ◻ Planned (phases 3–4 in architecture.md §13).

> **Delegated-to-MSAL note:** the loopback redirect, PKCE (S256), and CSRF `state`
> (FR-AUTH-2/3/4, NFR-SEC-7) are satisfied by MSAL's `acquireTokenInteractive` and marked
> ✅ "(MSAL)".

---

## 1. Functional requirements — capabilities (spec §6)

| Req | Summary | Module(s) | Test(s) | Status |
| --- | --- | --- | --- | --- |
| FR-C1-1 | List connected accounts as identities | `auth/tokenStore`, `capabilities/listAccounts` | `listAccounts.test`, `tokenStore.test` | ✅ |
| FR-C1-2 | Empty → non-error connect guidance | `capabilities/listAccounts` | `listAccounts.test` | ✅ |
| FR-C1-3 | Output feeds other tools' selectors | `capabilities/listAccounts` | `listAccounts.test` | ✅ |
| FR-C2-1 | Required query in native syntax | `search/translate`, `capabilities/searchConversations` | `translate.test`, `searchConversations.test` | ✅ |
| FR-C2-2 | Page size default 20, max 100 | `output/contract` (`clampPageSize`), `capabilities/searchConversations` | `outputContract.test`, `searchConversations.test` | ✅ |
| FR-C2-3 | Opaque next-page cursor (`@odata.nextLink`) | `capabilities/searchConversations` | `searchConversations.test` | ✅ |
| FR-C2-4 | Summary: id, subject, sender, date, snippet | `capabilities/searchConversations`, `graph/types` | `searchConversations.test` | ✅ |
| FR-C2-5 | Degrade a failed entry, not whole search | `capabilities/searchConversations` | `searchConversations.test` | ✅ (n/a: summaries come from the list response, no per-entry fetch to fail) |
| FR-C2-6 | Bound per-conversation fan-out | `capabilities/searchConversations` | `searchConversations.test` | ✅ (no fan-out: single `GET /me/messages`) |
| FR-C3-1 | Full conversation: headers, body, labels | `capabilities/readConversation`, `graph/types` | `readConversation.test` | ✅ |
| FR-C3-2 | Cap messages (newest) + body chars; `truncated` | `output/contract`, `capabilities/readConversation` | `readConversation.test` | ✅ |
| FR-C3-3 | HTML → readable plain text | `util/html`, `capabilities/readConversation` | `outputContract.test`, `readConversation.test` | ✅ |
| FR-C4-1 | Compose to/cc/bcc/subject/body/is_html | `mail/compose`, `capabilities/createDraft` | `compose.test` | ◻ |
| FR-C4-2 | Persist draft, do not send | `capabilities/createDraft` | `createDraft.test` | ◻ |
| FR-C4-3 | Attachment: path XOR inline base64 | `mail/attachments` | `attachments.test` | ◻ |
| FR-C4-4 | Reply drafting + `Re:` default subject | `mail/compose`, `capabilities/createDraft` | `compose.test` | ◻ |
| FR-C4-5 | Recipient forms + injection prevention | `mail/compose`, `mail/sanitize` | `sanitize.test` | ◻ |
| FR-C5-1 | Send: same inputs + reply-to ref | `capabilities/sendMessage` | `sendMessage.test` | ◻ |
| FR-C5-2 | Deliver immediately (irreversible) | `capabilities/sendMessage` | `sendMessage.test` | ◻ |
| FR-C5-3 | Annotated destructive | `index` (tool registration) | `tools.annotations.test` | ◻ |
| FR-C5-4 | No duplicate delivery under retry | `graph/retry`, `capabilities/sendMessage` | `graphRetry.test` (policy ✅), capability ◻ | 🟡 (policy done; send capability phase 3) |
| FR-C6-1 | List labels: id, name, type | `capabilities/listLabels` | `listLabels.test` | ◻ |
| FR-C6-2 | Output is id source for C8 | `capabilities/listLabels` | `listLabels.test` | ◻ |
| FR-C7-1 | Create user label (+nesting via folders) | `capabilities/createLabel` | `createLabel.test` | ◻ |
| FR-C8-1..6 | Organise mail (fan-out: category/move/isRead) | `organise/decompose`, `capabilities/organizeMail` | `decompose.test` | ◻ |

## 2. Account & identity model (spec §7)

| Req | Summary | Module(s) | Test(s) | Status |
| --- | --- | --- | --- | --- |
| FR-ID-1 | Optional account selector on every tool but C1 | `auth/accountRegistry`, `capabilities/search|readConversation` | `accountRegistry.test`, capability tests | ✅ (C2/C3 take the selector) |
| FR-ID-2 | Default rule (1 → use; ≥2 → disambiguate; 0 → guide) | `auth/accountRegistry` | `accountRegistry.test` | ✅ |
| FR-ID-3 | Unknown account → error listing accounts | `auth/accountRegistry` | `accountRegistry.test` | ✅ |
| FR-ID-4 | Case-insensitive identity (lower-cased key) | `auth/accountRegistry`, `auth/tokenStore` | `accountRegistry.test`, `tokenStore.test` | ✅ |
| FR-ID-5 | Refresh uses the issuing app registration | `auth/tokenProvider`, `auth/msalClient`, `auth/credentialSources` | `msalClient.test` | ✅ |
| FR-ID-6 | Multiple OAuth clients, auto-discovered | `auth/credentialSources` | `credentialSources.test` | ✅ |

## 3. Authentication & onboarding (spec §8)

| Req | Summary | Module(s) | Test(s) | Status |
| --- | --- | --- | --- | --- |
| FR-AUTH-1 | CLI connect via auth-code browser consent | `cli/connect`, `auth/msalClient` | `msalClient.test` | ✅ |
| FR-AUTH-2 | Loopback redirect on 127.0.0.1 (+fallback) | `auth/msalClient` (MSAL) | — | ✅ (MSAL) |
| FR-AUTH-3 | PKCE (S256) | `auth/msalClient` (MSAL) | — | ✅ (MSAL) |
| FR-AUTH-4 | CSRF `state` round-trip; neutral on forged | `auth/msalClient` (MSAL) | — | ✅ (MSAL) |
| FR-AUTH-5 | Offline access → refresh token | `auth/msalClient` | `msalClient.test` | ✅ |
| FR-AUTH-6 | Identify account (authenticated identity) | `cli/connect`, `auth/msalClient` | `msalClient.test` | ✅ |
| FR-AUTH-7 | Bounded consent wait (ref 5 min) | `auth/msalClient` (`withTimeout`) | `msalClient.test` | ✅ |
| FR-AUTH-8 | CLI list (w/ source) + remove | `cli/list`, `cli/remove`, `auth/tokenStore` | `tokenStore.test` | ✅ |
| FR-AUTH-9 | Re-consent repair, no restart needed | `auth/tokenStore`, `auth/tokenProvider` (persists rotated cache) | `tokenStore.test` | ✅ |
| FR-AUTH-10 | Least-privilege scopes | `auth/credentialSources` | `credentialSources.test` | ✅ |

## 4. Error handling & resilience (spec §9)

| Req | Summary | Module(s) | Test(s) | Status |
| --- | --- | --- | --- | --- |
| FR-ERR-1 | Provider errors → actionable messages | `graph/errors`, `auth/accountRegistry`, `index` (tool-error envelope) | `graphErrors.test`, `accountRegistry.test` | ✅ |
| FR-ERR-2 | Corrupt token store → "no accounts" + warning | `auth/tokenStore` | `tokenStore.test` | ✅ |
| FR-ERR-3 | Validation errors before any provider call | `search/translate`, `capabilities/search|readConversation`, `index` (zod) | `translate.test`, capability tests | 🟡 (read path done; write/organise phases 3–4) |

## 5. Output / response contract (spec §11)

| Req | Summary | Module(s) | Test(s) | Status |
| --- | --- | --- | --- | --- |
| FR-OUT-1 | Dual channel: human text + structured | `domain/types` (`ToolResult`), `index` (envelope), all capabilities | capability tests | ✅ |
| FR-OUT-2 | Structured authoritative when text truncated | `output/contract` (`clampText`) | `outputContract.test` | ✅ |
| FR-OUT-3 | Stable, documented field names | capabilities (`account`, `count`, `conversation_id`, `next_page_token`, `truncated`, `omitted_message_count`) | capability tests | ✅ |

## 6. Non-functional — security & privacy (spec §10.1)

| Req | Summary | Module(s) | Test(s) | Status |
| --- | --- | --- | --- | --- |
| NFR-SEC-1 | Tokens local only; 600 file in 700 dir | `auth/tokenStore` | `tokenStore.test` | ✅ |
| NFR-SEC-2 | Atomic temp+rename, cross-process lock | `auth/tokenStore`, `util/lock` | `lock.test`, `tokenStore.test` | ✅ |
| NFR-SEC-3 | Attachment path guard (allow-list, resolve) | `mail/attachments`, `config` | `attachments.test` | 🟡 (config done; reader phase 3) |
| NFR-SEC-4 | TOCTOU-safe attachment read | `mail/attachments` | `attachments.test` | ◻ |
| NFR-SEC-5 | Strip CR/LF from header-bound values | `mail/sanitize` | `sanitize.test` | ◻ |
| NFR-SEC-6 | Never log tokens/credentials/content | all modules (errors carry no secrets) | `graphErrors.test` (messages are generic) | 🟡 (policy followed; dedicated assertion later) |
| NFR-SEC-7 | Consent loopback hardening (PKCE+state) | `auth/msalClient` (MSAL) | — | ✅ (MSAL) |

## 7. Non-functional — reliability (spec §10.2)

| Req | Summary | Module(s) | Test(s) | Status |
| --- | --- | --- | --- | --- |
| NFR-REL-1 | Per-request timeout (AbortSignal, 30s) | `graph/client` | `graphClient.test` | ✅ |
| NFR-REL-2 | Bounded jittered backoff retry | `graph/retry` | `graphRetry.test` | ✅ |
| NFR-REL-3 | No duplicate side effects (send/draft policy) | `graph/retry` (policy), `capabilities/sendMessage|createDraft` (apply) | `graphRetry.test` | 🟡 (policy done + tested; applied in phase 3) |
| NFR-REL-4 | Concurrency bound on bulk fetches | `util/bounded` | `bounded.test` | ◻ (deferred to phase-4 organise fan-out; read path needs none) |

## 8. Non-functional — performance & bounds (spec §10.3)

| Req | Summary | Module(s) | Test(s) | Status |
| --- | --- | --- | --- | --- |
| NFR-PERF-1 | Response char budget (ref 25,000) | `output/contract` (`clampText`) | `outputContract.test` | ✅ |
| NFR-PERF-2 | Conversation cap: 100 msgs / 20,000 chars | `output/contract`, `capabilities/readConversation` | `readConversation.test` | ✅ |
| NFR-PERF-3 | Validate outgoing size locally | `mail/compose` | `compose.test` | ◻ |

## 9. Non-functional — compatibility & operability (spec §10.5)

| Req | Summary | Module(s) | Test(s) | Status |
| --- | --- | --- | --- | --- |
| NFR-OPS-1 | Node ≥ 18 on Windows + macOS | `package.json` (engines), CI matrix | CI (`.github/workflows/ci.yml`) | ✅ |
| NFR-OPS-2 | stdio transport; report accounts to stderr | `index` | manual smoke | ✅ |
| NFR-OPS-3 | All knobs via env vars | `config` | `config.test` | ✅ |
| NFR-OPS-4 | MCP behavioural annotations; send/organise destructive | `index` (tool registration) | `tools.annotations.test` | 🟡 (C1/C2/C3 annotated read-only; destructive tools phase 3) |

## 10. Constraints & assumptions (spec §14)

| Id | Summary | Where addressed | Status |
| --- | --- | --- | --- |
| CON-1 | One server = one provider (Outlook) | whole build; `package.json` name | ✅ |
| CON-2 | Host acts on destructive annotations | `index` declares only | 🟡 (read tools declared) |
| CON-3 | Unverified-app consent policy in onboarding docs | build phase 5 docs | ◻ |
| CON-4 | Large binaries via allow-listed path, not base64 | `mail/attachments`, `config` | 🟡 (config done) |
| ASM-1 | Operator can register Entra app + run CLI | onboarding docs | ◻ |
| ASM-2 | Egress to Graph/Entra available | runtime env | n/a |

---

## 11. Acceptance-criteria coverage (spec §13)

| §13 criterion | Proven by | Live (operator) |
| --- | --- | --- |
| 1. Multi-account selection | `accountRegistry.test` ✅ | — |
| 2. Capabilities C1–C8 | C1–C3: `listAccounts`/`searchConversations`/`readConversation` tests ✅; C4–C8 ◻ | ✓ real sandbox |
| 3. Onboarding (PKCE+state, refresh, re-consent) | `msalClient.test`, `tokenStore.test` (mocked) 🟡 | ✓ real consent |
| 4. No duplicate sends | `graphRetry.test` (policy ✅); send capability ◻ | — |
| 5. Attachment guard | `attachments.test` ◻ | — |
| 6. Safety annotations | C1–C3 annotated; destructive tools ◻ | ✓ host prompt |
| 7. Resilience (timeout + retry) | `graphClient.test`, `graphRetry.test` ✅ | — |
| 8. Local-only secrets (600/700) | `tokenStore.test` ✅ | — |
| 9. Bounded output | `outputContract.test`, `readConversation.test` ✅ | — |

> Criteria 2 and 3 have an offline (mocked) portion proven here and a live portion requiring a real
> Entra app registration + Outlook mailboxes, run locally by the operator. **Live search/read
> validation** (criterion 2 for C2/C3) will confirm the Graph `$search`/`$filter` translations and
> the `$count`/`ConsistencyLevel` behaviour against a real mailbox.
