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

**Current phase:** **Phase 1 complete** — auth core (credential sources, MSAL client, secure
token store + cross-process lock, account registry), the `connect`/`list`/`remove` CLI, and **C1
`list_accounts`**. 29 tests pass (Graph/MSAL mocked). Read/write/organise capabilities (C2–C8) and
the Graph client remain ◻ Planned (phases 2–4 in architecture.md §13).

> **Delegated-to-MSAL note:** the loopback redirect, PKCE (S256), and CSRF `state`
> (FR-AUTH-2/3/4, NFR-SEC-7) are satisfied by MSAL's `acquireTokenInteractive`, the chosen
> library (provider-mapping §4). They are marked ✅ "(MSAL)" — we rely on, rather than
> re-implement, that hardening.

---

## 1. Functional requirements — capabilities (spec §6)

| Req | Summary | Module(s) | Test(s) | Status |
| --- | --- | --- | --- | --- |
| FR-C1-1 | List connected accounts as identities | `auth/tokenStore`, `capabilities/listAccounts` | `listAccounts.test`, `tokenStore.test` | ✅ |
| FR-C1-2 | Empty → non-error connect guidance | `capabilities/listAccounts` | `listAccounts.test` | ✅ |
| FR-C1-3 | Output feeds other tools' selectors | `capabilities/listAccounts` | `listAccounts.test` | ✅ |
| FR-C2-1 | Required query in native syntax | `search/translate`, `capabilities/searchConversations` | `translate.test` | ◻ |
| FR-C2-2 | Page size default 20, max 100 | `capabilities/searchConversations` | `searchConversations.test` | ◻ |
| FR-C2-3 | Opaque next-page cursor (`@odata.nextLink`) | `capabilities/searchConversations` | `searchConversations.test` | ◻ |
| FR-C2-4 | Summary: id, subject, sender, date, snippet | `capabilities/searchConversations` | `searchConversations.test` | ◻ |
| FR-C2-5 | Degrade a failed entry, not whole search | `capabilities/searchConversations` | `searchConversations.test` | ◻ |
| FR-C2-6 | Bound per-conversation fan-out | `util/bounded`, `capabilities/searchConversations` | `bounded.test` | ◻ |
| FR-C3-1 | Full conversation: headers, body, labels | `capabilities/readConversation` | `readConversation.test` | ◻ |
| FR-C3-2 | Cap messages (newest) + body chars; `truncated` | `output/contract`, `capabilities/readConversation` | `readConversation.test` | ◻ |
| FR-C3-3 | HTML → readable plain text | `capabilities/readConversation` | `readConversation.test` | ◻ |
| FR-C4-1 | Compose to/cc/bcc/subject/body/is_html | `mail/compose`, `capabilities/createDraft` | `compose.test` | ◻ |
| FR-C4-2 | Persist draft, do not send | `capabilities/createDraft` | `createDraft.test` | ◻ |
| FR-C4-3 | Attachment: path XOR inline base64 | `mail/attachments` | `attachments.test` | ◻ |
| FR-C4-4 | Reply drafting + `Re:` default subject | `mail/compose`, `capabilities/createDraft` | `compose.test` | ◻ |
| FR-C4-5 | Recipient forms + injection prevention | `mail/compose`, `mail/sanitize` | `sanitize.test` | ◻ |
| FR-C5-1 | Send: same inputs + reply-to ref | `capabilities/sendMessage` | `sendMessage.test` | ◻ |
| FR-C5-2 | Deliver immediately (irreversible) | `capabilities/sendMessage` | `sendMessage.test` | ◻ |
| FR-C5-3 | Annotated destructive | `index` (tool registration) | `tools.annotations.test` | ◻ |
| FR-C5-4 | No duplicate delivery under retry | `graph/retry`, `capabilities/sendMessage` | `retry.test` (forced transient) | ◻ |
| FR-C6-1 | List labels: id, name, type | `capabilities/listLabels` | `listLabels.test` | ◻ |
| FR-C6-2 | Output is id source for C8 | `capabilities/listLabels` | `listLabels.test` | ◻ |
| FR-C7-1 | Create user label (+nesting via folders) | `capabilities/createLabel` | `createLabel.test` | ◻ |
| FR-C8-1 | Exactly one target (conversation XOR message) | `capabilities/organizeMail` | `organizeMail.test` | ◻ |
| FR-C8-2 | Add and/or remove; ≥1 change required | `organise/decompose`, `capabilities/organizeMail` | `decompose.test` | ◻ |
| FR-C8-3 | Derived intents: read/unread, archive | `organise/decompose` | `decompose.test` | ◻ |
| FR-C8-4 | Conversation → report union of labels | `capabilities/organizeMail` | `organizeMail.test` | ◻ |
| FR-C8-5 | Annotated destructive | `index` (tool registration) | `tools.annotations.test` | ◻ |
| FR-C8-6 | Decompose into correct Graph op combination | `organise/decompose` | `decompose.test` | ◻ |

## 2. Account & identity model (spec §7)

| Req | Summary | Module(s) | Test(s) | Status |
| --- | --- | --- | --- | --- |
| FR-ID-1 | Optional account selector on every tool but C1 | `domain/types` (`AccountSelectable`), `auth/accountRegistry` | `accountRegistry.test` | 🟡 (resolver done; selectors wired per-tool from C2) |
| FR-ID-2 | Default rule (1 → use; ≥2 → disambiguate; 0 → guide) | `auth/accountRegistry` | `accountRegistry.test` | ✅ |
| FR-ID-3 | Unknown account → error listing accounts | `auth/accountRegistry` | `accountRegistry.test` | ✅ |
| FR-ID-4 | Case-insensitive identity (lower-cased key) | `auth/accountRegistry`, `auth/tokenStore` | `accountRegistry.test`, `tokenStore.test` | ✅ |
| FR-ID-5 | Refresh uses the issuing app registration | `auth/msalClient`, `auth/tokenStore`, `auth/credentialSources` | `msalClient.test` (acquireToken), `tokenStore.test` | ✅ |
| FR-ID-6 | Multiple OAuth clients, auto-discovered | `auth/credentialSources` | `credentialSources.test` | ✅ |

## 3. Authentication & onboarding (spec §8)

| Req | Summary | Module(s) | Test(s) | Status |
| --- | --- | --- | --- | --- |
| FR-AUTH-1 | CLI connect via auth-code browser consent | `cli/connect`, `auth/msalClient` | `msalClient.test` (core flow) | ✅ |
| FR-AUTH-2 | Loopback redirect on 127.0.0.1 (+fallback) | `auth/msalClient` (MSAL) | — | ✅ (MSAL) |
| FR-AUTH-3 | PKCE (S256) | `auth/msalClient` (MSAL) | — | ✅ (MSAL) |
| FR-AUTH-4 | CSRF `state` round-trip; neutral on forged | `auth/msalClient` (MSAL) | — | ✅ (MSAL) |
| FR-AUTH-5 | Offline access → refresh token | `auth/msalClient`, `auth/credentialSources` | `msalClient.test` | ✅ |
| FR-AUTH-6 | Identify account (authenticated identity) | `cli/connect`, `auth/msalClient` | `msalClient.test` (abort-if-undetermined) | ✅ |
| FR-AUTH-7 | Bounded consent wait (ref 5 min) | `auth/msalClient` (`withTimeout`) | `msalClient.test` | ✅ |
| FR-AUTH-8 | CLI list (w/ source) + remove | `cli/list`, `cli/remove`, `auth/tokenStore` | `tokenStore.test` | ✅ |
| FR-AUTH-9 | Re-consent repair, no restart needed | `auth/tokenStore` (per-access read + atomic overwrite) | `tokenStore.test` | ✅ |
| FR-AUTH-10 | Least-privilege scopes | `auth/credentialSources` (`OUTLOOK_SCOPES`) | `credentialSources.test` | ✅ |

> FR-AUTH-6 identity = MSAL `account.username` (UPN), lower-cased as the store key. A `GET /me`
> enrichment to prefer primary SMTP (`mail`) is a phase-2 refinement once the Graph client lands;
> the key is stable (UPN == primary SMTP for typical M365 accounts), so no re-keying is required.

## 4. Error handling & resilience (spec §9)

| Req | Summary | Module(s) | Test(s) | Status |
| --- | --- | --- | --- | --- |
| FR-ERR-1 | Provider errors → actionable messages | `graph/errors` (Graph); `auth/accountRegistry` (selection) | `accountRegistry.test` | 🟡 (selection errors done; Graph mapping phase 2) |
| FR-ERR-2 | Corrupt token store → "no accounts" + warning | `auth/tokenStore` | `tokenStore.test` | ✅ |
| FR-ERR-3 | Validation errors before any provider call | `capabilities/*` (zod), `mail/*`, `organise/decompose` | per-capability tests | ◻ |

## 5. Output / response contract (spec §11)

| Req | Summary | Module(s) | Test(s) | Status |
| --- | --- | --- | --- | --- |
| FR-OUT-1 | Dual channel: human text + structured | `domain/types` (`ToolResult`), `capabilities/listAccounts`, `index` | `listAccounts.test` | 🟡 (C1 done; `output/contract` module pending) |
| FR-OUT-2 | Structured authoritative when text truncated | `output/contract` | `contract.test` | ◻ |
| FR-OUT-3 | Stable, documented field names | `capabilities/listAccounts`, `output/contract` | `listAccounts.test` | 🟡 (C1 fields stable; central module pending) |

## 6. Non-functional — security & privacy (spec §10.1)

| Req | Summary | Module(s) | Test(s) | Status |
| --- | --- | --- | --- | --- |
| NFR-SEC-1 | Tokens local only; 600 file in 700 dir | `auth/tokenStore` | `tokenStore.test` (stat modes) | ✅ |
| NFR-SEC-2 | Atomic temp+rename, cross-process lock | `auth/tokenStore`, `util/lock` | `lock.test`, `tokenStore.test` | ✅ |
| NFR-SEC-3 | Attachment path guard (allow-list, resolve) | `mail/attachments`, `config` | `attachments.test` | 🟡 (config done; reader phase 3) |
| NFR-SEC-4 | TOCTOU-safe: open once, validate via handle | `mail/attachments` | `attachments.test` | ◻ |
| NFR-SEC-5 | Strip CR/LF from header-bound values | `mail/sanitize` | `sanitize.test` | ◻ |
| NFR-SEC-6 | Never log tokens/credentials/content | all modules (logging policy) | `logging.test` (assertion) | 🟡 (policy followed; assertion test planned) |
| NFR-SEC-7 | Consent loopback hardening (PKCE+state) | `auth/msalClient` (MSAL) | — | ✅ (MSAL) |

## 7. Non-functional — reliability (spec §10.2)

| Req | Summary | Module(s) | Test(s) | Status |
| --- | --- | --- | --- | --- |
| NFR-REL-1 | Per-request timeout (AbortSignal, 30s) | `graph/client` | `client.test` (stalled request) | ◻ |
| NFR-REL-2 | Bounded jittered backoff retry | `graph/retry` | `retry.test` | ◻ |
| NFR-REL-3 | No duplicate side effects (send/draft policy) | `graph/retry`, `capabilities/sendMessage`, `capabilities/createDraft` | `retry.test` | ◻ |
| NFR-REL-4 | Concurrency bound on bulk fetches | `util/bounded` | `bounded.test` | ◻ |

## 8. Non-functional — performance & bounds (spec §10.3)

| Req | Summary | Module(s) | Test(s) | Status |
| --- | --- | --- | --- | --- |
| NFR-PERF-1 | Response char budget (ref 25,000) | `output/contract` | `contract.test` | ◻ |
| NFR-PERF-2 | Conversation cap: 100 msgs / 20,000 chars | `output/contract`, `capabilities/readConversation` | `readConversation.test` | ◻ |
| NFR-PERF-3 | Validate outgoing size locally | `mail/compose` | `compose.test` | ◻ |

## 9. Non-functional — compatibility & operability (spec §10.5)

| Req | Summary | Module(s) | Test(s) | Status |
| --- | --- | --- | --- | --- |
| NFR-OPS-1 | Node ≥ 18 on Windows + macOS | `package.json` (engines), `config` (path sep injectable), `auth/msalClient` (cross-platform open) | `config.test` | ✅ |
| NFR-OPS-2 | stdio transport; report accounts to stderr | `index` | manual smoke (CLI/list verified) | ✅ |
| NFR-OPS-3 | All knobs via env vars | `config` | `config.test` | ✅ |
| NFR-OPS-4 | MCP behavioural annotations; send/organise destructive | `index` (tool registration) | `tools.annotations.test` | 🟡 (C1 annotated; destructive tools from phase 3) |

## 10. Constraints & assumptions (spec §14)

| Id | Summary | Where addressed | Status |
| --- | --- | --- | --- |
| CON-1 | One server = one provider (Outlook) | whole build; `package.json` name | ✅ |
| CON-2 | Host acts on destructive annotations | `index` declares only | 🟡 (C1 declared) |
| CON-3 | Unverified-app consent policy in onboarding docs | build phase 5 docs | ◻ |
| CON-4 | Large binaries via allow-listed path, not base64 | `mail/attachments`, `config` | 🟡 (config done) |
| ASM-1 | Operator can register Entra app + run CLI | onboarding docs | ◻ |
| ASM-2 | Egress to Graph/Entra available | runtime env | n/a |

---

## 11. Acceptance-criteria coverage (spec §13)

| §13 criterion | Proven by | Live (operator) |
| --- | --- | --- |
| 1. Multi-account selection | `accountRegistry.test` ✅ | — |
| 2. Capabilities C1–C8 | C1: `listAccounts.test` ✅; C2–C8 ◻ | ✓ real sandbox |
| 3. Onboarding (PKCE+state, refresh, re-consent) | `msalClient.test`, `tokenStore.test` (mocked MSAL) 🟡 | ✓ real consent |
| 4. No duplicate sends | `retry.test` (forced transient) ◻ | — |
| 5. Attachment guard | `attachments.test` ◻ | — |
| 6. Safety annotations | C1 annotated; `tools.annotations.test` ◻ | ✓ host prompt |
| 7. Resilience (timeout + retry) | `client.test`, `retry.test` ◻ | — |
| 8. Local-only secrets (600/700) | `tokenStore.test` (stat) ✅ | — |
| 9. Bounded output | `contract.test`, `readConversation.test` ◻ | — |

> Criteria 2 and 3 have an offline (mocked) portion and a live portion that requires a real Entra
> app registration + Outlook mailboxes, run locally by the operator (this environment cannot hold
> live credentials). The interactive consent flow itself (FR-AUTH-2/3/4) is delegated to MSAL and
> verified live by the operator.
