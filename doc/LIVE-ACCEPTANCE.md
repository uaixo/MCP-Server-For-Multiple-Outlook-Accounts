# Live acceptance runbook

> The offline build is feature-complete and green, but the spec's §13.2/§13.3 acceptance must be
> confirmed against a **real** Entra app registration + Outlook mailbox — real consent, real
> Microsoft Graph calls. That can only run on an operator's machine with real credentials, so it
> isn't part of CI. This runbook turns the HANDOVER §4.1 checklist into ordered, concrete steps with
> pass/fail criteria. Work top to bottom; flip the matching row in
> [`traceability-matrix.md`](./traceability-matrix.md) §11 as each passes.

**Safety:** §2–§3 send real email and modify real mail. Use a **test mailbox** (or send only to
yourself), and prefer a throwaway conversation for the organise checks.

---

## 0. Prerequisites

1. Register the Entra app + create `credentials*.json` and connect a mailbox — follow
   [`ONBOARDING.md`](./ONBOARDING.md) §1–§5.
2. Build the server:
   ```bash
   npm install && npm run build
   ```
3. Confirm at least one account is connected:
   ```bash
   outlook-mcp-auth list
   ```

---

## 1. Read-only smoke (automated)

Runs the non-destructive tools end-to-end against your real mailbox and prints PASS/FAIL:

```bash
npm run live-smoke
# options: --account you@example.com   --query "is:unread"
```

It checks: the 8-tool surface, `list_accounts`, `search_conversations`, `read_conversation` (first
hit), and `list_labels`. Exit code is non-zero on any failure. This covers the read half of
acceptance criterion §13.2 for **C1/C2/C3** and **C6**.

| Confirms | Criterion |
| --- | --- |
| Accounts enumerate; tool surface complete | §13.1, §13.6 (read) |
| `$search`/`$filter` translation returns real results | §13.2 (C2), watch-item §7.2 |
| Conversation read renders headers + HTML→text, bounded | §13.2 (C3) |
| Categories **and** nested folders (full paths) list | §13.2 (C6) |

> The write/destructive tools (`create_draft`, `send_message`, `create_label`, `organize_mail`) are
> **not** automated here — run them per §3 via your MCP host (or extend `scripts/live-smoke.mjs`).

---

## 2. Onboarding & auth (manual)

| # | Action | Expected / pass criterion | Req |
| --- | --- | --- | --- |
| 2.1 | `outlook-mcp-auth connect` | Browser opens; consent (incl. the unverified-app prompt for your account type) completes; the loopback callback returns; terminal prints `Connected <you>`. | FR-AUTH-1..7, NFR-SEC-7 |
| 2.2 | Forge a junk callback hit during 2.1 (optional) | The genuine flow still completes; the forged hit is answered neutrally. | FR-AUTH-4 |
| 2.3 | `outlook-mcp-auth list` then `remove <acct>` then re-`connect` | Account shows with its source; remove deletes it; re-connect restores it. | FR-AUTH-8 |
| 2.4 | Leave the server running; `connect` the same account again in another shell; call any tool | The running server uses the refreshed cache **without a restart**. | FR-AUTH-9 |
| 2.5 | After ~1 h idle, call any tool | Silent token refresh succeeds (no re-consent). | FR-AUTH-5, FR-ID-5 |

---

## 3. Write & organise (manual — sends/modifies real mail)

Invoke each tool through your MCP host (e.g. Claude Desktop pointed at `outlook-mcp`) with the
arguments shown, or extend the smoke script. Verify the effect in Outlook/OWA.

### 3.1 Drafts & send (C4/C5)

| # | Tool + arguments | Expected / pass criterion | Req |
| --- | --- | --- | --- |
| 3.1a | `create_draft` `{ to:["you@…"], subject:"acceptance", body:"hello" }` | A draft appears in Drafts; nothing is sent; result returns a `draft_id`. | FR-C4-1/2 |
| 3.1b | `create_draft` with `reply_to_conversation_id` of an existing thread, subject omitted | Draft files into that thread; subject defaults to `Re: <original>`. **Open the sent/received reply and confirm it threads** (the `In-Reply-To`/`References` headers are honoured — see HANDOVER §4.2). | FR-C4-4 |
| 3.1c | `create_draft` with an inline base64 attachment ≤ 3 MB | Draft has the attachment (inline `fileAttachment`). | FR-C4-3 |
| 3.1d | `create_draft` (or `send_message`) with an attachment **> 3 MB** (allow-listed path or inline) | The attachment uploads via an **upload session** (chunked); the message carries the large file. Confirm it arrives intact. | FR-C4-3, NFR-PERF-3 |
| 3.1e | `send_message` `{ to:["you@…"], subject:"acceptance send", body:"hi" }` | Email is **delivered** (arrives in your inbox); appears in Sent Items. | FR-C5-1/2 |
| 3.1f | `send_message` with a **> 3 MB** attachment | create-draft → upload → send; delivered exactly once. | FR-C5, NFR-PERF-3 |
| 3.1g | No-duplicate check | Inducing a transient failure isn't easy live; instead confirm across 3.1e/3.1f that **exactly one** copy arrives and no orphaned duplicate is sent. (Offline forced-failure is covered by `sendMessage.test`.) | FR-C5-4, NFR-REL-3 |

### 3.2 Labels & organise (C7/C8)

| # | Tool + arguments | Expected / pass criterion | Req |
| --- | --- | --- | --- |
| 3.2a | `create_label` `{ name:"Acceptance", kind:"category" }` | Category appears in Outlook; `list_labels` now returns it. | FR-C7-1 |
| 3.2b | `create_label` `{ name:"Acme", kind:"folder", parent_folder_id:"<id>" }` | Nested folder created; `list_labels` shows it by full path (e.g. `Inbox/Acme`). | FR-C7-1, FR-C6-1 |
| 3.2c | `organize_mail` `{ message_id:"<id>", add_labels:["Acceptance"], mark_read:true }` | Message gains the category and is marked read (single PATCH). | FR-C8-1/2/6 |
| 3.2d | `organize_mail` `{ conversation_id:"<id>", archive:true }` | Every message in the thread moves to Archive; result reports the union. | FR-C8-4, §3.1 |
| 3.2e | `organize_mail` with `trash:true`, then another with `junk:true` (different messages) | Move to Deleted Items / Junk Email respectively; providing two of archive/trash/junk is rejected. | FR-C8-3 |
| 3.2f | Re-run 3.2c unchanged | Converges (idempotent); no error. | FR-C8 (idempotent) |

---

## 4. Host gating (CON-2)

| # | Action | Expected / pass criterion | Req |
| --- | --- | --- | --- |
| 4.1 | In your MCP host, invoke `send_message` and `organize_mail` | The host **prompts for confirmation** (they carry `destructiveHint: true`); read tools do not. | NFR-OPS-4, CON-2 |

---

## 5. Record results

For each row that passes, flip the corresponding criterion in
[`traceability-matrix.md`](./traceability-matrix.md) §11 from "✓ live" pending to confirmed, and note
the date + mailbox/tenant used. If anything fails, capture the tool's structured error (already
redacted) and the Outlook-side observation, and open an issue referencing the FR/§13 id.

> When §1–§4 all pass, the spec's acceptance criteria are satisfied end-to-end and the project is
> done — there is no remaining offline work.
