# MCP Server for Multiple Outlook Accounts

A **local, single-connector** [Model Context Protocol](https://modelcontextprotocol.io) server
that lets an AI assistant operate **several Outlook / Microsoft 365 mailboxes at once** —
searching, reading, drafting, sending, and organising mail — through one safety-annotated tool
surface, with OAuth tokens that **never leave your machine**.

This is the **Outlook / Microsoft Graph** provider variant of a provider-neutral specification.

- 📄 [`doc/business-specification.md`](./doc/business-specification.md) — provider-neutral business + functional spec (the contract).
- 🔌 [`doc/provider-mapping.md`](./doc/provider-mapping.md) — how neutral requirements bind to Microsoft Graph.
- 🏗️ [`doc/architecture.md`](./doc/architecture.md) — how this variant is designed and built.
- ✅ [`doc/traceability-matrix.md`](./doc/traceability-matrix.md) — every requirement → module → test.
- 🚀 [`doc/ONBOARDING.md`](./doc/ONBOARDING.md) — operator guide: register an Entra app, handle consent, connect a mailbox.
- 🤝 [`doc/CONTRIBUTING.md`](./doc/CONTRIBUTING.md) — handoff guide: how we work, workflow, and the per-phase record.
- 📋 [`doc/HANDOVER.md`](./doc/HANDOVER.md) — handover snapshot: what's done and the outstanding-work checklist.

---

## Project status: phase 5 (hardening + onboarding)

Built with **TypeScript 6.0.3**; build, typecheck, tests (148), and format all green. All eight
capabilities are implemented and the offline build is feature-complete. What exists today:

- ✅ Architecture design + requirements traceability matrix (`doc/`).
- ✅ **Auth core:** Entra credential-source discovery, MSAL public-client (consent + silent refresh),
  secure token store (`0600`/`0700`, atomic + cross-process-locked), account-selection registry.
- ✅ **Account-management CLI:** `outlook-mcp-auth connect | list | remove`.
- ✅ **Microsoft Graph client:** thin `fetch` wrapper with a per-request timeout, bounded jittered
  retry (with the no-duplicate-send policy), and actionable error mapping.
- ✅ **Read tools:** `list_accounts` (C1), `search_conversations` (C2), `read_conversation` (C3) — with
  Gmail-style search-operator translation and bounded, truncating output.
- ✅ **Write tools:** `create_draft` (C4) and `send_message` (C5) — recipient parsing
  (`Display Name <addr>`), header-injection stripping, allow-listed/TOCTOU-safe attachments,
  local outgoing-size validation, reply threading, and a single `sendMail` call under the
  `nonDuplicable` retry policy so a retry can never double-deliver.
- ✅ **Organise tools:** `list_labels` (C6), `create_label` (C7), and `organize_mail` (C8) — the
  label-decomposition fan-out that maps one neutral organise request to the right mix of Graph
  category-PATCH / `move` / read-state calls, applied per message across a conversation under a
  bounded concurrency limit, reporting the union of resulting labels.
- ✅ **Hardening + onboarding:** a secret-redaction boundary so tokens/credentials can never reach
  the logs (NFR-SEC-6), and the operator [onboarding guide](./doc/ONBOARDING.md) (Entra app
  registration, unverified-app consent, the connect CLI).

**Remaining:** the **live** acceptance runs (real browser consent + real Graph calls) that the
operator performs against an Entra app registration + Outlook mailbox. See `doc/architecture.md` §13.

> Tests mock Microsoft Graph and MSAL. Live `§13` acceptance — real browser consent and Graph calls —
> requires an Entra app registration + Outlook mailboxes and is run locally by the operator.

---

## Capabilities (spec §5)

| Tool                   | Purpose                    | Destructive? | Status  |
| ---------------------- | -------------------------- | ------------ | ------- |
| `list_accounts`        | List connected mailboxes   | No           | ✅ live |
| `search_conversations` | Search a mailbox (paged)   | No           | ✅ live |
| `read_conversation`    | Read a full conversation   | No           | ✅ live |
| `create_draft`         | Compose a draft (not sent) | No           | ✅ live |
| `send_message`         | Send immediately           | **Yes**      | ✅ live |
| `list_labels`          | List categories + folders  | No           | ✅ live |
| `create_label`         | Create a category/folder   | No           | ✅ live |
| `organize_mail`        | Tag / move / read-state    | **Yes**      | ✅ live |

Plus an out-of-band account-management CLI (`outlook-mcp-auth connect | list | remove`).

---

## Development

Requires **Node.js ≥ 18** (developed on Node 22).

```bash
npm install
npm run typecheck      # tsc --noEmit
npm run build          # compile to dist/
npm test               # vitest (Graph/MSAL mocked)
npm run format:check   # prettier
```

### Configuration

All operational knobs are environment variables (see [`.env.example`](./.env.example)):

| Env var                          | Meaning                                       | Default          |
| -------------------------------- | --------------------------------------------- | ---------------- |
| `OUTLOOK_MCP_DATA_DIR`           | tokens + app-registration configs             | `~/.outlook-mcp` |
| `OUTLOOK_OAUTH_CREDENTIALS`      | pin one app registration (disables discovery) | unset            |
| `OUTLOOK_MCP_ATTACHMENTS_DIR`    | allow-list for `path` attachments             | unset (disabled) |
| `OUTLOOK_MCP_LOCK_TIMEOUT_MS`    | token-store lock wait                         | `12000`          |
| `OUTLOOK_MCP_REQUEST_TIMEOUT_MS` | per-Graph-call timeout                        | `30000`          |

### Connecting a mailbox

> For the full walkthrough — Entra app registration, the unverified-app consent policy, and
> troubleshooting — see **[`doc/ONBOARDING.md`](./doc/ONBOARDING.md)**. The short version:

1. Register a **public-client** app in Entra ID with the redirect URI `http://localhost` and the
   delegated scopes `Mail.ReadWrite`, `Mail.Send`, `User.Read`, `offline_access`.
2. Drop a `credentials*.json` into the data dir (or point `OUTLOOK_OAUTH_CREDENTIALS` at it):

   ```json
   { "clientId": "<application-client-id>", "tenant": "common" }
   ```

   Multiple `credentials*.json` files are auto-discovered, so accounts under different app
   registrations each refresh with the client that authorised them.

3. Connect, list, and remove mailboxes:

   ```bash
   outlook-mcp-auth connect            # opens the browser for consent
   outlook-mcp-auth connect --source acme   # pick a specific app registration
   outlook-mcp-auth list
   outlook-mcp-auth remove user@example.com
   ```

### Security posture

OAuth tokens are stored only on the local machine (file mode `600` in a `700` data dir). Reading
local files by path for attachments is **disabled by default** and only allowed from an explicit
allow-list. The server never logs tokens, credentials, or message content. See
`doc/architecture.md` §9.

## License

MIT — see [LICENSE](./LICENSE).
