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

---

## ⚠️ Project status: scaffold + design

This repository currently contains the **project scaffold and design only**, for review before the
full capability build. What exists today:

- ✅ Buildable TypeScript 6.0.3 project (build, test, format all green).
- ✅ Neutral domain types and subsystem contracts (`src/domain/`).
- ✅ Tested configuration loader (`src/config.ts`).
- ✅ Server + CLI entry points (start and report status; **no capability tools yet**).
- ✅ Architecture design + requirements traceability matrix.

**Not yet implemented:** the eight capabilities (C1–C8), MSAL consent flow, secure token store, and
the Microsoft Graph client. These are designed in `doc/architecture.md` §13 and tracked as _Planned_
in the traceability matrix. They land in subsequent build phases once this design is approved.

---

## Planned capabilities (spec §5)

| Tool                   | Purpose                    | Destructive? |
| ---------------------- | -------------------------- | ------------ |
| `list_accounts`        | List connected mailboxes   | No           |
| `search_conversations` | Search a mailbox (paged)   | No           |
| `read_conversation`    | Read a full conversation   | No           |
| `create_draft`         | Compose a draft (not sent) | No           |
| `send_message`         | Send immediately           | **Yes**      |
| `list_labels`          | List categories + folders  | No           |
| `create_label`         | Create a category/folder   | No           |
| `organize_mail`        | Tag / move / read-state    | **Yes**      |

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

### Security posture

OAuth tokens are stored only on the local machine (file mode `600` in a `700` data dir). Reading
local files by path for attachments is **disabled by default** and only allowed from an explicit
allow-list. The server never logs tokens, credentials, or message content. See
`doc/architecture.md` §9.

## License

MIT — see [LICENSE](./LICENSE).
