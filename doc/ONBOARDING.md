# Onboarding — register an Entra app and connect a mailbox

> Operator guide for getting this server talking to a real Outlook / Microsoft 365 mailbox
> (spec §8, ASM-1). It covers the one-time **Entra app registration**, the **consent policy**
> for an unverified app (CON-3), and the **CLI** that connects accounts. The MCP server itself
> never starts an interactive sign-in — you run the CLI once per mailbox, out of band.

Companion docs: [`README.md`](../README.md) (quick start + config), [`architecture.md`](./architecture.md)
§5 (auth design), [`provider-mapping.md`](./provider-mapping.md) §4 (identity mapping).

---

## 0. Prerequisites

- **Node.js ≥ 18** to run the published server (≥ 20 to run the test suite).
- An **Outlook.com** or **Microsoft 365** mailbox you can sign into in a browser.
- Permission to **register an application** in Entra ID. Personal Microsoft accounts can always
  self-register; for a work/school tenant you may need an administrator (see §4).

---

## 1. Register a public-client app in Entra ID

1. Go to the [Entra admin center](https://entra.microsoft.com) → **Identity → Applications → App
   registrations → New registration**.
2. **Name:** anything (e.g. `outlook-mcp`).
3. **Supported account types:** pick the audience you need — this must match the `tenant` you put
   in the config file (§3):

   | Audience | `tenant` value |
   | --- | --- |
   | Any org + personal Microsoft accounts | `common` |
   | Work/school accounts in any org | `organizations` |
   | Personal Microsoft accounts only | `consumers` |
   | One specific tenant | the tenant (directory) GUID |

4. **Redirect URI:** add a **Mobile and desktop applications** platform with the URI
   **`http://localhost`**. MSAL drives a loopback redirect during consent (FR-AUTH-2); the exact
   port is chosen at runtime, so `http://localhost` (no port) is the correct entry.
5. Register, then on **Overview** copy the **Application (client) ID** — you need it in §3.
6. **Authentication →** confirm **Allow public client flows** is **Yes** (this is a public client;
   there is no client secret — auth uses PKCE, FR-AUTH-3).

### Delegated permissions (least privilege — FR-AUTH-10)

Under **API permissions**, add these **delegated** Microsoft Graph permissions:

- `Mail.ReadWrite` — read, draft, and organise mail
- `Mail.Send` — send messages
- `User.Read` — identify the signed-in mailbox (`GET /me`)
- `offline_access` — obtain a refresh token so the server keeps working without re-consent

These are exactly the scopes the server requests; they are **fixed in code**, not read from your
config file, so the consent screen always asks for the least-privilege set.

---

## 2. Where files live (the data dir)

Everything is kept under the **data dir**, default `~/.outlook-mcp` (override with
`OUTLOOK_MCP_DATA_DIR`). The CLI creates it **owner-only** (`700`) and writes the token store
`tokens.json` **owner-only** (`600`); tokens never leave your machine (NFR-SEC-1). See
[`.env.example`](../.env.example) for all knobs.

---

## 3. Add a credential config

Drop a `credentials*.json` file into the data dir (or point `OUTLOOK_OAUTH_CREDENTIALS` at one
specific file). Minimum contents:

```json
{
  "clientId": "00000000-0000-0000-0000-000000000000",
  "tenant": "common"
}
```

- `clientId` *(required)* — the Application (client) ID from §1.5.
- `tenant` *(optional, default `common`)* — must match the audience chosen in §1.3.
- `id` *(optional)* — a stable label recorded against each account it authorises; defaults to the
  filename without `.json`. Use `--source <id>` to pick it at connect time.

**Multiple app registrations.** Any file matching `credentials*.json` is auto-discovered, so you
can run accounts across different orgs/registrations side by side (e.g.
`credentials-acme.json`, `credentials-personal.json`). Each account is bound to the registration
that authorised it and always refreshes with that same client (FR-ID-5/6). Setting
`OUTLOOK_OAUTH_CREDENTIALS` pins exactly one file and disables discovery.

---

## 4. Consent — including the "unverified app" prompt (CON-3)

Because this is **your** app registration (not a Microsoft-verified publisher), the browser consent
screen will show an **"unverified"** / **"This app is not published by Microsoft"** notice. That is
expected. How to get past it depends on the account type:

- **Personal Microsoft account (Outlook.com).** You can review the requested permissions and click
  **Accept** to consent for yourself. No admin involved.
- **Work / school account.** Your tenant may restrict user consent to unverified apps. Then either:
  - an **administrator grants consent** — Entra admin center → **Enterprise applications →** your
    app → **Permissions → Grant admin consent**, or visit
    `https://login.microsoftonline.com/{tenant}/adminconsent?client_id={clientId}` while signed in
    as an admin; **or**
  - the admin **enables user consent** for the requested delegated permissions; **or**
  - you complete **[publisher verification](https://learn.microsoft.com/azure/active-directory/develop/publisher-verification-overview)**
    to remove the "unverified" banner (optional, org-policy dependent).

Consent is hardened end-to-end: the loopback callback binds to loopback only, is short-lived, and
enforces PKCE + a CSRF `state` (NFR-SEC-7); forged callbacks are answered neutrally without
aborting the genuine flow (FR-AUTH-4).

---

## 5. Connect, list, and remove mailboxes (the CLI)

After `npm run build` (or `npx outlook-mcp-auth …` from the published package):

```bash
# Connect a mailbox — opens the browser for consent, then stores the bound token cache.
outlook-mcp-auth connect

# If several credentials*.json exist, choose which app registration to use:
outlook-mcp-auth connect --source acme
outlook-mcp-auth connect --credentials /path/to/credentials-acme.json

# List the connected mailboxes (identity + which app registration authorised each).
outlook-mcp-auth list

# Remove a mailbox (deletes its stored token cache).
outlook-mcp-auth remove user@example.com
```

`connect` prints the connected identity on success; errors are written to stderr with any
token/credential material **redacted** (NFR-SEC-6). Connecting an already-connected account simply
repairs/refreshes its stored cache.

---

## 6. Point your MCP host at the server

The server speaks **stdio** and is launched by your MCP host (NFR-OPS-2). A typical host config:

```json
{
  "mcpServers": {
    "outlook": {
      "command": "outlook-mcp",
      "env": { "OUTLOOK_MCP_DATA_DIR": "/home/you/.outlook-mcp" }
    }
  }
}
```

On startup the server prints the connected accounts to **stderr** (never stdout, which is the
JSON-RPC channel; never secrets). Tools then accept an optional `account` selector — omit it when
exactly one mailbox is connected (FR-ID-2). Use `list_accounts` to see the valid identities.

---

## 7. Troubleshooting

| Symptom | Fix |
| --- | --- |
| `No accounts are connected` | Run `outlook-mcp-auth connect`. |
| `AADSTS65001` / consent required | Grant consent for the account type (see §4); for work accounts an admin may need to consent. |
| `AADSTS50011` redirect mismatch | The app registration is missing the **`http://localhost`** redirect under *Mobile and desktop applications* (§1.4). |
| `Authentication failed … Re-connect this account` | The refresh token was revoked/expired; re-run `outlook-mcp-auth connect` for that mailbox (the server picks up the new cache without a restart — FR-AUTH-9). |
| `Credential source "…" is missing` | The `credentials*.json` that authorised the account was moved/removed; restore it or re-connect. |
| Reading a `path` attachment is refused | Path attachments are disabled unless `OUTLOOK_MCP_ATTACHMENTS_DIR` is set to an allow-listed directory (NFR-SEC-3); inline base64 always works. |

---

*See [`business-specification.md`](./business-specification.md) §8 for the normative auth
requirements and [`architecture.md`](./architecture.md) §5 / §9 for the consent and security design.*
