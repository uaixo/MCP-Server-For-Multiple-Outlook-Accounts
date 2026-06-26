#!/usr/bin/env node
/**
 * Read-only live smoke test for the Outlook MCP server (doc/LIVE-ACCEPTANCE.md).
 *
 * Drives the BUILT server (`dist/index.js`) over stdio with a real MCP client and
 * exercises the NON-DESTRUCTIVE tools against whatever mailbox you've connected
 * with `outlook-mcp-auth connect`: list_accounts → search_conversations →
 * read_conversation → list_labels. It never drafts, sends, creates, or organises
 * anything — those (write/destructive) checks are run by hand per the runbook.
 *
 * Prerequisites:
 *   npm run build                 # produces dist/
 *   outlook-mcp-auth connect      # at least one real account
 * Run:
 *   npm run live-smoke
 *   npm run live-smoke -- --account you@example.com --query "is:unread"
 *
 * Exit code is non-zero if any step fails, so it's CI-able on the operator's box.
 */

import { existsSync } from "node:fs";
import { fileURLToPath } from "node:url";
import { dirname, join } from "node:path";
import { Client } from "@modelcontextprotocol/sdk/client/index.js";
import { StdioClientTransport } from "@modelcontextprotocol/sdk/client/stdio.js";

const root = join(dirname(fileURLToPath(import.meta.url)), "..");
const serverEntry = join(root, "dist", "index.js");

function arg(name, fallback) {
  const i = process.argv.indexOf(`--${name}`);
  return i >= 0 && process.argv[i + 1] ? process.argv[i + 1] : fallback;
}
const account = arg("account", undefined); // omit → default-account rule
const query = arg("query", "is:unread");

const results = [];
const record = (step, ok, detail) => {
  results.push({ step, ok, detail });
  console.log(`${ok ? "PASS" : "FAIL"}  ${step}${detail ? ` — ${detail}` : ""}`);
};

function structured(res) {
  if (res.isError) {
    const text = res.content?.find((c) => c.type === "text")?.text ?? "tool error";
    throw new Error(text);
  }
  return res.structuredContent ?? {};
}

async function main() {
  if (!existsSync(serverEntry)) {
    console.error(`Build first: ${serverEntry} not found. Run "npm run build".`);
    process.exit(2);
  }

  // Pass through env (only string values) so OUTLOOK_MCP_* config reaches the server.
  const env = Object.fromEntries(
    Object.entries(process.env).filter(([, v]) => typeof v === "string"),
  );

  const transport = new StdioClientTransport({
    command: process.execPath,
    args: [serverEntry],
    env,
  });
  const client = new Client({ name: "live-smoke", version: "1.0.0" });
  await client.connect(transport);

  try {
    // 0) Tool surface.
    const { tools } = await client.listTools();
    const names = new Set(tools.map((t) => t.name));
    const expected = [
      "list_accounts",
      "search_conversations",
      "read_conversation",
      "create_draft",
      "send_message",
      "list_labels",
      "create_label",
      "organize_mail",
    ];
    const missing = expected.filter((n) => !names.has(n));
    record(
      "tool surface (8 tools registered)",
      missing.length === 0,
      missing.length ? `missing ${missing.join(", ")}` : `${names.size} tools`,
    );

    // 1) list_accounts.
    const accounts = structured(await client.callTool({ name: "list_accounts", arguments: {} }));
    const count = accounts.account_count ?? 0;
    record("list_accounts", count > 0, `${count} account(s)`);
    if (count === 0) {
      console.error('\nNo accounts connected. Run "outlook-mcp-auth connect" first.');
      process.exit(1);
    }

    const sel = account ? { account } : {};

    // 2) search_conversations.
    const search = structured(
      await client.callTool({ name: "search_conversations", arguments: { ...sel, query } }),
    );
    const conversations = search.conversations ?? [];
    record(
      "search_conversations",
      true,
      `query "${query}" → ${conversations.length} conversation(s)`,
    );

    // 3) read_conversation (first hit, if any).
    if (conversations[0]?.conversation_id) {
      const read = structured(
        await client.callTool({
          name: "read_conversation",
          arguments: { ...sel, conversation_id: conversations[0].conversation_id },
        }),
      );
      record(
        "read_conversation",
        (read.message_count ?? 0) > 0,
        `${read.message_count} message(s)${read.truncated ? " (truncated)" : ""}`,
      );
    } else {
      record("read_conversation", true, "skipped — no conversation matched the query");
    }

    // 4) list_labels.
    const labels = structured(await client.callTool({ name: "list_labels", arguments: sel }));
    const folders = (labels.labels ?? []).filter((l) => l.kind === "folder").length;
    const categories = (labels.labels ?? []).filter((l) => l.kind === "category").length;
    record(
      "list_labels",
      (labels.label_count ?? 0) > 0,
      `${categories} categor(y/ies), ${folders} folder(s)`,
    );
  } finally {
    await client.close();
  }

  const failed = results.filter((r) => !r.ok);
  console.log(`\n${results.length - failed.length}/${results.length} read-only checks passed.`);
  console.log("Now run the WRITE/DESTRUCTIVE checks by hand — see doc/LIVE-ACCEPTANCE.md §3.");
  process.exit(failed.length ? 1 : 0);
}

main().catch((err) => {
  console.error(`\nlive-smoke failed: ${err instanceof Error ? err.message : String(err)}`);
  process.exit(1);
});
