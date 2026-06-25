import { defineConfig } from "vitest/config";

// Test strategy (per build decision): mocked unit + integration tests.
// Microsoft Graph and MSAL are mocked — no live Azure credentials are used.
// Live §13 acceptance against a real provider sandbox is run separately,
// locally, with a real Entra app registration and Outlook mailboxes.
export default defineConfig({
  test: {
    include: ["test/**/*.test.ts"],
    environment: "node",
    coverage: {
      reporter: ["text", "html"],
      include: ["src/**/*.ts"],
      exclude: ["src/**/*.d.ts"],
    },
  },
});
