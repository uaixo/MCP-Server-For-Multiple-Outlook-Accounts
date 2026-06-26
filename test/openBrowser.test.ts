import { afterEach, describe, expect, it, vi } from "vitest";
import { EventEmitter } from "node:events";

const spawnMock = vi.hoisted(() => vi.fn());
vi.mock("node:child_process", () => ({ spawn: spawnMock }));

const { defaultOpenBrowser } = await import("../src/auth/msalClient.js");

describe("defaultOpenBrowser (NFR-OPS-1)", () => {
  afterEach(() => vi.restoreAllMocks());

  it("does not crash when the opener binary is missing, and prints the URL to stderr", async () => {
    const fake = Object.assign(new EventEmitter(), { unref: vi.fn() });
    spawnMock.mockReturnValue(fake);
    const writes: string[] = [];
    vi.spyOn(process.stderr, "write").mockImplementation((chunk: unknown) => {
      writes.push(String(chunk));
      return true;
    });

    const url = "https://login.microsoftonline.com/common/consent?x=1";
    await defaultOpenBrowser(url);
    // Without the 'error' listener this emit would throw an uncaught exception.
    fake.emit("error", new Error("spawn xdg-open ENOENT"));

    expect(writes.join("")).toContain(url);
    expect(fake.unref).toHaveBeenCalled();
  });
});
