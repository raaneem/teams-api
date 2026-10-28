/**
 * CLI tests (src/cli.ts), run in-process.
 *
 * Nothing here touches the real Keychain, the real ~/.teams-api, a browser,
 * or Teams: `security` is mocked to behave like the locked login keychain of
 * an SSH session, HOME is a temporary directory, browser login is replaced
 * with a stub that fails the test, and `fetch` serves fixtures and rejects
 * any other host.
 */

import { mkdtempSync, readdirSync, rmSync, statSync } from "node:fs";
import { tmpdir } from "node:os";
import { join } from "node:path";
import { format } from "node:util";
import { describe, it, expect, vi, beforeEach, afterEach } from "vitest";

const execFileSync = vi.hoisted(() => vi.fn());

vi.mock("node:child_process", async (importOriginal) => ({
  ...(await importOriginal<typeof import("node:child_process")>()),
  execFileSync,
}));

vi.mock("../../src/auth/interactive.js", () => ({
  acquireTokenViaInteractiveLogin: vi.fn(async () => {
    throw new Error("browser login must not run in CLI tests");
  }),
}));

vi.mock("../../src/emoji-map.js", () => ({
  initializeEmojiMap: vi.fn().mockResolvedValue(undefined),
  resolveReactionKey: vi.fn((key: string) => key),
}));

import { runCli } from "../../src/cli.js";
import { resetTokenStoreState } from "../../src/token-store.js";

const cachedToken = {
  skypeToken: "CANARY-SKYPE-TOKEN-0123456789abcdef",
  region: "apac",
  bearerToken: "CANARY-BEARER-TOKEN-0123456789abcdef",
  substrateToken: "CANARY-SUBSTRATE-TOKEN-0123456789abcdef",
  amsToken: "CANARY-AMS-TOKEN-0123456789abcdef",
  sharePointHost: "contoso-my.sharepoint.com",
};
const secrets = [
  cachedToken.skypeToken,
  cachedToken.bearerToken,
  cachedToken.substrateToken,
  cachedToken.amsToken,
];

const lockedKeychainStderr =
  "security: SecKeychainItemCreateFromContent (<default>): User interaction is not allowed.\n";

const originalPlatform = process.platform;
const originalFetch = globalThis.fetch;

let home: string;
let stdout: string[];
let stderr: string[];
let unexpectedRequests: string[];
let routes: Map<string, () => unknown>;

function encodedCachedToken(): string {
  return Buffer.from(
    JSON.stringify({ ...cachedToken, acquiredAt: Date.now() }),
  ).toString("base64");
}

/** `security` as it behaves over SSH: reads may work, writes are refused. */
function simulateLockedKeychain(cached: string | null): void {
  execFileSync.mockImplementation((file: string, args: string[]) => {
    if (file !== "security") {
      throw new Error(`unexpected command: ${file}`);
    }
    if (args[0] === "find-generic-password" && cached) {
      return cached;
    }
    throw Object.assign(
      new Error(
        `Command failed: ${file} ${args.join(" ")}\n${lockedKeychainStderr}`,
      ),
      { status: 36, stderr: Buffer.from(lockedKeychainStderr) },
    );
  });
}

function captureStream(stream: NodeJS.WriteStream, sink: string[]): void {
  vi.spyOn(stream, "write").mockImplementation(((
    chunk: unknown,
    encodingOrCallback?: unknown,
    callback?: unknown,
  ) => {
    sink.push(String(chunk));
    const done =
      typeof encodingOrCallback === "function" ? encodingOrCallback : callback;
    if (typeof done === "function") {
      (done as () => void)();
    }
    return true;
  }) as typeof stream.write);
}

function jsonResponse(body: unknown, status = 200, statusText = "OK") {
  return {
    ok: status >= 200 && status < 300,
    status,
    statusText,
    headers: new Headers(),
    json: async () => body,
    text: async () => JSON.stringify(body),
  };
}

async function run(...args: string[]): Promise<number> {
  return runCli(["node", "teams-api", ...args]);
}

function allOutput(): string {
  return stdout.join("") + stderr.join("");
}

function expectNoSecrets(): void {
  const output = allOutput();
  for (const secret of secrets) {
    expect(output).not.toContain(secret);
  }
}

beforeEach(() => {
  home = mkdtempSync(join(tmpdir(), "teams-api-cli-home-"));
  vi.stubEnv("HOME", home);
  vi.stubEnv("TEAMS_READ_ONLY", "");
  vi.stubEnv("TEAMS_TOKEN_STORE", "");
  vi.stubEnv("TEAMS_NON_INTERACTIVE", "");
  Object.defineProperty(process, "platform", { value: "darwin" });
  resetTokenStoreState();
  execFileSync.mockReset();

  stdout = [];
  stderr = [];
  captureStream(process.stdout, stdout);
  captureStream(process.stderr, stderr);
  // vitest routes console.* around process streams; capture it too.
  for (const method of ["log", "info"] as const) {
    vi.spyOn(console, method).mockImplementation((...args: unknown[]) => {
      stdout.push(`${format(...args)}\n`);
    });
  }
  for (const method of ["error", "warn", "debug"] as const) {
    vi.spyOn(console, method).mockImplementation((...args: unknown[]) => {
      stderr.push(`${format(...args)}\n`);
    });
  }

  unexpectedRequests = [];
  routes = new Map();
  globalThis.fetch = vi.fn(async (input: unknown, init?: RequestInit) => {
    const url = String(input);
    const route = routes.get(`${init?.method ?? "GET"} ${url}`);
    if (!route) {
      unexpectedRequests.push(`${init?.method ?? "GET"} ${url}`);
      return jsonResponse({ error: "unexpected request" }, 599, "Unexpected");
    }
    return route();
  }) as unknown as typeof fetch;
});

afterEach(() => {
  vi.restoreAllMocks();
  vi.unstubAllEnvs();
  globalThis.fetch = originalFetch;
  Object.defineProperty(process, "platform", { value: originalPlatform });
  rmSync(home, { recursive: true, force: true });
});

describe("locked Keychain over SSH", () => {
  it("falls back to the private token file instead of crashing, and prints no token", async () => {
    simulateLockedKeychain(encodedCachedToken());
    routes.set(
      "GET https://emea.ng.msg.teams.microsoft.com/v1/users/ME/conversations?view=mychats&pageSize=50",
      () =>
        jsonResponse({
          conversations: [
            {
              id: "19:dev@thread.v2",
              version: 1,
              threadProperties: {
                topic: "Development team",
                threadType: "chat",
              },
              properties: { lastimreceivedtime: "2026-10-28T09:15:00.000Z" },
            },
          ],
        }),
    );

    // The region override makes the client re-save the cached token, which
    // the locked Keychain refuses: the exact path that used to print it.
    const exitCode = await run(
      "list-conversations",
      "--region",
      "emea",
      "--json",
    );

    expect(exitCode).toBe(0);
    expect(unexpectedRequests).toEqual([]);
    const envelope = JSON.parse(stdout.join(""));
    expect(envelope).toMatchObject({
      schemaVersion: 1,
      ok: true,
      command: "list-conversations",
      data: [
        {
          id: "19:dev@thread.v2",
          topic: "Development team",
          type: "chat",
          lastMessageAt: "2026-10-28T09:15:00.000Z",
        },
      ],
    });
    expect(stderr.join("")).toContain("User interaction is not allowed");
    expect(stderr.join("")).toContain("saved it to the token file");
    expectNoSecrets();

    const tokenDirectory = join(home, ".teams-api", "tokens");
    const files = readdirSync(tokenDirectory);
    expect(files).toHaveLength(1);
    expect(statSync(tokenDirectory).mode & 0o777).toBe(0o700);
    expect(statSync(join(tokenDirectory, files[0])).mode & 0o777).toBe(0o600);
  });

  it("reads the token back from the file cache on the next run", async () => {
    simulateLockedKeychain(encodedCachedToken());
    routes.set(
      "GET https://emea.ng.msg.teams.microsoft.com/v1/users/ME/conversations?view=mychats&pageSize=50",
      () => jsonResponse({ conversations: [] }),
    );
    await run("list-conversations", "--region", "emea", "--json");

    // Next run: the Keychain can no longer be read at all.
    resetTokenStoreState();
    simulateLockedKeychain(null);
    stdout.length = 0;
    const exitCode = await run("list-conversations", "--json");

    expect(exitCode).toBe(0);
    expect(JSON.parse(stdout.join("")).ok).toBe(true);
    expect(unexpectedRequests).toEqual([]);
    expectNoSecrets();
  });

  it("redacts tokens that appear in API errors", async () => {
    simulateLockedKeychain(encodedCachedToken());
    routes.set(
      "GET https://apac.ng.msg.teams.microsoft.com/v1/users/ME/conversations?view=mychats&pageSize=50",
      () =>
        jsonResponse(
          {},
          500,
          `echoed skypetoken=${cachedToken.skypeToken} ${cachedToken.bearerToken}`,
        ),
    );

    const exitCode = await run("list-conversations");

    expect(exitCode).toBe(1);
    expect(stderr.join("")).toContain("Failed to fetch conversations: 500");
    expect(stderr.join("")).toContain("[REDACTED]");
    expectNoSecrets();
  });
});

describe("read-only mode", () => {
  it.each([
    [
      "TEAMS_READ_ONLY",
      [
        "send-message",
        "--conversation-id",
        "19:x@thread.v2",
        "--content",
        "hi",
        "--json",
      ],
    ],
    [
      "--read-only",
      [
        "delete-message",
        "--conversation-id",
        "19:x@thread.v2",
        "--message-id",
        "1",
        "--read-only",
        "--json",
      ],
    ],
  ])(
    "refuses write commands before signing in (%s)",
    async (switchName, args) => {
      if (switchName === "TEAMS_READ_ONLY") {
        vi.stubEnv("TEAMS_READ_ONLY", "true");
      }
      simulateLockedKeychain(encodedCachedToken());

      const exitCode = await run(...args);

      expect(exitCode).toBe(4);
      expect(JSON.parse(stdout.join(""))).toMatchObject({
        ok: false,
        error: { code: "READ_ONLY" },
      });
      expect(execFileSync).not.toHaveBeenCalled();
      expect(globalThis.fetch).not.toHaveBeenCalled();
    },
  );

  it("still runs read commands", async () => {
    vi.stubEnv("TEAMS_READ_ONLY", "true");
    simulateLockedKeychain(encodedCachedToken());
    routes.set(
      "GET https://apac.ng.msg.teams.microsoft.com/v1/users/ME/conversations?view=mychats&pageSize=50",
      () => jsonResponse({ conversations: [] }),
    );

    await expect(run("list-conversations", "--json")).resolves.toBe(0);
  });

  it("refuses to print tokens", async () => {
    vi.stubEnv("TEAMS_READ_ONLY", "true");
    simulateLockedKeychain(encodedCachedToken());

    const exitCode = await run("auth", "--print-token");

    expect(exitCode).toBe(4);
    expectNoSecrets();
  });
});

describe("get-messages --json", () => {
  beforeEach(() => {
    simulateLockedKeychain(encodedCachedToken());
    routes.set(
      `GET https://apac.ng.msg.teams.microsoft.com/v1/users/ME/conversations/${encodeURIComponent("19:dev@thread.v2")}/messages?pageSize=200`,
      () =>
        jsonResponse({
          messages: [
            {
              id: "1761642900000",
              messagetype: "RichText/Html",
              from: "https://apac.ng.msg.teams.microsoft.com/v1/users/ME/contacts/8:orgid:aaaaaaaa-0000-0000-0000-000000000001",
              imdisplayname: "Alice Smith",
              content: "<p>The login page shows a blank screen</p>",
              originalarrivaltime: "2026-10-28T09:15:00.000Z",
              composetime: "2026-10-28T09:15:00.000Z",
              properties: {},
            },
            {
              id: "1761638400000",
              messagetype: "Text",
              from: "https://apac.ng.msg.teams.microsoft.com/v1/users/ME/contacts/8:orgid:aaaaaaaa-0000-0000-0000-000000000002",
              imdisplayname: "Bob Johnson",
              content: "Reports export is slow",
              originalarrivaltime: "2026-10-28T08:00:00.000Z",
              composetime: "2026-10-28T08:00:00.000Z",
              properties: {},
            },
            {
              id: "1761595200000",
              messagetype: "Text",
              from: "https://apac.ng.msg.teams.microsoft.com/v1/users/ME/contacts/8:orgid:aaaaaaaa-0000-0000-0000-000000000002",
              imdisplayname: "Bob Johnson",
              content: "Old message",
              originalarrivaltime: "2026-10-27T20:00:00.000Z",
              composetime: "2026-10-27T20:00:00.000Z",
              properties: {},
            },
          ],
          _metadata: {},
        }),
    );
  });

  it("returns stable IDs, links, and timestamps for messages since a time", async () => {
    const exitCode = await run(
      "get-messages",
      "--conversation-id",
      "19:dev@thread.v2",
      "--since",
      "2026-10-28T05:00:00+05:00",
      "--json",
    );

    expect(exitCode).toBe(0);
    expect(unexpectedRequests).toEqual([]);
    const envelope = JSON.parse(stdout.join(""));
    expect(envelope.meta).toEqual({
      conversationId: "19:dev@thread.v2",
      conversationLabel: "19:dev@thread.v2",
      threadRootId: null,
      since: "2026-10-28T00:00:00.000Z",
      order: "oldest-first",
      count: 2,
      newestCreatedAt: "2026-10-28T09:15:00.000Z",
      limitReached: false,
    });
    expect(envelope.data.map((message: { id: string }) => message.id)).toEqual([
      "1761638400000",
      "1761642900000",
    ]);
    expect(envelope.data[1]).toMatchObject({
      id: "1761642900000",
      conversationId: "19:dev@thread.v2",
      link: "https://teams.microsoft.com/l/message/19%3Adev%40thread.v2/1761642900000",
      createdAt: "2026-10-28T09:15:00.000Z",
      sender: {
        id: "8:orgid:aaaaaaaa-0000-0000-0000-000000000001",
        name: "Alice Smith",
      },
      text: "The login page shows a blank screen",
    });
  });

  it("rejects an invalid --since with a JSON error and exit code 2", async () => {
    const exitCode = await run(
      "get-messages",
      "--conversation-id",
      "19:dev@thread.v2",
      "--since",
      "yesterday",
      "--json",
    );

    expect(exitCode).toBe(2);
    expect(JSON.parse(stdout.join(""))).toMatchObject({
      ok: false,
      command: "get-messages",
      error: { code: "INVALID_ARGUMENT" },
    });
    // Rejected before signing in: no Keychain read, no browser, no network.
    expect(execFileSync).not.toHaveBeenCalled();
    expect(globalThis.fetch).not.toHaveBeenCalled();
  });
});

describe("auth", () => {
  it("prints which tokens are cached, never the tokens", async () => {
    simulateLockedKeychain(encodedCachedToken());

    const exitCode = await run("auth");

    expect(exitCode).toBe(0);
    const status = JSON.parse(stdout.join(""));
    expect(status).toMatchObject({
      authenticated: true,
      region: "apac",
      tokens: {
        skype: { present: true },
        middleTier: { present: true },
        substrate: { present: true },
        sharePoint: { present: false },
      },
    });
    expectNoSecrets();
  });
});

describe("usage errors", () => {
  it("rejects an unknown format with exit code 2", async () => {
    simulateLockedKeychain(encodedCachedToken());

    const exitCode = await run("list-conversations", "--format", "yaml");

    expect(exitCode).toBe(2);
    expect(stderr.join("")).toContain('Invalid format "yaml"');
  });

  it("requires --region with --token, without echoing the token", async () => {
    const exitCode = await run(
      "list-conversations",
      "--token",
      cachedToken.skypeToken,
    );

    expect(exitCode).toBe(2);
    expect(stderr.join("")).toContain("--region is required");
    expectNoSecrets();
  });
});
