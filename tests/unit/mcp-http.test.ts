/**
 * Tests for the Streamable HTTP mode of the MCP server (src/mcp-http.ts).
 *
 * Starts a real listener on an ephemeral localhost port with a fake Teams
 * client provider: no authentication and no Teams requests.
 */

import {
  mkdtempSync,
  readdirSync,
  rmSync,
  statSync,
  writeFileSync,
} from "node:fs";
import type { Server } from "node:http";
import type { AddressInfo } from "node:net";
import { tmpdir } from "node:os";
import { join } from "node:path";
import { describe, it, expect, vi, beforeEach, afterEach } from "vitest";
import { Client } from "@modelcontextprotocol/sdk/client/index.js";
import { StreamableHTTPClientTransport } from "@modelcontextprotocol/sdk/client/streamableHttp.js";
import {
  FileShare,
  loadOrCreateHttpToken,
  startMcpHttpServer,
} from "../../src/mcp-http.js";
import { createMcpServer, parseServerArguments } from "../../src/mcp-server.js";
import { FileCredentialStore } from "../../src/credential-store.js";
import type { TeamsClient } from "../../src/teams-client.js";

const token = "test-http-access-token-0123456789abcdefghij";

const fakeClient = {
  listConversations: vi.fn().mockResolvedValue([
    {
      id: "19:dev@thread.v2",
      topic: "Development team",
      threadType: "chat",
      version: 1,
      lastMessageTime: "2026-10-28T09:15:00.000Z",
      memberCount: null,
    },
  ]),
} as unknown as TeamsClient;

let server: Server;
let baseUrl: string;
let fileShare: FileShare;
let fileDirectory: string;
let clock: number;

beforeEach(async () => {
  vi.stubEnv("TEAMS_READ_ONLY", "");
  vi.stubEnv("TEAMS_TELEMETRY", "");
  fileDirectory = mkdtempSync(join(tmpdir(), "teams-api-http-files-"));
  clock = Date.parse("2026-10-29T12:00:00Z");
  const authManager = { getClient: vi.fn().mockResolvedValue(fakeClient) };
  server = await startMcpHttpServer({
    host: "127.0.0.1",
    port: 0,
    token,
    createMcpServer: () =>
      createMcpServer({ authManager, readOnly: true, fileShare }).server,
    get fileShare() {
      return fileShare;
    },
  });
  const { port } = server.address() as AddressInfo;
  baseUrl = `http://127.0.0.1:${port}`;
  fileShare = new FileShare(baseUrl, 60 * 60 * 1_000, () => clock);
});

afterEach(async () => {
  await new Promise<void>((resolve) => server.close(() => resolve()));
  rmSync(fileDirectory, { recursive: true, force: true });
  vi.unstubAllEnvs();
});

function post(path: string, headers: Record<string, string>) {
  return fetch(`${baseUrl}${path}`, {
    method: "POST",
    headers: {
      "Content-Type": "application/json",
      Accept: "application/json, text/event-stream",
      ...headers,
    },
    body: JSON.stringify({
      jsonrpc: "2.0",
      id: 1,
      method: "tools/list",
      params: {},
    }),
  });
}

async function connectClient(accessToken: string): Promise<Client> {
  const client = new Client({ name: "teams-api-http-test", version: "1.0.0" });
  await client.connect(
    new StreamableHTTPClientTransport(new URL(`${baseUrl}/mcp`), {
      requestInit: { headers: { Authorization: `Bearer ${accessToken}` } },
    }),
  );
  return client;
}

describe("MCP over HTTP", () => {
  it("serves the read-only tools to a client with the token", async () => {
    const client = await connectClient(token);

    const names = (await client.listTools()).tools.map((tool) => tool.name);
    const result = await client.callTool({
      name: "teams_list_conversations",
      arguments: { format: "json" },
    });
    await client.close();

    expect(names).toContain("teams_get_messages");
    expect(names).not.toContain("teams_send_message");
    const text = (result.content as Array<{ text: string }>)[0].text;
    expect(JSON.parse(text)).toMatchObject({
      ok: true,
      data: [{ id: "19:dev@thread.v2" }],
    });
  });

  it("rejects a missing or wrong token", async () => {
    const missing = await post("/mcp", {});
    const wrong = await post("/mcp", { Authorization: "Bearer not-the-token" });

    expect(missing.status).toBe(401);
    expect(missing.headers.get("www-authenticate")).toContain("Bearer");
    expect(wrong.status).toBe(401);
    await expect(connectClient("not-the-token")).rejects.toThrow();
  });

  it("refuses browser requests even with the token", async () => {
    const response = await post("/mcp", {
      Authorization: `Bearer ${token}`,
      Origin: "https://evil.example",
    });

    expect(response.status).toBe(403);
  });

  it("only serves /mcp, and only POST", async () => {
    const otherPath = await post("/other", {
      Authorization: `Bearer ${token}`,
    });
    const get = await fetch(`${baseUrl}/mcp`, {
      headers: { Authorization: `Bearer ${token}` },
    });

    expect(otherPath.status).toBe(404);
    expect(get.status).toBe(405);
  });
});

describe("file links", () => {
  function shareFile(fileName: string, contentType: string, body: string) {
    const path = join(fileDirectory, fileName);
    writeFileSync(path, body);
    return fileShare.share({ path, fileName, contentType });
  }

  it("serves a shared file without the MCP token", async () => {
    const link = shareFile("screenshot.png", "image/png", "PNGDATA");

    const response = await fetch(link);

    expect(link).toMatch(/\/files\/[A-Za-z0-9_-]{22}\/screenshot\.png$/);
    expect(response.status).toBe(200);
    expect(response.headers.get("content-type")).toBe("image/png");
    expect(response.headers.get("content-disposition")).toContain("inline");
    expect(await response.text()).toBe("PNGDATA");
  });

  it("forgets links after an hour, and never serves unknown ones", async () => {
    const link = shareFile("screenshot.png", "image/png", "PNGDATA");
    clock += 60 * 60 * 1_000;

    expect((await fetch(link)).status).toBe(404);
    expect((await fetch(`${baseUrl}/files/not-a-real-link/x.png`)).status).toBe(
      404,
    );
  });

  it("only allows GET and HEAD", async () => {
    const link = shareFile("screenshot.png", "image/png", "PNGDATA");

    expect((await fetch(link, { method: "POST" })).status).toBe(405);
    expect((await fetch(link, { method: "HEAD" })).status).toBe(200);
  });

  it("sends types a browser could run as downloads", async () => {
    const link = shareFile("page.html", "text/html", "<script>1</script>");

    const response = await fetch(link);

    expect(response.headers.get("content-type")).toBe(
      "application/octet-stream",
    );
    expect(response.headers.get("content-disposition")).toContain("attachment");
    expect(response.headers.get("content-security-policy")).toBe("sandbox");
  });

  it("adds a link to teams_download_file results", async () => {
    fakeClient.getMessages = vi.fn().mockResolvedValue([
      {
        id: "1761642900000",
        messageType: "RichText/Html",
        senderMri: "8:orgid:aaaaaaaa-0000-0000-0000-000000000001",
        senderDisplayName: "Alice Smith",
        content: "<p>see screenshot</p>",
        originalArrivalTime: "2026-10-28T09:15:00.000Z",
        composeTime: "2026-10-28T09:15:00.000Z",
        editTime: null,
        subject: null,
        isDeleted: false,
        reactions: [],
        followers: [],
        mentions: [],
        quotedMessageId: null,
        images: [
          {
            amsObjectId: "0-test-image",
            url: "https://example.test/imgo",
            fullSizeUrl: "https://example.test/full",
            width: 100,
            height: 50,
            contentPosition: 0,
          },
        ],
        files: [],
      },
    ]);
    fakeClient.downloadImage = vi.fn().mockResolvedValue({
      data: Buffer.from("PNGDATA"),
      contentType: "image/png",
      size: 7,
    });
    const client = await connectClient(token);

    const result = await client.callTool({
      name: "teams_download_file",
      arguments: {
        conversationId: "19:dev@thread.v2",
        messageId: "1761642900000",
        outputDirectory: fileDirectory,
      },
    });
    await client.close();

    const text = (result.content as Array<{ type: string; text?: string }>)
      .filter((block) => block.type === "text")
      .map((block) => block.text)
      .join("\n");
    const link = /Link \(valid for 1 hour\): (\S+)/.exec(text)?.[1];
    expect(link).toMatch(/^http:\/\/127\.0\.0\.1:\d+\/files\//);
    expect(await (await fetch(link!)).text()).toBe("PNGDATA");
  });
});

describe("loadOrCreateHttpToken", () => {
  let directory: string;

  beforeEach(() => {
    directory = mkdtempSync(join(tmpdir(), "teams-api-http-token-"));
  });

  afterEach(() => {
    rmSync(directory, { recursive: true, force: true });
  });

  it("creates a strong token once, stored privately, and reuses it", async () => {
    const store = new FileCredentialStore(join(directory, "tokens"));

    const first = await loadOrCreateHttpToken(store, {});
    const second = await loadOrCreateHttpToken(store, {});

    expect(first).toMatch(/^[A-Za-z0-9_-]{43}$/);
    expect(second).toBe(first);
    const [fileName] = readdirSync(join(directory, "tokens"));
    if (process.platform !== "win32") {
      expect(statSync(join(directory, "tokens", fileName)).mode & 0o777).toBe(
        0o600,
      );
    }
  });

  it("uses TEAMS_MCP_HTTP_TOKEN when set, and rejects short ones", async () => {
    const store = new FileCredentialStore(join(directory, "tokens"));

    await expect(
      loadOrCreateHttpToken(store, { TEAMS_MCP_HTTP_TOKEN: token }),
    ).resolves.toBe(token);
    await expect(
      loadOrCreateHttpToken(store, { TEAMS_MCP_HTTP_TOKEN: "short" }),
    ).rejects.toThrow("at least 32 characters");
  });
});

describe("parseServerArguments", () => {
  it("defaults to stdio", () => {
    expect(parseServerArguments([])).toEqual({
      http: false,
      host: "127.0.0.1",
      port: 8787,
      allowWrites: false,
      printHttpToken: false,
      autoRefresh: true,
      refreshMarginMs: 45 * 60 * 1_000,
      publicUrl: "http://host.docker.internal:8787",
    });
  });

  it("parses HTTP options", () => {
    expect(
      parseServerArguments(["--http", "--port", "9000", "--host", "0.0.0.0"]),
    ).toMatchObject({ http: true, port: 9000, host: "0.0.0.0" });
  });

  it("parses refresh options", () => {
    expect(
      parseServerArguments(["--http", "--refresh-margin", "2h"]),
    ).toMatchObject({
      autoRefresh: true,
      refreshMarginMs: 2 * 60 * 60 * 1_000,
    });
    expect(parseServerArguments(["--http", "--no-auto-refresh"])).toMatchObject(
      { autoRefresh: false },
    );
  });

  it.each([
    ["--port", "0"],
    ["--port", "http"],
    ["--htpp"],
    ["--refresh-margin", "soon"],
  ])("rejects %s", (...argv) => {
    expect(() => parseServerArguments(argv)).toThrow();
  });
});
