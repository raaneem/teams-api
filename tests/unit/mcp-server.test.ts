/**
 * MCP server tests (src/mcp-server.ts) over an in-memory transport,
 * with a fake client provider: no authentication and no network.
 */

import { describe, it, expect, vi, beforeEach, afterEach } from "vitest";
import { Client } from "@modelcontextprotocol/sdk/client/index.js";
import { InMemoryTransport } from "@modelcontextprotocol/sdk/inMemory.js";
import {
  createMcpServer,
  type TeamsClientProvider,
} from "../../src/mcp-server.js";
import type { TeamsClient } from "../../src/teams-client.js";

const writeTools = [
  "teams_send_message",
  "teams_edit_message",
  "teams_delete_message",
  "teams_add_reaction",
  "teams_remove_reaction",
  "teams_describe_image",
];

async function connect(readOnly: boolean, authManager: TeamsClientProvider) {
  const { server } = createMcpServer({ readOnly, authManager });
  const [clientTransport, serverTransport] =
    InMemoryTransport.createLinkedPair();
  const client = new Client({ name: "teams-api-test", version: "1.0.0" });
  await Promise.all([
    server.connect(serverTransport),
    client.connect(clientTransport),
  ]);
  return client;
}

function textOf(result: Awaited<ReturnType<Client["callTool"]>>): string {
  return (result.content as Array<{ type: string; text?: string }>)
    .map((block) => block.text ?? "")
    .join("");
}

beforeEach(() => {
  vi.stubEnv("TEAMS_READ_ONLY", "");
  vi.stubEnv("TEAMS_TELEMETRY", "");
});

afterEach(() => {
  vi.unstubAllEnvs();
});

describe("MCP server", () => {
  it("hides every tool that could change Teams in read-only mode", async () => {
    const client = await connect(true, { getClient: vi.fn() });

    const names = (await client.listTools()).tools.map((tool) => tool.name);

    expect(names).toContain("teams_get_messages");
    expect(names).toContain("teams_search_messages");
    for (const tool of writeTools) {
      expect(names).not.toContain(tool);
    }
    expect(client.getInstructions()).toContain("read-only mode");
  });

  it("exposes write tools when read-only mode is off", async () => {
    const client = await connect(false, { getClient: vi.fn() });

    const names = (await client.listTools()).tools.map((tool) => tool.name);

    for (const tool of writeTools) {
      expect(names).toContain(tool);
    }
  });

  it("refuses a write tool if read-only mode is switched on after startup", async () => {
    const getClient = vi.fn();
    const client = await connect(false, { getClient });
    vi.stubEnv("TEAMS_READ_ONLY", "true");

    const result = await client.callTool({
      name: "teams_send_message",
      arguments: { conversationId: "19:x@thread.v2", content: "hi" },
    });

    expect(result.isError).toBe(true);
    expect(textOf(result)).toContain("read-only mode");
    expect(getClient).not.toHaveBeenCalled();
  });

  it("returns redacted errors instead of passing raw messages to the SDK", async () => {
    const bundle = Buffer.from(
      JSON.stringify({ skypeToken: "CANARY-MCP-SKYPE-TOKEN-0123456789" }),
    ).toString("base64");
    const client = await connect(false, {
      getClient: vi.fn(async () => {
        throw new Error(
          `Command failed: security add-generic-password -a me@contoso.com -s teams-api -w ${bundle} -U\n` +
            "security: SecKeychainItemCreateFromContent (<default>): User interaction is not allowed.",
        );
      }),
    });

    const result = await client.callTool({
      name: "teams_list_conversations",
      arguments: {},
    });

    expect(result.isError).toBe(true);
    expect(JSON.stringify(result)).not.toContain(bundle);
    expect(textOf(result)).toContain("-w [REDACTED]");
  });

  it("supports the json format", async () => {
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
    const client = await connect(true, {
      getClient: vi.fn().mockResolvedValue(fakeClient),
    });

    const result = await client.callTool({
      name: "teams_list_conversations",
      arguments: { format: "json" },
    });

    expect(result.isError).toBeFalsy();
    expect(JSON.parse(textOf(result))).toMatchObject({
      schemaVersion: 1,
      ok: true,
      command: "list-conversations",
      data: [{ id: "19:dev@thread.v2", topic: "Development team" }],
    });
  });
});
