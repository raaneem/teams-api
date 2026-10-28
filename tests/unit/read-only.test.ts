/**
 * Unit tests for read-only mode (src/read-only.ts, src/runtime-flags.ts) and
 * its enforcement at the network and client layers.
 */

import { describe, it, expect, vi, beforeEach, afterEach } from "vitest";
import {
  isFlagEnabled,
  isReadOnlyMode,
  resolveTokenStoreMode,
} from "../../src/runtime-flags.js";
import {
  ReadOnlyModeError,
  isRequestAllowedInReadOnlyMode,
} from "../../src/read-only.js";
import { fetchWithRetry } from "../../src/api/common.js";
import {
  postMessage,
  editMessage,
  deleteMessage,
  addReaction,
  removeReaction,
  postScheduledMessage,
  fetchConversations,
} from "../../src/api/chat-service.js";
import { fetchProfiles } from "../../src/api/middle-tier.js";
import { searchMessages } from "../../src/api/message-search.js";
import { TeamsClient } from "../../src/teams-client.js";
import type { TeamsToken } from "../../src/types.js";

vi.mock("../../src/emoji-map.js", () => ({
  initializeEmojiMap: vi.fn().mockResolvedValue(undefined),
  resolveReactionKey: vi.fn((key: string) => key),
}));

const token: TeamsToken = {
  skypeToken: "read-only-test-skype-token",
  region: "apac",
  bearerToken: "read-only-test-bearer-token",
  substrateToken: "read-only-test-substrate-token",
};

const originalFetch = globalThis.fetch;
let fetchMock: ReturnType<typeof vi.fn>;

beforeEach(() => {
  fetchMock = vi.fn().mockResolvedValue({
    ok: true,
    status: 200,
    statusText: "OK",
    headers: new Headers(),
    json: async () => ({ conversations: [], value: [] }),
    text: async () => "",
  });
  globalThis.fetch = fetchMock as unknown as typeof fetch;
});

afterEach(() => {
  globalThis.fetch = originalFetch;
  vi.unstubAllEnvs();
});

describe("isFlagEnabled", () => {
  it.each(["true", "1", "yes", "on", "TRUE", "anything"])(
    "treats %s as on (fails closed)",
    (value) => {
      expect(isFlagEnabled(value)).toBe(true);
    },
  );

  it.each([undefined, "", "0", "false", "no", "off", " False "])(
    "treats %s as off",
    (value) => {
      expect(isFlagEnabled(value)).toBe(false);
    },
  );

  it("reads TEAMS_READ_ONLY", () => {
    expect(isReadOnlyMode({ TEAMS_READ_ONLY: "true" })).toBe(true);
    expect(isReadOnlyMode({})).toBe(false);
  });
});

describe("resolveTokenStoreMode", () => {
  it.each([
    [undefined, "auto"],
    ["auto", "auto"],
    ["file", "file"],
    ["FILE", "file"],
    ["system", "system"],
    ["keychain", "system"],
    ["unknown", "auto"],
  ])("maps %s to %s", (value, expected) => {
    expect(resolveTokenStoreMode({ TEAMS_TOKEN_STORE: value })).toBe(expected);
  });
});

describe("isRequestAllowedInReadOnlyMode", () => {
  it.each([
    [
      "GET",
      "https://apac.ng.msg.teams.microsoft.com/v1/users/ME/conversations",
    ],
    ["HEAD", "https://example.com/"],
    [
      "POST",
      "https://teams.cloud.microsoft/api/mt/apac/beta/users/fetchShortProfile?isMailAddress=false",
    ],
    [
      "POST",
      "https://substrate.office.com/search/api/v1/suggestions?scenario=peoplepicker.newChat",
    ],
    ["POST", "https://substrate.office.com/searchservice/api/v2/query"],
  ])("allows %s %s", (method, url) => {
    expect(isRequestAllowedInReadOnlyMode(url, method)).toBe(true);
  });

  it.each([
    [
      "POST",
      "https://apac.ng.msg.teams.microsoft.com/v1/users/ME/conversations/19:x/messages",
    ],
    [
      "PUT",
      "https://apac.ng.msg.teams.microsoft.com/v1/users/ME/conversations/19:x/messages/1",
    ],
    [
      "DELETE",
      "https://apac.ng.msg.teams.microsoft.com/v1/users/ME/conversations/19:x/messages/1",
    ],
    ["POST", "https://api.openai.com/v1/chat/completions"],
    ["POST", "http://substrate.office.com/searchservice/api/v2/query"],
    [
      "POST",
      "https://substrate.office.com.evil.test/searchservice/api/v2/query",
    ],
    [
      "PATCH",
      "https://teams.cloud.microsoft/api/mt/apac/beta/users/fetchShortProfile",
    ],
  ])("refuses %s %s", (method, url) => {
    expect(isRequestAllowedInReadOnlyMode(url, method)).toBe(false);
  });
});

describe("fetchWithRetry in read-only mode", () => {
  it("refuses a write before sending it", async () => {
    vi.stubEnv("TEAMS_READ_ONLY", "true");

    await expect(
      fetchWithRetry("https://apac.ng.msg.teams.microsoft.com/v1/x", {
        method: "POST",
      }),
    ).rejects.toBeInstanceOf(ReadOnlyModeError);
    expect(fetchMock).not.toHaveBeenCalled();
  });

  it("sends writes normally when read-only mode is off", async () => {
    await fetchWithRetry("https://apac.ng.msg.teams.microsoft.com/v1/x", {
      method: "POST",
    });
    expect(fetchMock).toHaveBeenCalledTimes(1);
  });
});

describe("Chat Service writes in read-only mode", () => {
  beforeEach(() => {
    vi.stubEnv("TEAMS_READ_ONLY", "true");
  });

  it.each([
    ["postMessage", () => postMessage(token, "19:x@thread.v2", "hi", "Alice")],
    [
      "editMessage",
      () => editMessage(token, "19:x@thread.v2", "1", "hi", "Alice"),
    ],
    ["deleteMessage", () => deleteMessage(token, "19:x@thread.v2", "1")],
    ["addReaction", () => addReaction(token, "19:x@thread.v2", "1", "like")],
    [
      "removeReaction",
      () => removeReaction(token, "19:x@thread.v2", "1", "like"),
    ],
    [
      "postScheduledMessage",
      () =>
        postScheduledMessage(
          token,
          "19:x@thread.v2",
          "hi",
          "Alice",
          new Date("2030-01-01T00:00:00Z"),
        ),
    ],
  ])("%s is refused without a request", async (_name, call) => {
    await expect(call()).rejects.toBeInstanceOf(ReadOnlyModeError);
    expect(fetchMock).not.toHaveBeenCalled();
  });

  it("still allows reads, profile lookups, and search", async () => {
    await fetchConversations(token, 10);
    await fetchProfiles(token, [
      "8:orgid:00000000-0000-0000-0000-000000000001",
    ]);
    fetchMock.mockResolvedValueOnce({
      ok: true,
      status: 200,
      json: async () => ({
        EntitySets: [{ EntityType: "Message", ResultSets: [] }],
      }),
    });
    await searchMessages(token, { query: "release" });

    expect(fetchMock).toHaveBeenCalledTimes(3);
  });
});

describe("TeamsClient writes in read-only mode", () => {
  beforeEach(() => {
    vi.stubEnv("TEAMS_READ_ONLY", "true");
  });

  it.each([
    ["sendMessage", (client: TeamsClient) => client.sendMessage("19:x", "hi")],
    [
      "editMessage",
      (client: TeamsClient) => client.editMessage("19:x", "1", "hi"),
    ],
    [
      "deleteMessage",
      (client: TeamsClient) => client.deleteMessage("19:x", "1"),
    ],
    [
      "addReaction",
      (client: TeamsClient) => client.addReaction("19:x", "1", "like"),
    ],
    [
      "removeReaction",
      (client: TeamsClient) => client.removeReaction("19:x", "1", "like"),
    ],
    [
      "scheduleMessage",
      (client: TeamsClient) =>
        client.scheduleMessage("19:x", "hi", new Date("2030-01-01T00:00:00Z")),
    ],
    [
      "sendMessageWithImages",
      (client: TeamsClient) =>
        client.sendMessageWithImages("19:x", [{ type: "text", text: "hi" }]),
    ],
    [
      "sendMessageWithFiles",
      (client: TeamsClient) =>
        client.sendMessageWithFiles("19:x", [{ type: "text", text: "hi" }]),
    ],
  ])("%s throws before any request", async (_name, call) => {
    const client = TeamsClient.fromToken(token);

    await expect(call(client)).rejects.toBeInstanceOf(ReadOnlyModeError);
    expect(fetchMock).not.toHaveBeenCalled();
  });
});
