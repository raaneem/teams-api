/**
 * Unit tests for the SSH forced-command wrapper (src/ssh-wrapper.ts).
 */

import { describe, it, expect } from "vitest";
import {
  ALLOWED_COMMANDS,
  buildCliArguments,
  buildCliEnvironment,
  splitCommandLine,
} from "../../src/ssh-wrapper.js";
import { InvalidArgumentError } from "../../src/errors.js";

describe("splitCommandLine", () => {
  it("splits on whitespace and honors quotes and escapes", () => {
    expect(
      splitCommandLine(
        `search-messages --query "release plan" --from 'a@b.c'  --x a\\ b ""`,
      ),
    ).toEqual([
      "search-messages",
      "--query",
      "release plan",
      "--from",
      "a@b.c",
      "--x",
      "a b",
      "",
    ]);
  });

  it("never expands anything", () => {
    expect(splitCommandLine("get-messages --chat $(rm -rf ~);`id`|*")).toEqual([
      "get-messages",
      "--chat",
      "$(rm",
      "-rf",
      "~);`id`|*",
    ]);
  });

  it("rejects an unterminated quote", () => {
    expect(() => splitCommandLine(`get-messages --chat "Dev`)).toThrow(
      InvalidArgumentError,
    );
  });
});

describe("buildCliArguments", () => {
  it("passes allowed read commands through", () => {
    expect(
      buildCliArguments(
        "get-messages --conversation-id 19:abc@thread.v2 --since 24h --json",
      ),
    ).toEqual([
      "get-messages",
      "--conversation-id",
      "19:abc@thread.v2",
      "--since",
      "24h",
      "--json",
    ]);
  });

  it("accepts a leading teams-api", () => {
    expect(buildCliArguments("teams-api list-conversations --json")).toEqual([
      "list-conversations",
      "--json",
    ]);
  });

  it.each([
    "send-message --chat Dev --content hi",
    "edit-message --chat Dev --message-id 1 --content hi",
    "delete-message --chat Dev --message-id 1",
    "add-reaction --chat Dev --message-id 1 --reaction-key like",
    "auth",
    "auth --print-token",
    "logout --email a@b.c",
    "download-file --chat Dev --message-id 1",
    "describe-image --ams-object-id x",
  ])("refuses %s", (command) => {
    expect(() => buildCliArguments(command)).toThrow(InvalidArgumentError);
  });

  it.each([
    "get-messages --chat Dev --token abc",
    "get-messages --chat Dev --token=abc",
    "get-messages --chat Dev --login",
    "get-messages --chat Dev --output /etc/passwd",
    "list-conversations --debug-port 9222",
  ])("refuses options that bypass the cache or write files: %s", (command) => {
    expect(() => buildCliArguments(command)).toThrow(/not allowed over SSH/);
  });

  it("refuses an empty command", () => {
    expect(() => buildCliArguments(undefined)).toThrow(/No command given/);
    expect(() => buildCliArguments("   ")).toThrow(/No command given/);
  });

  it("only allows read commands", () => {
    expect([...ALLOWED_COMMANDS].sort()).toEqual([
      "find-conversation",
      "find-one-on-one",
      "get-members",
      "get-messages",
      "get-transcript",
      "list-conversations",
      "search-messages",
      "whoami",
    ]);
  });
});

describe("buildCliEnvironment", () => {
  it("forces read-only, non-interactive, file-cache mode and drops overrides", () => {
    const environment = buildCliEnvironment({
      HOME: "/Users/someone",
      TEAMS_READ_ONLY: "false",
      TEAMS_TOKEN_STORE: "system",
      TEAMS_TOKEN: "injected",
      TEAMS_TELEMETRY: "true",
      DEBUG: "pw:*",
      NODE_OPTIONS: "--require /tmp/x.js",
    });

    expect(environment).toEqual({
      HOME: "/Users/someone",
      TEAMS_READ_ONLY: "true",
      TEAMS_NON_INTERACTIVE: "true",
      TEAMS_TOKEN_STORE: "file",
    });
  });
});
