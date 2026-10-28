#!/usr/bin/env npx tsx
/**
 * teams-api-ssh — forced-command wrapper for unattended, read-only use over SSH.
 *
 * Put it in front of a bot's SSH key in ~/.ssh/authorized_keys (absolute
 * paths, because non-interactive SSH sessions usually lack nvm on PATH):
 *
 *   restrict,command="/path/to/bin/node /path/to/lib/node_modules/teams-api/dist/ssh-wrapper.js" ssh-ed25519 AAAA... teams-bot
 *
 * The bot then runs, for example:
 *
 *   ssh mac 'get-messages --conversation-id 19:abc@thread.v2 --since 24h --json'
 *
 * A leading "teams-api" is optional. The wrapper:
 *   - allows only read commands (ALLOWED_COMMANDS)
 *   - refuses options that change authentication or write local files
 *   - forces TEAMS_READ_ONLY, TEAMS_NON_INTERACTIVE, and TEAMS_TOKEN_STORE=file
 *   - never runs a shell: the command line is split into arguments without
 *     any expansion and passed straight to the CLI
 *
 * `restrict` also disables PTY allocation, so stderr never mixes into stdout.
 */

import { spawnSync } from "node:child_process";
import { join } from "node:path";
import { InvalidArgumentError } from "./errors.js";
import { safeErrorMessage } from "./redact.js";

/** Read-only commands a bot may run over SSH. */
export const ALLOWED_COMMANDS: ReadonlySet<string> = new Set([
  "list-conversations",
  "find-conversation",
  "find-one-on-one",
  "get-messages",
  "search-messages",
  "get-members",
  "get-transcript",
  "whoami",
]);

/** Options that would bypass the cached session or write local files. */
const FORBIDDEN_OPTIONS: ReadonlySet<string> = new Set([
  "--token",
  "--bearer-token",
  "--substrate-token",
  "--login",
  "--auto",
  "--headless",
  "--debug",
  "--debug-port",
  "--output",
  "--print-token",
]);

/** Variables that must not reach the CLI from the SSH session. */
const REMOVED_VARIABLES = [
  "TEAMS_TOKEN",
  "TEAMS_BEARER_TOKEN",
  "TEAMS_SUBSTRATE_TOKEN",
  "TEAMS_TELEMETRY",
  "TEAMS_TELEMETRY_PATH",
  "DEBUG",
  "NODE_OPTIONS",
];

/**
 * Split a command line into arguments like a POSIX shell would for quoting
 * (single quotes, double quotes, backslashes), but with no expansion of any
 * kind: `$`, backticks, `;`, `|`, `*` are ordinary characters.
 */
export function splitCommandLine(commandLine: string): string[] {
  const parts: string[] = [];
  let current = "";
  let inArgument = false;
  let quote: "'" | '"' | null = null;

  for (let index = 0; index < commandLine.length; index++) {
    const character = commandLine[index];

    if (quote === "'") {
      if (character === "'") quote = null;
      else current += character;
      continue;
    }

    if (quote === '"') {
      if (character === '"') {
        quote = null;
      } else if (
        character === "\\" &&
        index + 1 < commandLine.length &&
        ['"', "\\", "$", "`"].includes(commandLine[index + 1])
      ) {
        current += commandLine[++index];
      } else {
        current += character;
      }
      continue;
    }

    if (character === "'" || character === '"') {
      quote = character;
      inArgument = true;
    } else if (character === "\\") {
      if (index + 1 < commandLine.length) {
        current += commandLine[++index];
        inArgument = true;
      }
    } else if (/\s/.test(character)) {
      if (inArgument) {
        parts.push(current);
        current = "";
        inArgument = false;
      }
    } else {
      current += character;
      inArgument = true;
    }
  }

  if (quote) {
    throw new InvalidArgumentError("Unterminated quote in the command.");
  }
  if (inArgument) {
    parts.push(current);
  }
  return parts;
}

/** Validate the SSH command and return the arguments for the CLI. */
export function buildCliArguments(commandLine: string | undefined): string[] {
  const allowed = [...ALLOWED_COMMANDS].join(", ");
  const parts = splitCommandLine(commandLine ?? "");
  if (parts[0] === "teams-api") {
    parts.shift();
  }

  const [command, ...rest] = parts;
  if (!command) {
    throw new InvalidArgumentError(
      `No command given. Allowed commands: ${allowed}.`,
    );
  }
  if (!ALLOWED_COMMANDS.has(command)) {
    throw new InvalidArgumentError(
      `"${command}" is not allowed over SSH. Allowed commands: ${allowed}.`,
    );
  }
  for (const part of rest) {
    const option = part.split("=", 1)[0];
    if (FORBIDDEN_OPTIONS.has(option)) {
      throw new InvalidArgumentError(`"${option}" is not allowed over SSH.`);
    }
  }
  return [command, ...rest];
}

/** The environment for the CLI: read-only, non-interactive, file token cache. */
export function buildCliEnvironment(
  base: NodeJS.ProcessEnv,
): NodeJS.ProcessEnv {
  const environment: NodeJS.ProcessEnv = {
    ...base,
    TEAMS_READ_ONLY: "true",
    TEAMS_NON_INTERACTIVE: "true",
    TEAMS_TOKEN_STORE: "file",
  };
  for (const name of REMOVED_VARIABLES) {
    delete environment[name];
  }
  return environment;
}

function main(): void {
  let cliArguments: string[];
  try {
    cliArguments = buildCliArguments(process.env.SSH_ORIGINAL_COMMAND);
  } catch (error) {
    process.stderr.write(`teams-api-ssh: ${safeErrorMessage(error)}\n`);
    process.exit(2);
  }

  const result = spawnSync(
    process.execPath,
    [join(__dirname, "cli.js"), ...cliArguments],
    { stdio: "inherit", env: buildCliEnvironment(process.env) },
  );
  if (result.error) {
    process.stderr.write(
      `teams-api-ssh: could not start teams-api (${safeErrorMessage(result.error)})\n`,
    );
    process.exit(1);
  }
  process.exit(result.status ?? 1);
}

const isMainModule =
  typeof require !== "undefined" &&
  typeof module !== "undefined" &&
  require.main === module;

if (isMainModule) {
  main();
}
