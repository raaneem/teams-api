#!/usr/bin/env npx tsx
/**
 * teams-api CLI
 *
 * Thin adapter that maps unified action definitions to CLI commands.
 * All commands, parameters, descriptions, and execution logic come
 * from `src/actions.ts` — the single source of truth.
 *
 * Special commands (auth, logout, guide) are defined here directly since
 * they handle authentication or documentation rather than Teams data operations.
 *
 * Everything written to stdout or stderr passes through `src/redact.ts`, and
 * failures map to stable exit codes (see EXIT_CODES).
 */

import { writeFileSync } from "node:fs";
import { createRequire } from "node:module";
import { resolve } from "node:path";
import { format } from "node:util";
import { Command, CommanderError } from "commander";
import { TeamsClient } from "./teams-client.js";
import { actions } from "./actions/definitions.js";
import { formatJsonError, formatOutput } from "./actions/formatters.js";
import type {
  ActionDefinition,
  ActionParameter,
  OutputFormat,
} from "./actions/formatters.js";
import { serverInstructions } from "./server-instructions.js";
import {
  ApiAuthError,
  ApiRateLimitError,
  ApiResponseError,
} from "./api/common.js";
import { AuthRequiredError, InvalidArgumentError } from "./errors.js";
import { ReadOnlyModeError } from "./read-only.js";
import { isReadOnlyMode } from "./runtime-flags.js";
import {
  redactKnownSecrets,
  redactSecrets,
  registerSecret,
  safeErrorMessage,
} from "./redact.js";
import { describeTokenStatus } from "./token-status.js";
import type {
  AutoLoginOptions,
  InteractiveLoginOptions,
  ManualTokenOptions,
} from "./types.js";

const VALID_FORMATS: OutputFormat[] = ["concise", "detailed", "json"];

/** Process exit codes. Documented in README ("Unattended use"). */
export const EXIT_CODES = {
  error: 1,
  invalidArgument: 2,
  authentication: 3,
  readOnly: 4,
  rateLimited: 5,
} as const;

const { version } = createRequire(__filename)("../package.json") as {
  version: string;
};

// ── Shared auth options ────────────────────────────────────────────────

interface AuthFlags {
  auto?: boolean;
  headless?: boolean;
  login?: boolean;
  debug?: boolean;
  email?: string;
  token?: string;
  bearerToken?: string;
  substrateToken?: string;
  debugPort: string;
  region?: string;
  readOnly?: boolean;
}

function addAuthOptions(command: Command): Command {
  return command
    .option("--auto", "Auto-acquire token via FIDO2 passkey (macOS)")
    .option("--headless", "Run FIDO2 auto-login without a visible browser")
    .option(
      "--login",
      "Interactive browser login (all platforms, no FIDO2 needed)",
    )
    .option("--debug", "Use Chrome debug session for token capture")
    .option(
      "--email <email>",
      "Corporate email (required with --auto, optional with --login)",
    )
    .option("--token <token>", "Use an existing skype token directly")
    .option(
      "--bearer-token <token>",
      "Optional middle-tier bearer token for profile/member resolution",
    )
    .option(
      "--substrate-token <token>",
      "Optional Substrate bearer token for people/chat search",
    )
    .option(
      "--debug-port <port>",
      "Chrome debug port for manual token capture",
      "9222",
    )
    .option(
      "--region <region>",
      "API region (optional for login/debug auth, required with --token)",
    )
    .option(
      "--read-only",
      "Refuse every command that changes Teams (same as TEAMS_READ_ONLY=true)",
    );
}

/** Progress and diagnostics go to stderr, redacted. */
function logToStderr(...messages: unknown[]): void {
  process.stderr.write(`${redactSecrets(format(...messages))}\n`);
}

/** Write to stdout and wait until it is flushed (pipes are async on macOS). */
function writeStdout(text: string): Promise<void> {
  return new Promise((resolvePromise, rejectPromise) => {
    process.stdout.write(`${text}\n`, (error) =>
      error ? rejectPromise(error) : resolvePromise(),
    );
  });
}

async function createClient(flags: AuthFlags): Promise<TeamsClient> {
  registerSecret(flags.token);
  registerSecret(flags.bearerToken);
  registerSecret(flags.substrateToken);

  if (flags.token) {
    if (!flags.region) {
      throw new InvalidArgumentError("--region is required when using --token");
    }
    const client = TeamsClient.fromToken(
      flags.token,
      flags.region,
      flags.bearerToken,
      flags.substrateToken,
    );
    if (flags.email) {
      client.setEmail(flags.email);
    }
    return client;
  }

  if (flags.auto) {
    if (!flags.email) {
      throw new InvalidArgumentError("--email is required when using --auto");
    }
    const autoLoginOptions: AutoLoginOptions = {
      email: flags.email,
      region: flags.region,
      headless: flags.headless === true,
      verbose: true,
      log: logToStderr,
    };
    return TeamsClient.create(autoLoginOptions);
  }

  if (flags.login) {
    const interactiveLoginOptions: InteractiveLoginOptions = {
      region: flags.region,
      email: flags.email,
      verbose: true,
      log: logToStderr,
    };
    const client = await TeamsClient.fromInteractiveLogin(
      interactiveLoginOptions,
    );
    if (flags.email) {
      client.setEmail(flags.email);
    }
    return client;
  }

  if (flags.debug) {
    const manualOptions: ManualTokenOptions = {
      debugPort: Number(flags.debugPort),
      region: flags.region,
    };
    return TeamsClient.fromDebugSession(manualOptions);
  }

  // Default: cached token, else headed interactive login on every platform.
  const client = await TeamsClient.connect({
    email: flags.email,
    region: flags.region,
    verbose: true,
    log: logToStderr,
  });
  if (flags.email) {
    client.setEmail(flags.email);
  }
  return client;
}

// ── Errors ─────────────────────────────────────────────────────────────

interface CliErrorDetails {
  code: string;
  exitCode: number;
}

function classifyError(error: unknown): CliErrorDetails {
  if (error instanceof ReadOnlyModeError) {
    return { code: "READ_ONLY", exitCode: EXIT_CODES.readOnly };
  }
  if (error instanceof InvalidArgumentError) {
    return { code: "INVALID_ARGUMENT", exitCode: EXIT_CODES.invalidArgument };
  }
  if (error instanceof AuthRequiredError) {
    return { code: "AUTH_REQUIRED", exitCode: EXIT_CODES.authentication };
  }
  if (error instanceof ApiAuthError) {
    return { code: "AUTH_FAILED", exitCode: EXIT_CODES.authentication };
  }
  if (error instanceof ApiRateLimitError) {
    return { code: "RATE_LIMITED", exitCode: EXIT_CODES.rateLimited };
  }
  if (error instanceof ApiResponseError) {
    return { code: "UPSTREAM_ERROR", exitCode: EXIT_CODES.error };
  }
  return { code: "ERROR", exitCode: EXIT_CODES.error };
}

/** Print a redacted error (JSON envelope on stdout, or text on stderr) and return the exit code. */
async function reportError(
  error: unknown,
  options: { json: boolean; command: string },
): Promise<number> {
  const { code, exitCode } = classifyError(error);
  const message = safeErrorMessage(error);
  if (options.json) {
    try {
      await writeStdout(formatJsonError(options.command, code, message));
    } catch {
      // stdout is gone; nothing else to report to.
    }
  } else {
    process.stderr.write(`Error: ${message}\n`);
  }
  return exitCode;
}

// ── Helpers ────────────────────────────────────────────────────────────

function camelToKebab(name: string): string {
  return name.replace(/([A-Z])/g, "-$1").toLowerCase();
}

function coerceParameter(
  value: string | string[] | boolean | undefined,
  type: ActionParameter["type"],
): unknown {
  if (value === undefined) return undefined;
  switch (type) {
    case "number": {
      const parsed = Number(value);
      if (Number.isNaN(parsed)) {
        throw new InvalidArgumentError(`Invalid number: "${String(value)}"`);
      }
      return parsed;
    }
    case "boolean":
      return value === true || value === "true";
    case "string[]":
      return Array.isArray(value) ? value : [value];
    default:
      return value;
  }
}

function resolveOutputFormat(flags: Record<string, unknown>): OutputFormat {
  if (flags.json === true) {
    return "json";
  }
  const rawFormat = flags.format as string | undefined;
  if (rawFormat === undefined) {
    return "concise";
  }
  if (!VALID_FORMATS.includes(rawFormat as OutputFormat)) {
    throw new InvalidArgumentError(
      `Invalid format "${rawFormat}". Valid formats: ${VALID_FORMATS.join(", ")}`,
    );
  }
  return rawFormat as OutputFormat;
}

function applyReadOnlyFlag(flags: { readOnly?: boolean }): void {
  if (flags.readOnly) {
    process.env.TEAMS_READ_ONLY = "true";
  }
}

// ── Program ────────────────────────────────────────────────────────────

interface CliState {
  exitCode: number;
}

function registerActionCommand(
  program: Command,
  action: ActionDefinition,
  state: CliState,
): void {
  const command = program.command(action.name).description(action.description);

  addAuthOptions(command);

  for (const parameter of action.parameters) {
    const flag = camelToKebab(parameter.name);

    if (parameter.type === "string[]") {
      command.option(`--${flag} <value...>`, parameter.description);
    } else if (parameter.type === "boolean") {
      if (parameter.default === true) {
        // Default-true boolean: define --no-* to allow opting out.
        // Commander auto-sets the value to true when --no-* is absent.
        command.option(`--no-${flag}`, `Disable: ${parameter.description}`);
      } else {
        command.option(`--${flag}`, parameter.description);
      }
    } else if (parameter.required) {
      command.requiredOption(`--${flag} <value>`, parameter.description);
    } else {
      command.option(
        `--${flag} <value>`,
        parameter.description,
        parameter.default !== undefined ? String(parameter.default) : undefined,
      );
    }
  }

  command.option(
    "--format <format>",
    "Output format (concise, detailed, json)",
  );
  command.option("--json", "Same as --format json");
  command.option(
    "--output <file>",
    "Export output to file (default format: concise)",
  );

  command.action(async (flags: Record<string, unknown>) => {
    const json = flags.json === true || flags.format === "json";
    let progressShown = false;

    try {
      applyReadOnlyFlag(flags);
      const outputFormat = resolveOutputFormat(flags);
      if (action.access !== "read" && isReadOnlyMode()) {
        throw new ReadOnlyModeError(`run "${action.name}"`);
      }

      const actionParameters: Record<string, unknown> = {};
      for (const parameter of action.parameters) {
        actionParameters[parameter.name] = coerceParameter(
          flags[parameter.name] as string | string[] | boolean | undefined,
          parameter.type,
        );
      }

      action.validateParameters?.(actionParameters);

      // Progress for message fetching, only for people watching a terminal.
      if (action.name === "get-messages" && process.stderr.isTTY && !json) {
        actionParameters.onProgress = (count: number) => {
          progressShown = true;
          process.stderr.write(`\r  ${count} messages fetched...`);
        };
      }

      const client = await createClient(flags as unknown as AuthFlags);
      const result = await action.execute(client, actionParameters);
      if (progressShown) {
        process.stderr.write("\n");
        progressShown = false;
      }

      const output = redactKnownSecrets(
        formatOutput(action, result, outputFormat),
      );

      if (flags.output) {
        const outputPath = resolve(flags.output as string);
        writeFileSync(outputPath, output, "utf-8");
        if (json) {
          process.stderr.write(`Output written to ${outputPath}\n`);
        } else {
          await writeStdout(`Output written to ${outputPath}`);
        }
      } else {
        await writeStdout(output);
      }
    } catch (error) {
      if (progressShown) {
        process.stderr.write("\n");
      }
      state.exitCode = await reportError(error, {
        json,
        command: action.name,
      });
    }
  });
}

export function createProgram(state: CliState = { exitCode: 0 }): Command {
  const program = new Command();

  program
    .name("teams-api")
    .description("AI-native Microsoft Teams integration CLI")
    .version(version)
    .exitOverride();

  for (const action of actions) {
    registerActionCommand(program, action, state);
  }

  // ── auth command (special — creates a client, not a data operation) ───

  addAuthOptions(
    program
      .command("auth")
      .description(
        "Sign in if needed, cache the Teams tokens, and print which tokens are available " +
          "(never the tokens themselves unless --print-token is given)",
      ),
  )
    .option(
      "--print-token",
      "Print the raw token bundle to stdout. Treat it as a password. Refused in read-only mode.",
    )
    .option(
      "--refresh",
      "Refresh the cached sign-in now from the saved browser session (the browser opens briefly)",
    )
    .action(
      async (
        flags: AuthFlags & { printToken?: boolean; refresh?: boolean },
      ) => {
        try {
          applyReadOnlyFlag(flags);
          if (flags.printToken && isReadOnlyMode()) {
            throw new ReadOnlyModeError("print tokens");
          }
          const client = await createClient(flags);
          if (flags.refresh) {
            await client.refreshTokenNow();
          }
          const token = client.getToken();
          await writeStdout(
            flags.printToken
              ? JSON.stringify(token, null, 2)
              : JSON.stringify(describeTokenStatus(token), null, 2),
          );
        } catch (error) {
          state.exitCode = await reportError(error, {
            json: false,
            command: "auth",
          });
        }
      },
    );

  // ── logout command (special — clears cached token) ────────────────────

  program
    .command("logout")
    .description(
      "Clear the cached token from the credential store and file cache",
    )
    .requiredOption("--email <email>", "Email whose cached token to clear")
    .action(async (flags: { email: string }) => {
      try {
        await TeamsClient.clearCachedToken(flags.email);
        await writeStdout(`Cached token for ${flags.email} cleared.`);
      } catch (error) {
        state.exitCode = await reportError(error, {
          json: false,
          command: "logout",
        });
      }
    });

  // ── guide command (special — prints usage guidance) ───────────────────

  program
    .command("guide")
    .description(
      "Print workflow guidance, tips, and important notes for using the Teams API tools",
    )
    .action(async () => {
      await writeStdout(serverInstructions);
    });

  return program;
}

/** Run the CLI and return the process exit code. */
export async function runCli(argv: string[] = process.argv): Promise<number> {
  const state: CliState = { exitCode: 0 };
  const program = createProgram(state);
  try {
    await program.parseAsync(argv);
  } catch (error) {
    if (error instanceof CommanderError) {
      // Commander already printed usage errors, help, or the version.
      return error.exitCode;
    }
    process.stderr.write(`Fatal error: ${safeErrorMessage(error)}\n`);
    return EXIT_CODES.error;
  }
  return state.exitCode;
}

function exitWithRedactedError(error: unknown): void {
  process.stderr.write(`Fatal error: ${safeErrorMessage(error)}\n`);
  process.exit(EXIT_CODES.error);
}

const isMainModule =
  typeof require !== "undefined" &&
  typeof module !== "undefined" &&
  require.main === module;

if (isMainModule) {
  // Replace Node's default printers, which dump stacks and error properties.
  process.on("uncaughtException", exitWithRedactedError);
  process.on("unhandledRejection", exitWithRedactedError);

  // Exit explicitly: browser and emoji-map handles can keep the event loop alive.
  runCli().then((exitCode) => process.exit(exitCode), exitWithRedactedError);
}
