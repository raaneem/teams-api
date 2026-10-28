#!/usr/bin/env npx tsx
/**
 * MCP Server for Teams API
 *
 * Thin adapter that maps unified action definitions to MCP tools.
 * All tools, parameters, descriptions, and execution logic come
 * from `src/actions.ts` — the single source of truth.
 *
 * Configuration:
 *   Set environment variables for authentication:
 *     TEAMS_TOKEN           — Use an existing skype token
 *     TEAMS_BEARER_TOKEN    — Optional middle-tier bearer token for profile resolution
 *     TEAMS_SUBSTRATE_TOKEN — Optional Substrate bearer token for people/chat search
 *     TEAMS_REGION          — API region (required with TEAMS_TOKEN, optional otherwise)
 *     TEAMS_EMAIL           — Corporate email (optional; the server prompts the AI agent if needed)
 *     TEAMS_AUTO            — Set to "true" to use auto-login (macOS + FIDO2)
 *     TEAMS_HEADLESS        — Set to "true" to run FIDO2 auto-login headlessly
 *     TEAMS_LOGIN           — Set to "true" to use interactive browser login (all platforms)
 *     TEAMS_DEBUG_PORT      — Chrome debug port (default: 9222)
 *     TEAMS_READ_ONLY       — Set to "true" to expose only tools that read Teams
 *     TEAMS_TOKEN_STORE     — auto (default), system, or file (see src/token-store.ts)
 *     TEAMS_NON_INTERACTIVE — Set to "true" to fail fast instead of waiting for a person to sign in
 *     TEAMS_MCP_HTTP_TOKEN  — Bearer token for --http (default: generated once and stored)
 *     TEAMS_TELEMETRY       — Set to "true" to enable full debug telemetry (contributor use)
 *     TEAMS_IMAGE_DESCRIPTION_API_KEY — Optional vision API key for teams_describe_image
 *
 * Options:
 *   --http               Serve Streamable HTTP on http://<host>:<port>/mcp instead of stdio.
 *                        Read-only unless --allow-writes; every request needs the bearer token.
 *   --host <address>     HTTP listen address (default: 127.0.0.1)
 *   --port <port>        HTTP port (default: 8787)
 *   --allow-writes       Allow write tools over HTTP (TEAMS_READ_ONLY still wins)
 *   --public-url <url>   With --http, base URL clients use to reach this server, for links to
 *                        downloaded files (default: http://host.docker.internal:<port>)
 *   --refresh-margin <d> With --http, refresh the cached sign-in this long before it
 *                        expires (default: 45m)
 *   --no-auto-refresh    With --http, do not refresh the sign-in in the background
 *   --print-http-token   Print the HTTP bearer token (created on first use) and exit
 *
 * Usage in VS Code settings (mcp config):
 *   {
 *     "mcpServers": {
 *       "teams": {
 *         "command": "npx",
 *         "args": ["-y", "teams-api"],
 *         "env": {
 *           "TEAMS_LOGIN": "true"
 *         }
 *       }
 *     }
 *   }
 */

import { createRequire } from "node:module";
import { format, parseArgs } from "node:util";
import { McpServer } from "@modelcontextprotocol/sdk/server/mcp.js";
import { StdioServerTransport } from "@modelcontextprotocol/sdk/server/stdio.js";
import type { ContentBlock } from "@modelcontextprotocol/sdk/types.js";
import { z } from "zod";
import { TeamsClient } from "./teams-client.js";
import { actions } from "./actions/definitions.js";
import { formatOutput } from "./actions/formatters.js";
import type { ActionParameter, OutputFormat } from "./actions/formatters.js";
import type { DownloadResult } from "./actions/file-actions.js";
import { serverInstructions } from "./server-instructions.js";
import { recordToolCall, recordToolError } from "./telemetry.js";
import {
  buildDownloadContentBlocks,
  createDownloadOutputResults,
} from "./model-context-protocol-download-content.js";
import { McpAuthManager } from "./mcp-auth.js";
import type { AuthLogFunction } from "./types.js";
import { ReadOnlyModeError } from "./read-only.js";
import { isReadOnlyMode } from "./runtime-flags.js";
import { parseDurationMs } from "./since.js";
import {
  DEFAULT_TOKEN_ACCOUNT,
  getCachedSessionExpiry,
  getSessionExpiry,
} from "./token-store.js";
import {
  DEFAULT_REFRESH_MARGIN_MS,
  startTokenRefresher,
} from "./token-refresher.js";
import { InvalidArgumentError } from "./errors.js";
import {
  DEFAULT_HTTP_HOST,
  DEFAULT_HTTP_PORT,
  FileShare,
  MCP_HTTP_PATH,
  loadOrCreateHttpToken,
  startMcpHttpServer,
} from "./mcp-http.js";
import {
  redactKnownSecrets,
  redactSecrets,
  registerSecret,
  safeErrorMessage,
} from "./redact.js";

const { version } = createRequire(__filename)("../package.json") as {
  version: string;
};

const READ_ONLY_INSTRUCTIONS =
  "\nThis server is running in read-only mode (TEAMS_READ_ONLY). Only tools that read Teams are available; " +
  "sending, editing, deleting, reacting, and image description are disabled.\n";

function parameterToZod(parameter: ActionParameter): z.ZodTypeAny {
  let schema: z.ZodTypeAny;
  switch (parameter.type) {
    case "string":
      schema = z.string();
      break;
    case "number":
      schema = z.number();
      break;
    case "boolean":
      schema = z.boolean();
      break;
    case "string[]":
      schema = z.array(z.string());
      break;
  }
  schema = schema.describe(parameter.description);
  if (!parameter.required) {
    schema = schema.optional();
  }
  return schema;
}

function formatLogMessage(messages: unknown[]): string {
  return redactSecrets(
    messages
      .map((message) =>
        message instanceof Error ? message.message : String(message),
      )
      .join(" "),
  );
}

function createMcpAuthLogFunction(mcpServer: McpServer): AuthLogFunction {
  return (...messages: unknown[]) => {
    const data = formatLogMessage(messages);
    void mcpServer
      .sendLoggingMessage({
        level: "info",
        logger: "teams-api.auth",
        data,
      })
      .catch((error: unknown) => {
        if (process.env.TEAMS_TELEMETRY === "true") {
          console.error(
            "Failed to send MCP auth logging message:",
            safeErrorMessage(error),
          );
        }
      });
  };
}

/** Supplies an authenticated client to tool calls. */
export interface TeamsClientProvider {
  getClient(email?: string): Promise<TeamsClient>;
  authenticateOnStartup?(): Promise<void>;
}

export interface CreateMcpServerOptions {
  /** Defaults to an `McpAuthManager` configured from `TEAMS_*` variables. */
  authManager?: TeamsClientProvider;
  /** Defaults to `TEAMS_READ_ONLY`. */
  readOnly?: boolean;
  /** HTTP mode: gives downloaded files a short-lived link. */
  fileShare?: FileShare;
}

/** Build the MCP server with one tool per action. */
export function createMcpServer(options: CreateMcpServerOptions = {}): {
  server: McpServer;
  authManager: TeamsClientProvider;
} {
  const readOnly = options.readOnly ?? isReadOnlyMode();
  const server = new McpServer(
    {
      name: "teams-api",
      version,
    },
    {
      instructions: readOnly
        ? `${serverInstructions}${READ_ONLY_INSTRUCTIONS}`
        : serverInstructions,
    },
  );

  const authManager =
    options.authManager ??
    new McpAuthManager<TeamsClient>({
      clientFactory: TeamsClient,
      log: createMcpAuthLogFunction(server),
    });

  // ── Register all actions as MCP tools ───────────────────────────────

  for (const action of actions) {
    // In read-only mode the agent never sees tools that could change anything.
    if (readOnly && action.access !== "read") {
      continue;
    }

    const toolName = `teams_${action.name.replace(/-/g, "_")}`;

    const inputSchema: Record<string, z.ZodTypeAny> = {};
    for (const parameter of action.parameters) {
      inputSchema[parameter.name] = parameterToZod(parameter);
    }
    inputSchema["format"] = z
      .enum(["concise", "detailed", "json"])
      .describe(
        "Output format. " +
          '"concise" (default): light Markdown with actionable IDs and key decision fields; nested collections may be summarized. ' +
          '"detailed": full JSON for programmatic processing or inspecting exact field values. ' +
          '"json": stable, versioned JSON envelope with message IDs, links, and timestamps.',
      )
      .optional();
    inputSchema["email"] = z
      .string()
      .describe(
        "Corporate email address for Teams login. " +
          "Only needed if the server asks for it.",
      )
      .optional();

    server.registerTool(
      toolName,
      {
        title: action.title,
        description: action.description,
        inputSchema,
      },
      async (parameters) => {
        const outputFormat = (parameters.format as OutputFormat) ?? "concise";
        const start = Date.now();

        try {
          if (action.access !== "read" && isReadOnlyMode()) {
            throw new ReadOnlyModeError(`run "${action.name}"`);
          }
          action.validateParameters?.(parameters as Record<string, unknown>);
          const client = await authManager.getClient(
            parameters.email as string | undefined,
          );
          const result = await action.execute(
            client,
            parameters as Record<string, unknown>,
          );

          const outputResult =
            action.name === "download-file" && Array.isArray(result)
              ? createDownloadOutputResults(
                  result as DownloadResult[],
                  options.fileShare,
                )
              : result;
          const output = redactKnownSecrets(
            formatOutput(action, outputResult, outputFormat),
          );
          const durationMs = Date.now() - start;

          recordToolCall({
            tool: action.name,
            format: outputFormat,
            parameters: parameters as Record<string, unknown>,
            result: outputResult,
            output,
            durationMs,
          });

          // Build content blocks — file downloads get inline file content
          const contentBlocks: ContentBlock[] = [
            { type: "text" as const, text: output },
          ];

          if (action.name === "download-file" && Array.isArray(result)) {
            const downloads = result as DownloadResult[];
            contentBlocks.push(...buildDownloadContentBlocks(downloads));
          }

          return {
            content: contentBlocks,
          };
        } catch (error) {
          recordToolError({
            tool: action.name,
            format: outputFormat,
            parameters: parameters as Record<string, unknown>,
            error,
            durationMs: Date.now() - start,
          });
          // Return the redacted message ourselves: the SDK would otherwise
          // send `error.message` to the client verbatim.
          return {
            content: [{ type: "text" as const, text: safeErrorMessage(error) }],
            isError: true,
          };
        }
      },
    );
  }

  return { server, authManager };
}

export interface ServerArguments {
  http: boolean;
  host: string;
  port: number;
  allowWrites: boolean;
  printHttpToken: boolean;
  autoRefresh: boolean;
  refreshMarginMs: number;
  /** Base URL clients use to reach this server, for file links. */
  publicUrl: string;
}

/** Parse `teams-api-mcp` options. Throws `InvalidArgumentError` on bad input. */
export function parseServerArguments(argv: string[]): ServerArguments {
  let values;
  try {
    ({ values } = parseArgs({
      args: argv,
      options: {
        http: { type: "boolean", default: false },
        host: { type: "string", default: DEFAULT_HTTP_HOST },
        port: { type: "string", default: String(DEFAULT_HTTP_PORT) },
        "allow-writes": { type: "boolean", default: false },
        "print-http-token": { type: "boolean", default: false },
        "public-url": { type: "string" },
        "refresh-margin": { type: "string" },
        "no-auto-refresh": { type: "boolean", default: false },
      },
      strict: true,
      allowPositionals: false,
    }));
  } catch (error) {
    throw new InvalidArgumentError(safeErrorMessage(error));
  }

  const port = Number(values.port);
  if (!Number.isInteger(port) || port < 1 || port > 65_535) {
    throw new InvalidArgumentError(`Invalid --port "${values.port}"`);
  }
  return {
    http: values.http,
    host: values.host,
    port,
    allowWrites: values["allow-writes"],
    printHttpToken: values["print-http-token"],
    autoRefresh: !values["no-auto-refresh"],
    publicUrl: values["public-url"] ?? `http://host.docker.internal:${port}`,
    refreshMarginMs:
      values["refresh-margin"] === undefined
        ? DEFAULT_REFRESH_MARGIN_MS
        : parseDurationMs(values["refresh-margin"]),
  };
}

function logToStderr(...messages: unknown[]): void {
  process.stderr.write(`${redactSecrets(format(...messages))}\n`);
}

async function serveHttp(options: ServerArguments): Promise<void> {
  // Anything that can reach the port could otherwise send messages.
  if (!options.allowWrites) {
    process.env.TEAMS_READ_ONLY = "true";
  }
  const token = await loadOrCreateHttpToken();
  registerSecret(token);

  const readOnly = isReadOnlyMode();
  const authManager = new McpAuthManager<TeamsClient>({
    clientFactory: TeamsClient,
    log: logToStderr,
  });
  const fileShare = new FileShare(options.publicUrl);
  await startMcpHttpServer({
    host: options.host,
    port: options.port,
    token,
    createMcpServer: () =>
      createMcpServer({ authManager, readOnly, fileShare }).server,
    fileShare,
    log: logToStderr,
  });
  logToStderr(
    `teams-api MCP server listening on http://${options.host}:${options.port}${MCP_HTTP_PATH} ` +
      `(${readOnly ? "read-only" : "read and write"}; bearer token required)`,
  );

  if (options.autoRefresh) {
    startTokenRefresher({
      marginMs: options.refreshMarginMs,
      getExpiry: () => getCachedSessionExpiry(DEFAULT_TOKEN_ACCOUNT),
      refresh: async () => {
        const client = await authManager.getClient();
        // Signing in just now already produced a fresh token.
        const expiresAt = getSessionExpiry(client.getToken());
        if (
          expiresAt !== undefined &&
          expiresAt - Date.now() > options.refreshMarginMs
        ) {
          return;
        }
        await client.refreshTokenNow();
      },
      log: logToStderr,
    });
    logToStderr(
      `Keeping the Teams sign-in fresh: refreshing ${Math.round(options.refreshMarginMs / 60_000)} minutes before it expires.`,
    );
  }
}

async function main() {
  const options = parseServerArguments(process.argv.slice(2));

  if (options.printHttpToken) {
    // The HTTP access token, which the user pastes into their MCP client.
    process.stdout.write(`${await loadOrCreateHttpToken()}\n`);
    return;
  }

  registerSecret(process.env.TEAMS_TOKEN);
  registerSecret(process.env.TEAMS_BEARER_TOKEN);
  registerSecret(process.env.TEAMS_SUBSTRATE_TOKEN);

  if (options.http) {
    await serveHttp(options);
    return;
  }

  const { server, authManager } = createMcpServer();
  const transport = new StdioServerTransport();
  await server.connect(transport);
  void authManager.authenticateOnStartup?.();
}

function exitWithRedactedError(error: unknown): void {
  console.error("MCP server error:", safeErrorMessage(error));
  process.exit(error instanceof InvalidArgumentError ? 2 : 1);
}

const isMainModule =
  typeof require !== "undefined" &&
  typeof module !== "undefined" &&
  require.main === module;

if (isMainModule) {
  process.on("uncaughtException", exitWithRedactedError);
  process.on("unhandledRejection", exitWithRedactedError);
  main().catch(exitWithRedactedError);
}
