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
 *     TEAMS_TELEMETRY       — Set to "true" to enable full debug telemetry (contributor use)
 *     TEAMS_IMAGE_DESCRIPTION_API_KEY — Optional vision API key for teams_describe_image
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
              ? createDownloadOutputResults(result as DownloadResult[])
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

async function main() {
  registerSecret(process.env.TEAMS_TOKEN);
  registerSecret(process.env.TEAMS_BEARER_TOKEN);
  registerSecret(process.env.TEAMS_SUBSTRATE_TOKEN);

  const { server, authManager } = createMcpServer();
  const transport = new StdioServerTransport();
  await server.connect(transport);
  void authManager.authenticateOnStartup?.();
}

function exitWithRedactedError(error: unknown): void {
  console.error("MCP server error:", safeErrorMessage(error));
  process.exit(1);
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
