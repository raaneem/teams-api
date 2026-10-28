/**
 * Streamable HTTP transport for the MCP server (`teams-api-mcp --http`).
 *
 * For MCP clients that connect by URL instead of launching a process, such as
 * apps that run in Docker and reach this machine at `host.docker.internal`.
 *
 * - Every request needs `Authorization: Bearer <token>`. The token is generated
 *   once and kept in the private file cache (see `loadOrCreateHttpToken`).
 * - Requests with an `Origin` header (browsers) are refused, so web pages
 *   cannot drive the server even through DNS rebinding.
 * - The server is stateless: each POST gets its own MCP server instance, and
 *   all instances share one authenticated Teams client.
 */

import { createHash, randomBytes, timingSafeEqual } from "node:crypto";
import {
  createServer,
  type IncomingMessage,
  type Server,
  type ServerResponse,
} from "node:http";
import type { McpServer } from "@modelcontextprotocol/sdk/server/mcp.js";
import { StreamableHTTPServerTransport } from "@modelcontextprotocol/sdk/server/streamableHttp.js";
import {
  FileCredentialStore,
  type CredentialStore,
} from "./credential-store.js";
import { InvalidArgumentError } from "./errors.js";
import { safeErrorMessage } from "./redact.js";

/** Key of the HTTP access token in the private file cache. */
const HTTP_TOKEN_ACCOUNT = "teams-api-mcp-http-access-token";
const MINIMUM_TOKEN_LENGTH = 32;

export const DEFAULT_HTTP_HOST = "127.0.0.1";
export const DEFAULT_HTTP_PORT = 8787;
export const MCP_HTTP_PATH = "/mcp";

/**
 * The bearer token clients must send. `TEAMS_MCP_HTTP_TOKEN` overrides it;
 * otherwise it is created on first use and stored in ~/.teams-api/tokens.
 */
export async function loadOrCreateHttpToken(
  store: CredentialStore = new FileCredentialStore(),
  environment: NodeJS.ProcessEnv = process.env,
): Promise<string> {
  const override = environment.TEAMS_MCP_HTTP_TOKEN?.trim();
  if (override) {
    if (override.length < MINIMUM_TOKEN_LENGTH) {
      throw new InvalidArgumentError(
        `TEAMS_MCP_HTTP_TOKEN must be at least ${MINIMUM_TOKEN_LENGTH} characters.`,
      );
    }
    return override;
  }

  const existing = await store.load(HTTP_TOKEN_ACCOUNT);
  if (existing) {
    return existing;
  }
  const token = randomBytes(32).toString("base64url");
  await store.save(HTTP_TOKEN_ACCOUNT, token);
  return token;
}

export interface McpHttpServerOptions {
  host: string;
  port: number;
  token: string;
  /** Build a fresh MCP server for one request. */
  createMcpServer: () => McpServer;
  log?: (message: string) => void;
}

function sendError(
  response: ServerResponse,
  status: number,
  message: string,
): void {
  response.writeHead(status, { "Content-Type": "application/json" });
  response.end(
    JSON.stringify({
      jsonrpc: "2.0",
      error: { code: -32000, message },
      id: null,
    }),
  );
}

/** Start the HTTP server; resolves once it is listening. */
export function startMcpHttpServer(
  options: McpHttpServerOptions,
): Promise<Server> {
  const expectedDigest = createHash("sha256").update(options.token).digest();

  function isAuthorized(header: string | undefined): boolean {
    const match = /^Bearer\s+(\S+)\s*$/i.exec(header ?? "");
    if (!match) {
      return false;
    }
    const digest = createHash("sha256").update(match[1]).digest();
    return timingSafeEqual(digest, expectedDigest);
  }

  async function handleRequest(
    request: IncomingMessage,
    response: ServerResponse,
  ): Promise<void> {
    const { pathname } = new URL(request.url ?? "/", "http://localhost");
    if (pathname !== MCP_HTTP_PATH) {
      sendError(response, 404, "Not found");
      return;
    }
    if (request.headers.origin !== undefined) {
      sendError(response, 403, "Browser requests are not allowed");
      return;
    }
    if (!isAuthorized(request.headers.authorization)) {
      response.setHeader("WWW-Authenticate", 'Bearer realm="teams-api"');
      sendError(response, 401, "Unauthorized");
      return;
    }
    if (request.method !== "POST") {
      response.setHeader("Allow", "POST");
      sendError(response, 405, "Method not allowed");
      return;
    }

    const mcpServer = options.createMcpServer();
    const transport = new StreamableHTTPServerTransport({
      sessionIdGenerator: undefined,
      enableJsonResponse: true,
    });
    response.on("close", () => {
      void mcpServer.close();
    });
    await mcpServer.connect(transport);
    await transport.handleRequest(request, response);
  }

  const server = createServer((request, response) => {
    handleRequest(request, response).catch((error: unknown) => {
      options.log?.(`MCP HTTP request failed: ${safeErrorMessage(error)}`);
      if (response.headersSent) {
        response.end();
      } else {
        sendError(response, 500, "Internal server error");
      }
    });
  });

  return new Promise((resolvePromise, rejectPromise) => {
    server.once("error", rejectPromise);
    server.listen(options.port, options.host, () => {
      server.off("error", rejectPromise);
      resolvePromise(server);
    });
  });
}
