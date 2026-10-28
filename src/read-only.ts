/**
 * Read-only mode — refuses every request that could change Teams.
 *
 * Enabled with `TEAMS_READ_ONLY=true` (or the CLI `--read-only` flag).
 * Enforced in three layers so one missed check cannot let a write through:
 *   - actions: the CLI and MCP server refuse actions whose access is not "read"
 *   - client: TeamsClient write methods throw before doing any work
 *   - network: fetchWithRetry refuses every request except GET/HEAD and a
 *     short allowlist of read-only POST endpoints (profile lookup and search)
 */

import { isReadOnlyMode } from "./runtime-flags.js";

export class ReadOnlyModeError extends Error {
  constructor(operation: string) {
    super(
      `teams-api is in read-only mode (TEAMS_READ_ONLY), so it refused to ${operation}.`,
    );
    this.name = "ReadOnlyModeError";
  }
}

/** Throw `ReadOnlyModeError` when read-only mode is on. */
export function assertWriteAllowed(operation: string): void {
  if (isReadOnlyMode()) {
    throw new ReadOnlyModeError(operation);
  }
}

/** POST endpoints that only read data, so they stay available in read-only mode. */
const READ_ONLY_POST_ENDPOINTS: ReadonlyArray<{
  hostname: string;
  pathname: RegExp;
}> = [
  // Profile lookup (src/api/middle-tier.ts)
  {
    hostname: "teams.cloud.microsoft",
    pathname: /^\/api\/mt\/[^/]+\/beta\/users\/fetchShortProfile$/,
  },
  // People and chat search (src/api/substrate.ts)
  {
    hostname: "substrate.office.com",
    pathname: /^\/search\/api\/v1\/suggestions$/,
  },
  // Message search (src/api/message-search.ts)
  {
    hostname: "substrate.office.com",
    pathname: /^\/searchservice\/api\/v2\/query$/,
  },
];

/** Whether a request may be sent while read-only mode is on. */
export function isRequestAllowedInReadOnlyMode(
  url: string,
  method: string,
): boolean {
  const normalizedMethod = method.toUpperCase();
  if (normalizedMethod === "GET" || normalizedMethod === "HEAD") {
    return true;
  }
  if (normalizedMethod !== "POST") {
    return false;
  }

  let parsed: URL;
  try {
    parsed = new URL(url);
  } catch {
    return false;
  }
  if (parsed.protocol !== "https:") {
    return false;
  }
  const hostname = parsed.hostname.toLowerCase();
  return READ_ONLY_POST_ENDPOINTS.some(
    (endpoint) =>
      endpoint.hostname === hostname && endpoint.pathname.test(parsed.pathname),
  );
}

/** Throw `ReadOnlyModeError` when read-only mode is on and the request could write. */
export function assertRequestAllowed(url: string, method: string): void {
  if (!isReadOnlyMode() || isRequestAllowedInReadOnlyMode(url, method)) {
    return;
  }
  let hostname = "an unknown host";
  try {
    hostname = new URL(url).hostname;
  } catch {
    // Keep the generic description for malformed URLs.
  }
  throw new ReadOnlyModeError(
    `send a ${method.toUpperCase()} request to ${hostname}`,
  );
}
