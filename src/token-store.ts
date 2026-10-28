/**
 * Token persistence across the platform credential store and a private file cache.
 *
 * Tokens are base64-encoded JSON with an acquisition timestamp.
 * Expired tokens (older than TOKEN_LIFETIME) are automatically cleared.
 *
 * `TEAMS_TOKEN_STORE` selects where tokens live:
 *   - `auto` (default): the platform store (macOS Keychain, Windows Credential
 *     Manager, Linux secret-tool), falling back to the private file cache in
 *     ~/.teams-api/tokens when the platform store refuses the write — for
 *     example the locked login Keychain in an SSH session. Reads check both
 *     and use the newest valid token.
 *   - `system` (alias `keychain`): the platform store only.
 *   - `file`: the private file cache only.
 *
 * Saving never throws. When no store accepts the token, a warning goes to
 * stderr and the caller keeps using the in-memory token.
 */

import type { TeamsToken } from "./types.js";
import type { CredentialStore } from "./credential-store.js";
import {
  createCredentialStore,
  createFileCredentialStore,
} from "./credential-store.js";
import {
  redactSecrets,
  registerSecret,
  registerTokenSecrets,
  safeErrorMessage,
} from "./redact.js";
import { resolveTokenStoreMode } from "./runtime-flags.js";

const TOKEN_LIFETIME = 23 * 60 * 60 * 1_000; // 23 hours (tokens last ~24h, 1h safety margin)
const JWT_EXPIRY_MARGIN = 5 * 60 * 1_000;

/** Cache key used when no email is known (see `TeamsClient.connect`). */
export const DEFAULT_TOKEN_ACCOUNT = "_default";

interface StoredToken {
  skypeToken: string;
  region: string;
  chatBearerToken?: string;
  bearerToken?: string;
  substrateToken?: string;
  amsToken?: string;
  sharePointToken?: string;
  sharePointHost?: string;
  acquiredAt: number;
}

let platformStore: CredentialStore | null = null;
let fileStore: CredentialStore | null = null;
/** Stores that refused a write in this process; skipped while a fallback exists. */
const unavailableStores = new Set<CredentialStore>();
const emittedWarnings = new Set<string>();

function getTokenStores(): CredentialStore[] {
  platformStore ??= createCredentialStore();
  fileStore ??= createFileCredentialStore();
  switch (resolveTokenStoreMode()) {
    case "file":
      return [fileStore];
    case "system":
      return [platformStore];
    case "auto":
      return [platformStore, fileStore];
  }
}

function allTokenStores(): CredentialStore[] {
  platformStore ??= createCredentialStore();
  fileStore ??= createFileCredentialStore();
  return [platformStore, fileStore];
}

function warnOnce(message: string): void {
  const safeMessage = redactSecrets(message);
  if (emittedWarnings.has(safeMessage)) return;
  emittedWarnings.add(safeMessage);
  process.stderr.write(`teams-api: ${safeMessage}\n`);
}

/** Forget cached stores and warnings. Intended for tests. */
export function resetTokenStoreState(): void {
  platformStore = null;
  fileStore = null;
  unavailableStores.clear();
  emittedWarnings.clear();
}

/** The `exp` claim of a JWT in epoch milliseconds, or undefined if it has none. */
function jwtExpiry(token: string | undefined): number | undefined {
  if (!token) return undefined;
  try {
    const payload = JSON.parse(
      Buffer.from(token.split(".")[1] ?? "", "base64url").toString("utf-8"),
    ) as { exp?: unknown };
    return typeof payload.exp === "number" ? payload.exp * 1_000 : undefined;
  } catch {
    return undefined;
  }
}

/**
 * When a token bundle stops working for chat, in epoch milliseconds: the
 * earliest `exp` of the Skype or chat Bearer token, capped at TOKEN_LIFETIME
 * after `acquiredAt`. Teams can hand back a token it issued hours earlier, so
 * the time of capture alone overstates how long it lasts.
 *
 * The short-lived search token is deliberately left out; search refreshes it
 * on demand.
 */
export function getSessionExpiry(
  token: Pick<TeamsToken, "skypeToken" | "chatBearerToken">,
  acquiredAt?: number,
): number | undefined {
  const candidates = [
    jwtExpiry(token.skypeToken),
    jwtExpiry(token.chatBearerToken),
    acquiredAt === undefined ? undefined : acquiredAt + TOKEN_LIFETIME,
  ].filter((value): value is number => value !== undefined);
  return candidates.length > 0 ? Math.min(...candidates) : undefined;
}

function decodeStoredToken(encoded: string): StoredToken | null {
  try {
    const decoded = JSON.parse(
      Buffer.from(encoded, "base64").toString("utf-8"),
    ) as StoredToken;
    return typeof decoded === "object" && decoded !== null ? decoded : null;
  } catch {
    return null;
  }
}

function isUsable(storedToken: StoredToken): boolean {
  const expiresAt = getSessionExpiry(storedToken, storedToken.acquiredAt);
  return expiresAt === undefined || expiresAt - Date.now() >= JWT_EXPIRY_MARGIN;
}

function toTeamsToken(storedToken: StoredToken): TeamsToken {
  return {
    skypeToken: storedToken.skypeToken,
    region: storedToken.region,
    chatBearerToken: storedToken.chatBearerToken,
    bearerToken: storedToken.bearerToken,
    substrateToken: storedToken.substrateToken,
    amsToken: storedToken.amsToken,
    sharePointToken: storedToken.sharePointToken,
    sharePointHost: storedToken.sharePointHost,
  };
}

/**
 * Cache a token. Never throws: see the module comment.
 *
 * @returns whether any store accepted the token.
 */
export async function saveToken(
  email: string,
  token: TeamsToken,
): Promise<boolean> {
  const storedToken: StoredToken = {
    skypeToken: token.skypeToken,
    region: token.region,
    chatBearerToken: token.chatBearerToken,
    bearerToken: token.bearerToken,
    substrateToken: token.substrateToken,
    amsToken: token.amsToken,
    sharePointToken: token.sharePointToken,
    sharePointHost: token.sharePointHost,
    acquiredAt: Date.now(),
  };
  const encoded = Buffer.from(JSON.stringify(storedToken)).toString("base64");
  registerTokenSecrets(token);
  registerSecret(encoded);

  const stores = getTokenStores();
  const failures: string[] = [];
  for (const store of stores) {
    if (stores.length > 1 && unavailableStores.has(store)) continue;
    try {
      await store.save(email, encoded);
      if (failures.length > 0) {
        warnOnce(
          `could not save the Teams token to the ${failures.join("; ")}; saved it to the ${store.label} instead.`,
        );
      }
      return true;
    } catch (error) {
      if (stores.length > 1) {
        unavailableStores.add(store);
      }
      failures.push(`${store.label} (${safeErrorMessage(error)})`);
    }
  }

  warnOnce(
    `could not cache the Teams token${failures.length > 0 ? ` in the ${failures.join("; ")}` : ""}. Continuing with the token in memory only.`,
  );
  return false;
}

export async function loadToken(email: string): Promise<TeamsToken | null> {
  let newest: StoredToken | null = null;

  for (const store of getTokenStores()) {
    let encoded: string | null;
    try {
      encoded = await store.load(email);
    } catch (error) {
      warnOnce(safeErrorMessage(error));
      continue;
    }
    if (!encoded) continue;
    registerSecret(encoded);

    const storedToken = decodeStoredToken(encoded);
    if (!storedToken || !isUsable(storedToken)) {
      await clearFromStore(store, email);
      continue;
    }
    if (!newest || storedToken.acquiredAt > newest.acquiredAt) {
      newest = storedToken;
    }
  }

  if (!newest) {
    return null;
  }

  const token = toTeamsToken(newest);
  registerTokenSecrets(token);
  return token;
}

/**
 * The newest cached token for the account, even if it has expired, without
 * clearing anything. Used to decide when to refresh the sign-in.
 */
export async function inspectCachedToken(
  email: string,
): Promise<{ token: TeamsToken; acquiredAt: number } | null> {
  let newest: StoredToken | null = null;
  for (const store of getTokenStores()) {
    let encoded: string | null;
    try {
      encoded = await store.load(email);
    } catch {
      continue;
    }
    if (!encoded) continue;
    registerSecret(encoded);
    const storedToken = decodeStoredToken(encoded);
    if (
      storedToken &&
      (!newest || storedToken.acquiredAt > newest.acquiredAt)
    ) {
      newest = storedToken;
    }
  }
  if (!newest) {
    return null;
  }
  const token = toTeamsToken(newest);
  registerTokenSecrets(token);
  return { token, acquiredAt: newest.acquiredAt };
}

/** When the cached sign-in stops working, or null if nothing is cached. */
export async function getCachedSessionExpiry(
  email: string,
): Promise<number | null> {
  const cached = await inspectCachedToken(email);
  if (!cached) {
    return null;
  }
  return getSessionExpiry(cached.token, cached.acquiredAt) ?? null;
}

async function clearFromStore(
  store: CredentialStore,
  email: string,
): Promise<void> {
  try {
    await store.clear(email);
  } catch {
    // Clearing is best effort.
  }
}

/** Remove the cached token from every store, whatever TEAMS_TOKEN_STORE says. */
export async function clearToken(email: string): Promise<void> {
  for (const store of allTokenStores()) {
    await clearFromStore(store, email);
  }
}
