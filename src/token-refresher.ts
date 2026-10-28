/**
 * Keeps the cached Teams sign-in fresh by refreshing it shortly before it
 * expires, so callers rarely meet an expired token.
 *
 * Used by `teams-api-mcp --http`, which runs all the time in the user's
 * session (where the browser can open). Each refresh loads Teams in the saved
 * browser profile for a few seconds.
 *
 * Teams sometimes hands back a token it issued earlier, so a refresh before
 * expiry may not extend the sign-in. The refresher then tries once more right
 * after expiry. When a refresh fails (for example because Microsoft wants the
 * user to sign in again), it waits RETRY_AFTER_FAILURE_MS instead of opening
 * the browser over and over.
 */

import { safeErrorMessage } from "./redact.js";

export const DEFAULT_REFRESH_MARGIN_MS = 45 * 60 * 1_000;
export const DEFAULT_CHECK_INTERVAL_MS = 5 * 60 * 1_000;
export const RETRY_AFTER_FAILURE_MS = 2 * 60 * 60 * 1_000;

export interface RefreshDecisionInput {
  now: number;
  /** When the cached sign-in stops working (epoch milliseconds). */
  expiresAt: number;
  marginMs: number;
  /** When the previous refresh attempt started, if any. */
  lastAttemptAt: number | undefined;
  lastAttemptFailed: boolean;
}

/**
 * Whether to refresh now. At most one attempt while the sign-in is about to
 * expire, one more right after it expires, and then one every
 * RETRY_AFTER_FAILURE_MS. After a failure, always wait RETRY_AFTER_FAILURE_MS.
 */
export function isRefreshDue({
  now,
  expiresAt,
  marginMs,
  lastAttemptAt,
  lastAttemptFailed,
}: RefreshDecisionInput): boolean {
  if (expiresAt - now > marginMs) {
    return false;
  }
  if (lastAttemptAt === undefined) {
    return true;
  }
  const sinceLastAttempt = now - lastAttemptAt;
  if (lastAttemptFailed) {
    return sinceLastAttempt >= RETRY_AFTER_FAILURE_MS;
  }
  if (now < expiresAt) {
    return lastAttemptAt < expiresAt - marginMs;
  }
  return (
    lastAttemptAt < expiresAt || sinceLastAttempt >= RETRY_AFTER_FAILURE_MS
  );
}

export interface TokenRefresherOptions {
  /** When the cached sign-in stops working, or null if nobody has signed in yet. */
  getExpiry: () => Promise<number | null>;
  /** Refresh the sign-in and save it to the cache. */
  refresh: () => Promise<void>;
  log: (message: string) => void;
  marginMs?: number;
  intervalMs?: number;
  now?: () => number;
}

export interface TokenRefresher {
  /** Resolves when the check made at startup has finished. */
  ready: Promise<void>;
  /** Run one check now (also runs on a timer). */
  check: () => Promise<void>;
  stop: () => void;
}

/** Check the sign-in now and then every `intervalMs`, refreshing when due. */
export function startTokenRefresher(
  options: TokenRefresherOptions,
): TokenRefresher {
  const marginMs = options.marginMs ?? DEFAULT_REFRESH_MARGIN_MS;
  const intervalMs = options.intervalMs ?? DEFAULT_CHECK_INTERVAL_MS;
  const now = options.now ?? Date.now;
  let lastAttemptAt: number | undefined;
  let lastAttemptFailed = false;
  let checking = false;

  async function check(): Promise<void> {
    if (checking) {
      return;
    }
    checking = true;
    try {
      const expiresAt = await options.getExpiry();
      if (expiresAt === null) {
        return;
      }
      const startedAt = now();
      if (
        !isRefreshDue({
          now: startedAt,
          expiresAt,
          marginMs,
          lastAttemptAt,
          lastAttemptFailed,
        })
      ) {
        return;
      }

      lastAttemptAt = startedAt;
      const expiry = new Date(expiresAt).toISOString();
      options.log(
        expiresAt > startedAt
          ? `Teams sign-in expires at ${expiry}; refreshing it now.`
          : `Teams sign-in expired at ${expiry}; refreshing it now.`,
      );
      try {
        await options.refresh();
        lastAttemptFailed = false;
      } catch (error) {
        lastAttemptFailed = true;
        options.log(
          `Refreshing the Teams sign-in failed: ${safeErrorMessage(error)}. Trying again in 2 hours.`,
        );
        return;
      }

      const renewedAt = await options.getExpiry();
      options.log(
        renewedAt !== null && renewedAt > expiresAt
          ? `Teams sign-in refreshed; valid until ${new Date(renewedAt).toISOString()}.`
          : `Teams kept the current sign-in; trying again after ${expiry}.`,
      );
    } catch (error) {
      options.log(
        `Checking the Teams sign-in failed: ${safeErrorMessage(error)}`,
      );
    } finally {
      checking = false;
    }
  }

  const timer = setInterval(() => void check(), intervalMs);
  timer.unref?.();

  return {
    ready: check(),
    check,
    stop: () => clearInterval(timer),
  };
}
