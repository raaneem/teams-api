/**
 * Secret redaction for everything teams-api prints or logs.
 *
 * Two mechanisms work together:
 *   - Exact values: every token teams-api holds is registered here, and its
 *     value is replaced wherever it appears. This catches encodings that no
 *     pattern recognizes, such as the base64 token bundle written to the
 *     credential store.
 *   - Patterns: JWTs, Skype and Bearer authorization values, `security -w`
 *     arguments, OAuth URL parameters, and long base64 runs, for secrets that
 *     were never registered (for example inside third-party error messages).
 *
 * Patterns are applied to diagnostics only (errors and logs). Command output
 * goes through `redactKnownSecrets`, which never alters message content.
 */

import type { TeamsToken } from "./types.js";

const REDACTED = "[REDACTED]";

/** Shorter values are not registered, to avoid redacting ordinary words. */
const MINIMUM_SECRET_LENGTH = 16;

const knownSecrets = new Set<string>();

/** Register a secret value so it is redacted from all output. */
export function registerSecret(value: string | null | undefined): void {
  if (typeof value === "string" && value.length >= MINIMUM_SECRET_LENGTH) {
    knownSecrets.add(value);
  }
}

/** Register every credential in a token bundle (region and host are not secrets). */
export function registerTokenSecrets(
  token: Partial<TeamsToken> | null | undefined,
): void {
  if (!token) return;
  registerSecret(token.skypeToken);
  registerSecret(token.chatBearerToken);
  registerSecret(token.bearerToken);
  registerSecret(token.substrateToken);
  registerSecret(token.amsToken);
  registerSecret(token.sharePointToken);
}

/** Forget all registered secrets. Intended for tests. */
export function resetRegisteredSecrets(): void {
  knownSecrets.clear();
}

/** Replace registered secret values only. Safe for command output. */
export function redactKnownSecrets(text: string): string {
  if (knownSecrets.size === 0) return text;
  let result = text;
  // Longest first, so a bundle that contains a shorter token is replaced whole.
  const secrets = [...knownSecrets].sort((a, b) => b.length - a.length);
  for (const secret of secrets) {
    if (result.includes(secret)) {
      result = result.split(secret).join(REDACTED);
    }
  }
  return result;
}

// Characters that appear in JWTs, base64, and base64url values. Quotes and
// backslashes are excluded so redaction never breaks surrounding JSON.
const TOKEN_CHARACTERS = "[A-Za-z0-9._~+/=-]";

const SECRET_PATTERNS: ReadonlyArray<readonly [RegExp, string]> = [
  [/eyJ[A-Za-z0-9_-]{8,}\.[A-Za-z0-9_-]{8,}\.[A-Za-z0-9_-]*/g, REDACTED],
  [new RegExp(`(skypetoken=)${TOKEN_CHARACTERS}+`, "gi"), `$1${REDACTED}`],
  [new RegExp(`(skype_token\\s+)${TOKEN_CHARACTERS}+`, "gi"), `$1${REDACTED}`],
  [new RegExp(`(Bearer\\s+)${TOKEN_CHARACTERS}{20,}`, "gi"), `$1${REDACTED}`],
  [/((?:x-skypetoken|cookie|set-cookie)\s*[:=]\s*)[^\s"']+/gi, `$1${REDACTED}`],
  // macOS `security add-generic-password ... -w <secret>` command lines.
  [new RegExp(`(\\s-w\\s+)${TOKEN_CHARACTERS}+`, "g"), `$1${REDACTED}`],
  [
    /([?&#](?:code|access_token|id_token|refresh_token|client_info|session_state|state|nonce|login_hint|sid|tempauth|token)=)[^&#\s"']+/gi,
    `$1${REDACTED}`,
  ],
  [/[A-Za-z0-9+/_-]{200,}={0,2}/g, REDACTED],
];

/** Redact registered secrets and anything that looks like a credential. */
export function redactSecrets(text: string): string {
  let result = redactKnownSecrets(text);
  for (const [pattern, replacement] of SECRET_PATTERNS) {
    result = result.replace(pattern, replacement);
  }
  return result;
}

/**
 * A printable, redacted description of an error.
 *
 * Only the message is used: stacks and attached properties (such as the
 * `stdout`/`stderr` buffers on child-process errors) are never included.
 */
export function safeErrorMessage(error: unknown): string {
  const message =
    error instanceof Error
      ? error.message
      : typeof error === "string"
        ? error
        : "Unknown error";
  return redactSecrets(message);
}

/** Drop the query string and fragment, which can carry OAuth codes and state. */
export function redactUrl(url: string): string {
  try {
    const parsed = new URL(url);
    const suffix = parsed.search || parsed.hash ? `?${REDACTED}` : "";
    const base =
      parsed.origin === "null"
        ? `${parsed.protocol}${parsed.pathname}`
        : `${parsed.origin}${parsed.pathname}`;
    return `${base}${suffix}`;
  } catch {
    return "[unparseable URL]";
  }
}
