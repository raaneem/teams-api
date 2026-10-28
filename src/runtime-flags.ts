/**
 * Environment-driven runtime switches shared by the CLI, MCP server, and SDK.
 *
 * Boolean switches fail closed: any value other than an explicit "off" value
 * ("", "0", "false", "no", "off") turns the switch on, so a typo such as
 * `TEAMS_READ_ONLY=yes` never silently leaves writes enabled.
 */

const OFF_VALUES = new Set(["", "0", "false", "no", "off"]);

/** Interpret an environment variable as an on/off switch. */
export function isFlagEnabled(value: string | undefined): boolean {
  return value !== undefined && !OFF_VALUES.has(value.trim().toLowerCase());
}

/** True when `TEAMS_READ_ONLY` is set: every request that changes Teams is refused. */
export function isReadOnlyMode(
  environment: NodeJS.ProcessEnv = process.env,
): boolean {
  return isFlagEnabled(environment.TEAMS_READ_ONLY);
}

/**
 * True when `TEAMS_NON_INTERACTIVE` is set: browser login may renew tokens
 * from an existing session but must not wait for a person to sign in.
 */
export function isNonInteractiveMode(
  environment: NodeJS.ProcessEnv = process.env,
): boolean {
  return isFlagEnabled(environment.TEAMS_NON_INTERACTIVE);
}

/**
 * Where cached tokens live.
 *
 * - `auto` (default): the platform store, falling back to the private file
 *   cache when the platform store refuses a write.
 * - `system`: the platform store only (`keychain` is accepted as an alias).
 * - `file`: the private file cache only.
 */
export type TokenStoreMode = "auto" | "system" | "file";

/** Resolve `TEAMS_TOKEN_STORE`. Unrecognized values fall back to `auto`. */
export function resolveTokenStoreMode(
  environment: NodeJS.ProcessEnv = process.env,
): TokenStoreMode {
  switch (environment.TEAMS_TOKEN_STORE?.trim().toLowerCase()) {
    case "file":
      return "file";
    case "system":
    case "keychain":
      return "system";
    default:
      return "auto";
  }
}
