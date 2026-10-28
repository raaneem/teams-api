/**
 * Shared error types with stable meanings.
 *
 * The CLI maps them to exit codes and JSON error codes (see `src/cli.ts`):
 * `InvalidArgumentError` → 2 / `INVALID_ARGUMENT`,
 * `AuthRequiredError` → 3 / `AUTH_REQUIRED`.
 */

/** A command-line or API argument was invalid. */
export class InvalidArgumentError extends Error {
  constructor(message: string) {
    super(message);
    this.name = "InvalidArgumentError";
  }
}

/** Teams needs a person to sign in again before tokens can be captured. */
export class AuthRequiredError extends Error {
  constructor(message: string) {
    super(message);
    this.name = "AuthRequiredError";
  }
}
