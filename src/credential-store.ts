/**
 * Cross-platform credential storage.
 *
 * Provides a unified interface for securely storing and retrieving
 * credentials across macOS, Windows, and Linux.
 *
 * | Platform | Mechanism                                              |
 * |----------|--------------------------------------------------------|
 * | macOS    | System Keychain via `security` CLI                     |
 * | Windows  | Windows Credential Manager via keytar (native wincred) |
 * | Linux    | `secret-tool` (libsecret) if available, else file with |
 * |          | 0o600 perms at ~/.config/teams-api/                    |
 *
 * `FileCredentialStore` is a private file cache (~/.teams-api/tokens) used
 * when the platform store is unavailable, such as the locked login Keychain
 * in an SSH session. See `src/token-store.ts` for how the stores combine.
 *
 * Errors thrown from this module never contain the stored data.
 */

import { execFileSync } from "node:child_process";
import {
  chmodSync,
  closeSync,
  existsSync,
  fsyncSync,
  lstatSync,
  mkdirSync,
  openSync,
  readFileSync,
  renameSync,
  rmSync,
  writeFileSync,
  unlinkSync,
  type Stats,
} from "node:fs";
import { homedir } from "node:os";
import { join } from "node:path";
import { createHash, randomBytes } from "node:crypto";
import { redactSecrets } from "./redact.js";

export interface CredentialStore {
  /** Human-readable name used in warnings (e.g. "macOS Keychain"). */
  readonly label: string;
  save(account: string, data: string): Promise<void>;
  load(account: string): Promise<string | null>;
  clear(account: string): Promise<void>;
}

/** A credential store operation failed. The message never contains stored data. */
export class CredentialStoreError extends Error {
  constructor(message: string) {
    super(message);
    this.name = "CredentialStoreError";
  }
}

// ── macOS: System Keychain ────────────────────────────────────────────

const CREDENTIAL_SERVICE = "teams-api";

/**
 * Summarize a failed `security` invocation from its stderr only.
 *
 * Node's own error message for a failed `execFileSync` repeats the full
 * command line, which includes the `-w <data>` argument, so it must never be
 * surfaced. `security` writes lines such as
 * "security: SecKeychainItemCreateFromContent (<default>): User interaction is not allowed."
 */
function describeSecurityFailure(error: unknown): string {
  const failure = error as { stderr?: unknown; status?: unknown };
  const stderr =
    typeof failure.stderr === "string"
      ? failure.stderr
      : Buffer.isBuffer(failure.stderr)
        ? failure.stderr.toString("utf-8")
        : "";
  const firstLine = stderr
    .split("\n")
    .map((line) => line.trim())
    .find((line) => line.length > 0);
  if (firstLine) {
    const detail = firstLine
      .replace(/^security:\s*/, "")
      .replace(/^[A-Za-z]+\s*(?:\([^)]*\))?:\s*/, "");
    return redactSecrets(detail).slice(0, 200);
  }
  return typeof failure.status === "number"
    ? `security exited with status ${failure.status}`
    : "security could not be run";
}

class KeychainStore implements CredentialStore {
  readonly label = "macOS Keychain";

  async save(account: string, data: string): Promise<void> {
    try {
      execFileSync(
        "security",
        [
          "add-generic-password",
          "-a",
          account,
          "-s",
          CREDENTIAL_SERVICE,
          "-w",
          data,
          "-U",
        ],
        { stdio: ["ignore", "pipe", "pipe"] },
      );
    } catch (error) {
      throw new CredentialStoreError(describeSecurityFailure(error));
    }
  }

  async load(account: string): Promise<string | null> {
    try {
      return execFileSync(
        "security",
        [
          "find-generic-password",
          "-a",
          account,
          "-s",
          CREDENTIAL_SERVICE,
          "-w",
        ],
        { encoding: "utf-8", stdio: ["ignore", "pipe", "pipe"] },
      ).trim();
    } catch {
      // Missing item or locked keychain (e.g. in an SSH session).
      return null;
    }
  }

  async clear(account: string): Promise<void> {
    try {
      execFileSync(
        "security",
        ["delete-generic-password", "-a", account, "-s", CREDENTIAL_SERVICE],
        { stdio: ["ignore", "pipe", "pipe"] },
      );
    } catch {
      // Entry may not exist, that's fine
    }
  }
}

// ── Windows: Windows Credential Manager via keytar ───────────────────
//
// The previous implementation spawned PowerShell with an inline script that
// loaded System.Security.Cryptography.ProtectedData to call DPAPI.
// That pattern (PowerShell + Add-Type + Protect/Unprotect on binary data) is
// a textbook ransomware behavioral signature and is reliably flagged by
// Windows Defender. keytar calls the native wincred API directly from C++,
// with no PowerShell or inline scripting involved.

function getKeytar(): typeof import("keytar") {
  try {
    // eslint-disable-next-line @typescript-eslint/no-require-imports
    return require("keytar") as typeof import("keytar");
  } catch {
    throw new Error(
      "keytar is required on Windows for credential storage. " +
        "Install it with: npm install keytar",
    );
  }
}

function accountKey(account: string): string {
  return createHash("sha256").update(account).digest("hex");
}

function accountToFileName(account: string): string {
  return createHash("sha256").update(account).digest("hex") + ".dat";
}

// Windows Credential Manager limits each credential blob to 2560 bytes.
// Teams tokens easily exceed this (~4500+ bytes for 5 JWTs).  We chunk
// large payloads across multiple credential entries.
const WINCRED_MAX_BLOB = 2560;
const CHUNK_MARKER = "CHUNKED:";

class WinCredStore implements CredentialStore {
  readonly label = "Windows Credential Manager";

  async save(account: string, data: string): Promise<void> {
    const key = accountKey(account);
    const kt = getKeytar();

    // Clear any previous chunks before saving
    await this.clearChunks(key, kt);

    if (data.length <= WINCRED_MAX_BLOB) {
      await kt.setPassword(CREDENTIAL_SERVICE, key, data);
      return;
    }

    const chunks: string[] = [];
    for (let i = 0; i < data.length; i += WINCRED_MAX_BLOB) {
      chunks.push(data.slice(i, i + WINCRED_MAX_BLOB));
    }

    // Primary entry stores the chunk count
    await kt.setPassword(
      CREDENTIAL_SERVICE,
      key,
      `${CHUNK_MARKER}${chunks.length}`,
    );
    for (let i = 0; i < chunks.length; i++) {
      await kt.setPassword(CREDENTIAL_SERVICE, `${key}:${i}`, chunks[i]);
    }
  }

  async load(account: string): Promise<string | null> {
    const key = accountKey(account);
    const kt = getKeytar();

    const value = await kt.getPassword(CREDENTIAL_SERVICE, key);
    if (!value) return null;

    if (!value.startsWith(CHUNK_MARKER)) {
      return value; // Single-entry credential (backward compatible)
    }

    const chunkCount = parseInt(value.slice(CHUNK_MARKER.length), 10);
    const parts: string[] = [];
    for (let i = 0; i < chunkCount; i++) {
      const chunk = await kt.getPassword(CREDENTIAL_SERVICE, `${key}:${i}`);
      if (chunk === null) return null; // Corrupted — missing chunk
      parts.push(chunk);
    }
    return parts.join("");
  }

  async clear(account: string): Promise<void> {
    const key = accountKey(account);
    const kt = getKeytar();
    await this.clearChunks(key, kt);
    try {
      await kt.deletePassword(CREDENTIAL_SERVICE, key);
    } catch {
      // Entry may not exist
    }
  }

  private async clearChunks(
    key: string,
    kt: typeof import("keytar"),
  ): Promise<void> {
    try {
      const value = await kt.getPassword(CREDENTIAL_SERVICE, key);
      if (value?.startsWith(CHUNK_MARKER)) {
        const count = parseInt(value.slice(CHUNK_MARKER.length), 10);
        for (let i = 0; i < count; i++) {
          try {
            await kt.deletePassword(CREDENTIAL_SERVICE, `${key}:${i}`);
          } catch {
            // Chunk may not exist
          }
        }
      }
    } catch {
      // Primary entry may not exist
    }
  }
}

// ── Linux: secret-tool or file-based fallback ────────────────────────

function getLinuxStorePath(): string {
  const configDir =
    process.env.XDG_CONFIG_HOME ??
    join(process.env.HOME ?? homedir(), ".config");
  return join(configDir, "teams-api");
}

function hasSecretTool(): boolean {
  try {
    execFileSync("which", ["secret-tool"], { stdio: "pipe" });
    return true;
  } catch {
    return false;
  }
}

class LinuxStore implements CredentialStore {
  private useSecretTool: boolean;

  constructor() {
    this.useSecretTool = hasSecretTool();
  }

  get label(): string {
    return this.useSecretTool
      ? "Linux secret service"
      : "token file in ~/.config/teams-api";
  }

  private getFilePath(account: string): string {
    const dir = getLinuxStorePath();
    return join(dir, accountToFileName(account));
  }

  private ensureDir(): void {
    const dir = getLinuxStorePath();
    if (!existsSync(dir)) {
      mkdirSync(dir, { recursive: true, mode: 0o700 });
    }
  }

  async save(account: string, data: string): Promise<void> {
    if (this.useSecretTool) {
      try {
        execFileSync(
          "secret-tool",
          [
            "store",
            "--label",
            "teams-api",
            "service",
            CREDENTIAL_SERVICE,
            "account",
            account,
          ],
          { input: data, encoding: "utf-8" },
        );
        return;
      } catch {
        // Fall through to file-based storage
        this.useSecretTool = false;
      }
    }

    this.ensureDir();
    writeFileSync(this.getFilePath(account), data, {
      encoding: "utf-8",
      mode: 0o600,
    });
  }

  async load(account: string): Promise<string | null> {
    if (this.useSecretTool) {
      try {
        return execFileSync(
          "secret-tool",
          ["lookup", "service", CREDENTIAL_SERVICE, "account", account],
          { encoding: "utf-8" },
        ).trim();
      } catch {
        // Secret not found or secret-tool failed
        return null;
      }
    }

    const filePath = this.getFilePath(account);
    if (!existsSync(filePath)) {
      return null;
    }

    try {
      return readFileSync(filePath, { encoding: "utf-8" }).trim();
    } catch {
      return null;
    }
  }

  async clear(account: string): Promise<void> {
    if (this.useSecretTool) {
      try {
        execFileSync("secret-tool", [
          "clear",
          "service",
          CREDENTIAL_SERVICE,
          "account",
          account,
        ]);
      } catch {
        // Entry may not exist
      }
    }

    // Also clean up the file-based fallback
    const filePath = this.getFilePath(account);
    try {
      unlinkSync(filePath);
    } catch {
      // File may not exist
    }
  }
}

// ── Private file cache (all platforms) ───────────────────────────────
//
// One file per account in ~/.teams-api/tokens. The directory is 0700 and each
// file is 0600, written atomically (exclusive create, fsync, rename). Paths
// that are symlinks, owned by another user, or readable by other users are
// refused. This matches how the browser profile in ~/.teams-api is protected.

/** Directory of the private token file cache. */
export function getTokenCacheDirectory(): string {
  return join(homedir(), ".teams-api", "tokens");
}

function describeFileError(error: unknown): string {
  const code = (error as { code?: unknown })?.code;
  return typeof code === "string" ? code : "unknown error";
}

function isOwnedByCurrentUser(stats: Stats): boolean {
  return typeof process.getuid !== "function" || stats.uid === process.getuid();
}

function hasPrivatePermissions(stats: Stats): boolean {
  return process.platform === "win32" || (stats.mode & 0o077) === 0;
}

export class FileCredentialStore implements CredentialStore {
  /** Omit the directory to follow the current home directory on every call. */
  constructor(private readonly directoryOverride?: string) {}

  get label(): string {
    const home = homedir();
    const displayDirectory = this.directory.startsWith(home)
      ? `~${this.directory.slice(home.length)}`
      : this.directory;
    return `token file in ${displayDirectory}`;
  }

  private get directory(): string {
    return this.directoryOverride ?? getTokenCacheDirectory();
  }

  async save(account: string, data: string): Promise<void> {
    this.ensurePrivateDirectory();
    const filePath = this.getFilePath(account);
    const temporaryPath = `${filePath}.${process.pid}.${randomBytes(6).toString("hex")}.tmp`;
    try {
      const descriptor = openSync(temporaryPath, "wx", 0o600);
      try {
        writeFileSync(descriptor, data, "utf-8");
        fsyncSync(descriptor);
      } finally {
        closeSync(descriptor);
      }
      renameSync(temporaryPath, filePath);
    } catch (error) {
      rmSync(temporaryPath, { force: true });
      throw new CredentialStoreError(
        `could not write the token cache file (${describeFileError(error)})`,
      );
    }
  }

  async load(account: string): Promise<string | null> {
    const filePath = this.getFilePath(account);
    let stats: Stats;
    try {
      stats = lstatSync(filePath);
    } catch {
      return null;
    }
    if (!stats.isFile() || !isOwnedByCurrentUser(stats)) {
      throw new CredentialStoreError(
        `ignored ${filePath} because it is not a regular file owned by you`,
      );
    }
    if (!hasPrivatePermissions(stats)) {
      throw new CredentialStoreError(
        `ignored ${filePath} because other users can access it (fix with: chmod 600 "${filePath}")`,
      );
    }
    try {
      const data = readFileSync(filePath, "utf-8").trim();
      return data.length > 0 ? data : null;
    } catch {
      return null;
    }
  }

  async clear(account: string): Promise<void> {
    try {
      rmSync(this.getFilePath(account), { force: true });
    } catch {
      // Nothing to remove
    }
  }

  private getFilePath(account: string): string {
    return join(this.directory, `${accountKey(account)}.token`);
  }

  private ensurePrivateDirectory(): void {
    let stats: Stats;
    try {
      mkdirSync(this.directory, { recursive: true, mode: 0o700 });
      stats = lstatSync(this.directory);
    } catch (error) {
      throw new CredentialStoreError(
        `could not create the token cache directory (${describeFileError(error)})`,
      );
    }
    if (!stats.isDirectory() || !isOwnedByCurrentUser(stats)) {
      throw new CredentialStoreError(
        `${this.directory} is not a directory owned by you`,
      );
    }
    if (!hasPrivatePermissions(stats)) {
      chmodSync(this.directory, 0o700);
    }
  }
}

// ── Factory ──────────────────────────────────────────────────────────

/** The platform credential store (Keychain, Credential Manager, or secret-tool). */
export function createCredentialStore(): CredentialStore {
  switch (process.platform) {
    case "darwin":
      return new KeychainStore();
    case "win32":
      return new WinCredStore();
    default:
      return new LinuxStore();
  }
}

/** The private file cache in ~/.teams-api/tokens. */
export function createFileCredentialStore(): CredentialStore {
  return new FileCredentialStore();
}
