/**
 * Unit tests for the macOS Keychain store (src/credential-store.ts).
 *
 * `security` is mocked to fail the way it does in an SSH session, where the
 * login keychain is locked: the real Keychain is never touched.
 */

import { inspect } from "node:util";
import { describe, it, expect, vi, beforeEach, afterEach } from "vitest";

const execFileSync = vi.hoisted(() => vi.fn());

vi.mock("node:child_process", async (importOriginal) => ({
  ...(await importOriginal<typeof import("node:child_process")>()),
  execFileSync,
}));

import {
  createCredentialStore,
  CredentialStoreError,
} from "../../src/credential-store.js";

const originalPlatform = process.platform;
const lockedKeychainStderr =
  "security: SecKeychainItemCreateFromContent (<default>): User interaction is not allowed.\n";
const tokenBundle = Buffer.from(
  JSON.stringify({
    skypeToken: "CANARY-SKYPE-TOKEN-1234567890",
    region: "apac",
  }),
).toString("base64");

/** Throw the same error shape Node's execFileSync throws for a non-zero exit. */
function failLikeLockedKeychain(file: string, args: string[]): never {
  throw Object.assign(
    new Error(
      `Command failed: ${file} ${args.join(" ")}\n${lockedKeychainStderr}`,
    ),
    {
      status: 36,
      signal: null,
      pid: 4242,
      stdout: Buffer.from(""),
      stderr: Buffer.from(lockedKeychainStderr),
      output: [null, Buffer.from(""), Buffer.from(lockedKeychainStderr)],
    },
  );
}

beforeEach(() => {
  execFileSync.mockReset();
  Object.defineProperty(process, "platform", { value: "darwin" });
});

afterEach(() => {
  Object.defineProperty(process, "platform", { value: originalPlatform });
});

describe("KeychainStore", () => {
  it("reports a locked keychain without the token bundle", async () => {
    execFileSync.mockImplementation(failLikeLockedKeychain);
    const store = createCredentialStore();

    const error = await store
      .save("user@contoso.com", tokenBundle)
      .catch((caught: unknown) => caught);

    expect(error).toBeInstanceOf(CredentialStoreError);
    expect((error as Error).message).toBe("User interaction is not allowed.");
    expect(String((error as Error).stack)).not.toContain(tokenBundle);
    expect(inspect(error, { depth: 10 })).not.toContain(tokenBundle);
    for (const property of ["stdout", "stderr", "output", "cmd", "spawnargs"]) {
      expect(error).not.toHaveProperty(property);
    }
  });

  it("describes a failure with no stderr by exit status", async () => {
    execFileSync.mockImplementation(() => {
      throw Object.assign(new Error("Command failed: security ... -w secret"), {
        status: 51,
        stderr: Buffer.from(""),
      });
    });

    await expect(
      createCredentialStore().save("user@contoso.com", tokenBundle),
    ).rejects.toThrow("security exited with status 51");
  });

  it("never lets security write to the inherited stderr", async () => {
    execFileSync.mockReturnValue("");
    const store = createCredentialStore();

    await store.save("user@contoso.com", tokenBundle);
    await store.load("user@contoso.com");
    await store.clear("user@contoso.com");

    for (const call of execFileSync.mock.calls) {
      expect(call[2]).toMatchObject({ stdio: ["ignore", "pipe", "pipe"] });
    }
  });

  it("returns null when the keychain cannot be read", async () => {
    execFileSync.mockImplementation(failLikeLockedKeychain);

    await expect(
      createCredentialStore().load("user@contoso.com"),
    ).resolves.toBeNull();
  });

  it("ignores clear failures", async () => {
    execFileSync.mockImplementation(failLikeLockedKeychain);

    await expect(
      createCredentialStore().clear("user@contoso.com"),
    ).resolves.toBeUndefined();
  });
});
