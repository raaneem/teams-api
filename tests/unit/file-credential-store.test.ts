/**
 * Unit tests for the private token file cache (src/credential-store.ts).
 * Uses a temporary directory; the real ~/.teams-api is never touched.
 */

import {
  chmodSync,
  mkdirSync,
  mkdtempSync,
  readdirSync,
  rmSync,
  statSync,
  symlinkSync,
  writeFileSync,
} from "node:fs";
import { tmpdir } from "node:os";
import { join } from "node:path";
import { describe, it, expect, beforeEach, afterEach } from "vitest";
import {
  CredentialStoreError,
  FileCredentialStore,
} from "../../src/credential-store.js";

const account = "user@contoso.com";
const secret = "eyJza3lwZVRva2VuIjoiQ0FOQVJZLUZJTEUtU0VDUkVUIn0=";

let root: string;
let directory: string;

beforeEach(() => {
  root = mkdtempSync(join(tmpdir(), "teams-api-file-store-"));
  directory = join(root, "tokens");
});

afterEach(() => {
  rmSync(root, { recursive: true, force: true });
});

function tokenFiles(): string[] {
  return readdirSync(directory).filter((name) => name.endsWith(".token"));
}

describe("FileCredentialStore", () => {
  it("saves and loads a token", async () => {
    const store = new FileCredentialStore(directory);

    await store.save(account, secret);

    await expect(store.load(account)).resolves.toBe(secret);
  });

  it("returns null when nothing is cached", async () => {
    await expect(
      new FileCredentialStore(directory).load(account),
    ).resolves.toBeNull();
  });

  it("removes the token on clear", async () => {
    const store = new FileCredentialStore(directory);
    await store.save(account, secret);

    await store.clear(account);

    await expect(store.load(account)).resolves.toBeNull();
    expect(tokenFiles()).toEqual([]);
  });

  it("names files by a hash, not the email", async () => {
    await new FileCredentialStore(directory).save(account, secret);

    const [fileName] = tokenFiles();
    expect(fileName).toMatch(/^[0-9a-f]{64}\.token$/);
  });

  it("leaves no temporary files behind", async () => {
    const store = new FileCredentialStore(directory);

    await store.save(account, secret);
    await store.save(account, `${secret}-updated`);

    expect(readdirSync(directory)).toHaveLength(1);
    await expect(store.load(account)).resolves.toBe(`${secret}-updated`);
  });

  it("describes itself without the full home path", () => {
    expect(new FileCredentialStore(directory).label).toContain("token file in");
  });

  describe.skipIf(process.platform === "win32")("permissions", () => {
    it("creates a 0700 directory and a 0600 file", async () => {
      await new FileCredentialStore(directory).save(account, secret);

      expect(statSync(directory).mode & 0o777).toBe(0o700);
      const [fileName] = tokenFiles();
      expect(statSync(join(directory, fileName)).mode & 0o777).toBe(0o600);
    });

    it("tightens an existing directory that others can read", async () => {
      mkdirSync(directory, { mode: 0o755 });
      chmodSync(directory, 0o755);

      await new FileCredentialStore(directory).save(account, secret);

      expect(statSync(directory).mode & 0o777).toBe(0o700);
    });

    it("refuses a token file that other users can read", async () => {
      const store = new FileCredentialStore(directory);
      await store.save(account, secret);
      const [fileName] = tokenFiles();
      chmodSync(join(directory, fileName), 0o644);

      const error = await store
        .load(account)
        .catch((caught: unknown) => caught);

      expect(error).toBeInstanceOf(CredentialStoreError);
      expect((error as Error).message).toContain("chmod 600");
      expect((error as Error).message).not.toContain(secret);
    });

    it("refuses a token path that is a symlink", async () => {
      const store = new FileCredentialStore(directory);
      await store.save(account, secret);
      const [fileName] = tokenFiles();
      const decoy = join(root, "decoy");
      writeFileSync(decoy, "decoy", { mode: 0o600 });
      rmSync(join(directory, fileName));
      symlinkSync(decoy, join(directory, fileName));

      await expect(store.load(account)).rejects.toBeInstanceOf(
        CredentialStoreError,
      );
    });
  });

  it("reports write failures without the token", async () => {
    // A regular file where the directory should be makes mkdir fail.
    writeFileSync(directory, "not a directory");

    const error = await new FileCredentialStore(directory)
      .save(account, secret)
      .catch((caught: unknown) => caught);

    expect(error).toBeInstanceOf(CredentialStoreError);
    expect((error as Error).message).not.toContain(secret);
  });
});
