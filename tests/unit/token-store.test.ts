/**
 * Unit tests for the token store (src/token-store.ts).
 *
 * These tests mock the credential stores to verify that the token store
 * correctly serializes/deserializes tokens, handles expiry, and falls back
 * to the private file cache without ever printing a token.
 */

import { describe, it, expect, vi, beforeEach, afterEach } from "vitest";
import type { TeamsToken } from "../../src/types.js";

const mockStore = vi.hoisted(() => ({
  label: "macOS Keychain",
  save: vi.fn(),
  load: vi.fn(),
  clear: vi.fn(),
}));

const mockFileStore = vi.hoisted(() => ({
  label: "token file in ~/.teams-api/tokens",
  save: vi.fn(),
  load: vi.fn(),
  clear: vi.fn(),
}));

vi.mock("../../src/credential-store.js", () => ({
  createCredentialStore: () => mockStore,
  createFileCredentialStore: () => mockFileStore,
}));

import {
  saveToken,
  loadToken,
  clearToken,
  getCachedSessionExpiry,
  getSessionExpiry,
  inspectCachedToken,
  resetTokenStoreState,
} from "../../src/token-store.js";

const testEmail = "user@company.com";
const testToken: TeamsToken = {
  skypeToken: "test-skype-token-abc123",
  region: "apac",
};

const testTokenWithAuxiliaryTokens: TeamsToken = {
  skypeToken: "test-skype-token-abc123",
  region: "apac",
  bearerToken: "test-bearer-token",
  substrateToken: "test-substrate-token",
};

let stderrWrite: { mock: { calls: unknown[][] }; mockRestore: () => void };

beforeEach(() => {
  vi.resetAllMocks();
  resetTokenStoreState();
  vi.stubEnv("TEAMS_TOKEN_STORE", "");
  stderrWrite = vi
    .spyOn(process.stderr, "write")
    .mockImplementation(() => true);
  vi.useFakeTimers();
  vi.setSystemTime(new Date("2026-03-17T12:00:00.000Z"));
});

afterEach(() => {
  vi.useRealTimers();
  vi.unstubAllEnvs();
  stderrWrite.mockRestore();
});

function stderrOutput(): string {
  return stderrWrite.mock.calls
    .map((call: unknown[]) => String(call[0]))
    .join("");
}

function encodeStoredToken(token: TeamsToken, acquiredAt: number): string {
  return Buffer.from(JSON.stringify({ ...token, acquiredAt })).toString(
    "base64",
  );
}

function makeStoredTokenBase64(
  overrides: { acquiredAt?: number } = {},
): string {
  const storedToken = {
    skypeToken: testToken.skypeToken,
    region: testToken.region,
    acquiredAt:
      overrides.acquiredAt ?? new Date("2026-03-17T12:00:00.000Z").getTime(),
  };
  return Buffer.from(JSON.stringify(storedToken)).toString("base64");
}

describe("saveToken", () => {
  it("should call credential store save with correct arguments", async () => {
    await saveToken(testEmail, testToken);

    expect(mockStore.save).toHaveBeenCalledTimes(1);
    expect(mockStore.save).toHaveBeenCalledWith(testEmail, expect.any(String));
  });

  it("should store base64-encoded JSON with acquiredAt timestamp", async () => {
    await saveToken(testEmail, testToken);

    const encodedValue = mockStore.save.mock.calls[0][1] as string;
    const decoded = JSON.parse(
      Buffer.from(encodedValue, "base64").toString("utf-8"),
    ) as {
      skypeToken: string;
      region: string;
      acquiredAt: number;
    };

    expect(decoded.skypeToken).toBe(testToken.skypeToken);
    expect(decoded.region).toBe(testToken.region);
    expect(decoded.acquiredAt).toBe(
      new Date("2026-03-17T12:00:00.000Z").getTime(),
    );
  });

  it("should persist optional bearer and substrate tokens", async () => {
    await saveToken(testEmail, testTokenWithAuxiliaryTokens);

    const encodedValue = mockStore.save.mock.calls[0][1] as string;
    const decoded = JSON.parse(
      Buffer.from(encodedValue, "base64").toString("utf-8"),
    ) as {
      skypeToken: string;
      region: string;
      bearerToken?: string;
      substrateToken?: string;
      acquiredAt: number;
    };

    expect(decoded.bearerToken).toBe("test-bearer-token");
    expect(decoded.substrateToken).toBe("test-substrate-token");
  });
});

describe("loadToken", () => {
  it("should return token when cached and not expired", async () => {
    const encoded = makeStoredTokenBase64();
    mockStore.load.mockResolvedValueOnce(encoded);

    const result = await loadToken(testEmail);

    expect(result).toEqual(testToken);
    expect(mockStore.load).toHaveBeenCalledWith(testEmail);
  });

  it("should restore optional bearer and substrate tokens", async () => {
    const encoded = Buffer.from(
      JSON.stringify({
        skypeToken: testTokenWithAuxiliaryTokens.skypeToken,
        region: testTokenWithAuxiliaryTokens.region,
        bearerToken: testTokenWithAuxiliaryTokens.bearerToken,
        substrateToken: testTokenWithAuxiliaryTokens.substrateToken,
        acquiredAt: new Date("2026-03-17T12:00:00.000Z").getTime(),
      }),
    ).toString("base64");
    mockStore.load.mockResolvedValueOnce(encoded);

    const result = await loadToken(testEmail);

    expect(result).toEqual(testTokenWithAuxiliaryTokens);
  });

  it("should return null when no token exists in credential store", async () => {
    mockStore.load.mockResolvedValueOnce(null);

    const result = await loadToken(testEmail);

    expect(result).toBeNull();
  });

  it("should return null and clear token when expired (older than 23 hours)", async () => {
    const twentyFourHoursAgo =
      new Date("2026-03-17T12:00:00.000Z").getTime() - 24 * 60 * 60 * 1_000;
    const encoded = makeStoredTokenBase64({ acquiredAt: twentyFourHoursAgo });
    mockStore.load.mockResolvedValueOnce(encoded);

    const result = await loadToken(testEmail);

    expect(result).toBeNull();
    expect(mockStore.clear).toHaveBeenCalledWith(testEmail);
  });

  it("should return token when acquired exactly 22 hours ago", async () => {
    const twentyTwoHoursAgo =
      new Date("2026-03-17T12:00:00.000Z").getTime() - 22 * 60 * 60 * 1_000;
    const encoded = makeStoredTokenBase64({
      acquiredAt: twentyTwoHoursAgo,
    });
    mockStore.load.mockResolvedValueOnce(encoded);

    const result = await loadToken(testEmail);

    expect(result).toEqual(testToken);
  });

  it("should return null and clear token when data is corrupted", async () => {
    mockStore.load.mockResolvedValueOnce("not-valid-base64!!!");

    const result = await loadToken(testEmail);

    expect(result).toBeNull();
    expect(mockStore.clear).toHaveBeenCalledWith(testEmail);
  });

  describe("chat Bearer token", () => {
    const now = new Date("2026-03-17T12:00:00.000Z").getTime();

    function makeJwt(expiresAtMs: number): string {
      const payload = Buffer.from(
        JSON.stringify({ exp: Math.floor(expiresAtMs / 1_000) }),
      ).toString("base64url");
      return `header.${payload}.signature`;
    }

    function encodeBearerOnlyToken(chatBearerToken: string): string {
      return Buffer.from(
        JSON.stringify({
          skypeToken: "",
          chatBearerToken,
          region: "apac",
          acquiredAt: now,
        }),
      ).toString("base64");
    }

    it("should restore a chat Bearer token that is still valid", async () => {
      const chatBearerToken = makeJwt(now + 60 * 60 * 1_000);
      mockStore.load.mockResolvedValueOnce(
        encodeBearerOnlyToken(chatBearerToken),
      );

      const result = await loadToken(testEmail);

      expect(result).toEqual({
        skypeToken: "",
        chatBearerToken,
        region: "apac",
      });
    });

    it("should clear a chat Bearer token that expires within the safety margin", async () => {
      mockStore.load.mockResolvedValueOnce(
        encodeBearerOnlyToken(makeJwt(now + 2 * 60 * 1_000)),
      );

      const result = await loadToken(testEmail);

      expect(result).toBeNull();
      expect(mockStore.clear).toHaveBeenCalledWith(testEmail);
    });

    it("should keep a chat Bearer token that is not a decodable JWT", async () => {
      mockStore.load.mockResolvedValueOnce(
        encodeBearerOnlyToken("opaque-token"),
      );

      const result = await loadToken(testEmail);

      expect(result?.chatBearerToken).toBe("opaque-token");
    });
  });
});

describe("fallback to the private file cache", () => {
  it("saves to the file cache when the Keychain refuses the write, without printing the token", async () => {
    // The raw Node error for a failed `security` call repeats the command
    // line, including `-w <token bundle>`: exactly what leaked over SSH.
    mockStore.save.mockImplementation(async (account: string, data: string) => {
      throw new Error(
        `Command failed: security add-generic-password -a ${account} -s teams-api -w ${data} -U\n` +
          "security: SecKeychainItemCreateFromContent (<default>): User interaction is not allowed.",
      );
    });
    mockFileStore.save.mockResolvedValue(undefined);

    await expect(
      saveToken(testEmail, testTokenWithAuxiliaryTokens),
    ).resolves.toBe(true);

    const savedData = mockFileStore.save.mock.calls[0][1] as string;
    expect(mockFileStore.save).toHaveBeenCalledWith(testEmail, savedData);
    const output = stderrOutput();
    expect(output).toContain("saved it to the token file");
    expect(output).toContain("User interaction is not allowed");
    expect(output).not.toContain(savedData);
    expect(output).not.toContain(testTokenWithAuxiliaryTokens.skypeToken);
    expect(output).not.toContain(testTokenWithAuxiliaryTokens.bearerToken);
    expect(output).not.toContain(testTokenWithAuxiliaryTokens.substrateToken);
  });

  it("stops retrying the Keychain after it refused once in this process", async () => {
    mockStore.save.mockRejectedValue(
      new Error("User interaction is not allowed."),
    );
    mockFileStore.save.mockResolvedValue(undefined);

    await saveToken(testEmail, testToken);
    await saveToken("_default", testToken);

    expect(mockStore.save).toHaveBeenCalledTimes(1);
    expect(mockFileStore.save).toHaveBeenCalledTimes(2);
    expect(stderrOutput().match(/saved it to the token file/g)).toHaveLength(1);
  });

  it("never throws when every store refuses the write", async () => {
    mockStore.save.mockRejectedValue(
      new Error("User interaction is not allowed."),
    );
    mockFileStore.save.mockRejectedValue(new Error("EACCES"));

    await expect(saveToken(testEmail, testToken)).resolves.toBe(false);
    expect(stderrOutput()).toContain(
      "Continuing with the token in memory only",
    );
  });

  it("uses only the file cache when TEAMS_TOKEN_STORE=file", async () => {
    vi.stubEnv("TEAMS_TOKEN_STORE", "file");
    mockFileStore.save.mockResolvedValue(undefined);
    mockFileStore.load.mockResolvedValue(
      encodeStoredToken(testToken, Date.now()),
    );

    await saveToken(testEmail, testToken);
    const result = await loadToken(testEmail);

    expect(result).toEqual(testToken);
    expect(mockStore.save).not.toHaveBeenCalled();
    expect(mockStore.load).not.toHaveBeenCalled();
  });

  it("uses only the platform store when TEAMS_TOKEN_STORE=system", async () => {
    vi.stubEnv("TEAMS_TOKEN_STORE", "system");
    mockStore.save.mockRejectedValue(
      new Error("User interaction is not allowed."),
    );

    await expect(saveToken(testEmail, testToken)).resolves.toBe(false);
    expect(mockFileStore.save).not.toHaveBeenCalled();
  });

  it("loads the newest valid token across both stores", async () => {
    const now = Date.now();
    const olderToken: TeamsToken = {
      skypeToken: "older-skype-token-value",
      region: "apac",
    };
    const newerToken: TeamsToken = {
      skypeToken: "newer-skype-token-value",
      region: "apac",
    };
    mockStore.load.mockResolvedValue(
      encodeStoredToken(olderToken, now - 2 * 60 * 60 * 1_000),
    );
    mockFileStore.load.mockResolvedValue(
      encodeStoredToken(newerToken, now - 60 * 60 * 1_000),
    );

    await expect(loadToken(testEmail)).resolves.toEqual(newerToken);
  });

  it("clears only the store whose token expired", async () => {
    const now = Date.now();
    mockStore.load.mockResolvedValue(
      encodeStoredToken(testToken, now - 24 * 60 * 60 * 1_000),
    );
    mockFileStore.load.mockResolvedValue(encodeStoredToken(testToken, now));

    await expect(loadToken(testEmail)).resolves.toEqual(testToken);
    expect(mockStore.clear).toHaveBeenCalledWith(testEmail);
    expect(mockFileStore.clear).not.toHaveBeenCalled();
  });

  it("warns and continues when a store cannot be read", async () => {
    mockFileStore.load.mockRejectedValue(
      new Error("ignored token file because other users can access it"),
    );
    mockStore.load.mockResolvedValue(encodeStoredToken(testToken, Date.now()));

    await expect(loadToken(testEmail)).resolves.toEqual(testToken);
    expect(stderrOutput()).toContain("other users can access it");
  });
});

describe("session expiry", () => {
  const hour = 60 * 60 * 1_000;

  function makeJwt(expiresAtMs: number): string {
    const payload = Buffer.from(
      JSON.stringify({ exp: Math.floor(expiresAtMs / 1_000) }),
    ).toString("base64url");
    return `header.${payload}.signature`;
  }

  it("uses the Skype token's own expiry when it comes before the cache lifetime", () => {
    const acquiredAt = Date.now();
    const skypeExpiry = acquiredAt + 5 * hour;

    expect(
      getSessionExpiry({ skypeToken: makeJwt(skypeExpiry) }, acquiredAt),
    ).toBe(skypeExpiry);
  });

  it("falls back to 23 hours after capture for tokens without an expiry", () => {
    const acquiredAt = Date.now();

    expect(getSessionExpiry({ skypeToken: "opaque" }, acquiredAt)).toBe(
      acquiredAt + 23 * hour,
    );
    expect(getSessionExpiry({ skypeToken: "opaque" })).toBeUndefined();
  });

  it("does not use a cached Skype token past its own expiry", async () => {
    mockStore.load.mockResolvedValueOnce(
      encodeStoredToken(
        { skypeToken: makeJwt(Date.now() + 60 * 1_000), region: "apac" },
        Date.now() - hour,
      ),
    );

    await expect(loadToken(testEmail)).resolves.toBeNull();
    expect(mockStore.clear).toHaveBeenCalledWith(testEmail);
  });

  it("inspects an expired token without clearing it", async () => {
    const expiredAt = Date.now() - hour;
    mockStore.load.mockResolvedValue(
      encodeStoredToken(
        { skypeToken: makeJwt(expiredAt), region: "apac" },
        Date.now() - 20 * hour,
      ),
    );

    const inspected = await inspectCachedToken(testEmail);

    expect(inspected?.token.region).toBe("apac");
    await expect(getCachedSessionExpiry(testEmail)).resolves.toBe(expiredAt);
    expect(mockStore.clear).not.toHaveBeenCalled();
  });

  it("reports no expiry when nothing is cached", async () => {
    await expect(getCachedSessionExpiry(testEmail)).resolves.toBeNull();
  });
});

describe("clearToken", () => {
  it("should call credential store clear", async () => {
    await clearToken(testEmail);

    expect(mockStore.clear).toHaveBeenCalledWith(testEmail);
    expect(mockFileStore.clear).toHaveBeenCalledWith(testEmail);
  });

  it("should not throw when token does not exist in credential store", async () => {
    mockStore.clear.mockImplementation(() => {
      // no-op, simulating non-existent entry
    });

    await expect(clearToken(testEmail)).resolves.not.toThrow();
  });
});
