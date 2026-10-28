/**
 * Unit tests for secret redaction (src/redact.ts).
 */

import { describe, it, expect, beforeEach } from "vitest";
import {
  redactKnownSecrets,
  redactSecrets,
  redactUrl,
  registerSecret,
  registerTokenSecrets,
  resetRegisteredSecrets,
  safeErrorMessage,
} from "../../src/redact.js";

const fakeJwt =
  "eyJhbGciOiJSUzI1NiIsInR5cCI6IkpXVCJ9.eyJzdWIiOiJ0ZXN0LXVzZXIiLCJleHAiOjE3NjE2OTk2MDB9.c2lnbmF0dXJlLXBhcnQtb2YtdGhlLWp3dA";

beforeEach(() => {
  resetRegisteredSecrets();
});

describe("redactKnownSecrets", () => {
  it("replaces every registered secret, longest first", () => {
    const token = "skype-token-value-1234567890";
    const bundle = `bundle-containing-${token}-and-more`;
    registerSecret(token);
    registerSecret(bundle);

    expect(redactKnownSecrets(`a ${bundle} b ${token} c`)).toBe(
      "a [REDACTED] b [REDACTED] c",
    );
  });

  it("ignores values too short to be credentials", () => {
    registerSecret("apac");
    expect(redactKnownSecrets("region apac")).toBe("region apac");
  });

  it("registers every credential of a token bundle", () => {
    registerTokenSecrets({
      skypeToken: "skype-token-value-1234567890",
      chatBearerToken: "chat-bearer-value-1234567890",
      bearerToken: "middle-tier-value-1234567890",
      substrateToken: "substrate-value-1234567890",
      amsToken: "ams-token-value-1234567890",
      sharePointToken: "sharepoint-value-1234567890",
      region: "apac",
      sharePointHost: "contoso-my.sharepoint.com",
    });

    const output = redactKnownSecrets(
      "skype-token-value-1234567890 chat-bearer-value-1234567890 middle-tier-value-1234567890 " +
        "substrate-value-1234567890 ams-token-value-1234567890 sharepoint-value-1234567890 " +
        "apac contoso-my.sharepoint.com",
    );

    expect(output).toBe(
      "[REDACTED] [REDACTED] [REDACTED] [REDACTED] [REDACTED] [REDACTED] apac contoso-my.sharepoint.com",
    );
  });

  it("tolerates a missing token bundle", () => {
    expect(() => registerTokenSecrets(undefined)).not.toThrow();
  });
});

describe("redactSecrets patterns", () => {
  it.each([
    ["a JWT", `token ${fakeJwt} end`, "token [REDACTED] end"],
    [
      "a Skype authentication header",
      "Authentication: skypetoken=abc.def-ghi_jkl",
      "Authentication: skypetoken=[REDACTED]",
    ],
    [
      "an AMS skype_token header",
      "Authorization: skype_token abcdefghijklmnop",
      "Authorization: skype_token [REDACTED]",
    ],
    [
      "a Bearer header",
      "Authorization: Bearer abcdefghijklmnopqrstuvwxyz0123",
      "Authorization: Bearer [REDACTED]",
    ],
    [
      "a cookie header",
      "Cookie: skypetoken_asm=secretvalue",
      "Cookie: [REDACTED]",
    ],
    [
      "the security -w argument",
      "Command failed: security add-generic-password -a me@contoso.com -s teams-api -w eyJza3lwZVRva2VuIjoi+/= -U",
      "Command failed: security add-generic-password -a me@contoso.com -s teams-api -w [REDACTED] -U",
    ],
    [
      "OAuth URL parameters",
      "https://teams.cloud.microsoft/v2/#code=0.AbC-def&state=xyz&client_info=eyJ1aWQi",
      "https://teams.cloud.microsoft/v2/#code=[REDACTED]&state=[REDACTED]&client_info=[REDACTED]",
    ],
    [
      "a long base64 run",
      `blob ${"QUJD".repeat(60)} end`,
      "blob [REDACTED] end",
    ],
  ])("redacts %s", (_label, input, expected) => {
    expect(redactSecrets(input)).toBe(expected);
  });

  it("leaves ordinary error messages alone", () => {
    const message =
      "Bearer token is missing — re-authentication required for profile resolution. Authentication failed: 401 Unauthorized";
    expect(redactSecrets(message)).toBe(message);
  });

  it("keeps surrounding JSON valid", () => {
    const json = JSON.stringify({
      header: "Authorization: Bearer abcdefghijklmnopqrstuvwxyz0123",
      jwt: fakeJwt,
    });
    expect(() => JSON.parse(redactSecrets(json))).not.toThrow();
  });
});

describe("safeErrorMessage", () => {
  it("uses only the redacted message, never the stack or attached output", () => {
    const secret = "registered-secret-value-1234567890";
    registerSecret(secret);
    const error = Object.assign(new Error(`failed with ${secret}`), {
      stdout: Buffer.from(secret),
      stderr: Buffer.from(secret),
    });

    const message = safeErrorMessage(error);

    expect(message).toBe("failed with [REDACTED]");
  });

  it("describes non-error values without dumping them", () => {
    expect(safeErrorMessage({ token: "x" })).toBe("Unknown error");
    expect(safeErrorMessage("plain text")).toBe("plain text");
  });
});

describe("redactUrl", () => {
  it("drops query strings and fragments", () => {
    expect(
      redactUrl(
        "https://login.microsoftonline.com/common/oauth2/v2.0/authorize?client_id=x&login_hint=me#state=y",
      ),
    ).toBe(
      "https://login.microsoftonline.com/common/oauth2/v2.0/authorize?[REDACTED]",
    );
  });

  it("keeps clean URLs unchanged", () => {
    expect(redactUrl("https://teams.cloud.microsoft/v2/")).toBe(
      "https://teams.cloud.microsoft/v2/",
    );
  });

  it("handles unparseable and opaque URLs", () => {
    expect(redactUrl("not a url")).toBe("[unparseable URL]");
    expect(redactUrl("about:blank")).toBe("about:blank");
  });
});
