/**
 * A secret-free description of a token bundle: which tokens are present and
 * when they expire. Used by `teams-api auth` instead of printing the tokens.
 */

import type { TeamsToken } from "./types.js";

function readJwtClaim(token: string | undefined, claim: string): unknown {
  if (!token) return undefined;
  try {
    const payload = JSON.parse(
      Buffer.from(token.split(".")[1] ?? "", "base64url").toString("utf-8"),
    ) as Record<string, unknown>;
    return payload[claim];
  } catch {
    return undefined;
  }
}

function describeToken(token: string | undefined): {
  present: boolean;
  expiresAt?: string | null;
} {
  if (!token) return { present: false };
  const expiry = readJwtClaim(token, "exp");
  return {
    present: true,
    expiresAt:
      typeof expiry === "number"
        ? new Date(expiry * 1_000).toISOString()
        : null,
  };
}

export function describeTokenStatus(token: TeamsToken) {
  const account = readJwtClaim(token.bearerToken, "upn");
  return {
    authenticated: Boolean(token.skypeToken || token.chatBearerToken),
    account: typeof account === "string" ? account : null,
    region: token.region,
    tokens: {
      skype: describeToken(token.skypeToken),
      chatBearer: describeToken(token.chatBearerToken),
      middleTier: describeToken(token.bearerToken),
      substrate: describeToken(token.substrateToken),
      ams: describeToken(token.amsToken),
      sharePoint: describeToken(token.sharePointToken),
    },
  };
}
