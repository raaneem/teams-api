/**
 * Unit tests for non-interactive behavior of interactive login
 * (src/auth/interactive.ts). The browser is a stub: nothing is launched.
 */

import { describe, it, expect, vi, afterEach } from "vitest";

vi.mock("playwright", () => ({ chromium: {} }));

const page = vi.hoisted(() => ({
  goto: vi.fn(),
  url: vi.fn(
    () =>
      "https://login.microsoftonline.com/common/oauth2/v2.0/authorize?client_id=x&login_hint=alice%40contoso.com",
  ),
  evaluate: vi.fn(),
  waitForURL: vi.fn(),
  waitForSelector: vi.fn(),
  close: vi.fn(),
}));

const context = vi.hoisted(() => ({
  pages: vi.fn(() => [page]),
  newPage: vi.fn(),
  close: vi.fn(),
}));

vi.mock("../../src/browser-runtime.js", () => ({
  launchInteractiveBrowserContext: vi.fn(async () => context),
  getDefaultBrowserProfileDir: vi.fn(() => "/tmp/teams-api-test-profile"),
}));

import { acquireTokenViaInteractiveLogin } from "../../src/auth/interactive.js";
import { AuthRequiredError } from "../../src/errors.js";

afterEach(() => {
  vi.useRealTimers();
  vi.unstubAllEnvs();
});

describe("acquireTokenViaInteractiveLogin with TEAMS_NON_INTERACTIVE", () => {
  it("fails fast with AuthRequiredError when the saved session needs a person", async () => {
    vi.stubEnv("TEAMS_NON_INTERACTIVE", "true");
    vi.useFakeTimers();
    page.goto.mockResolvedValue(undefined);
    page.evaluate.mockResolvedValue("Pick an account");
    context.close.mockResolvedValue(undefined);

    const outcome = acquireTokenViaInteractiveLogin({
      email: "alice@contoso.com",
    }).catch((error: unknown) => error);
    await vi.advanceTimersByTimeAsync(50_000);
    const error = await outcome;

    expect(error).toBeInstanceOf(AuthRequiredError);
    expect((error as Error).message).toContain('Run "teams-api auth"');
    expect((error as Error).message).not.toContain("login_hint");
    expect(page.goto).toHaveBeenCalledWith(
      "https://teams.cloud.microsoft/",
      expect.objectContaining({ timeout: 45_000 }),
    );
    // Nobody is there to use a pre-filled email, so it does not wait for one.
    expect(page.waitForURL).not.toHaveBeenCalled();
    expect(context.close).toHaveBeenCalled();
  });
});
