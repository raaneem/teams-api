/**
 * Interactive browser login for Teams.
 *
 * Opens a visible Chromium window (Playwright's bundled browser),
 * navigates to Teams, and waits for the user to complete login manually.
 * Works on all platforms (macOS, Windows, Linux).
 *
 * With TEAMS_NON_INTERACTIVE=true (unattended use, e.g. over SSH), the saved
 * browser session may still renew tokens silently, but login fails fast with
 * AuthRequiredError instead of waiting for a person.
 */

import type { TeamsToken, InteractiveLoginOptions } from "../types.js";
import { resolveTeamsRegion } from "../region.js";
import { diagnosePageState } from "./page-diagnostics.js";
import { AuthRequiredError } from "../errors.js";
import { redactUrl } from "../redact.js";
import { isNonInteractiveMode } from "../runtime-flags.js";
import {
  launchInteractiveBrowserContext,
  getDefaultBrowserProfileDir,
} from "../browser-runtime.js";
import {
  captureTokensFromPage,
  TOKEN_INTERCEPT_TIMEOUT,
  PAGE_STATE_POLL_INTERVAL,
  type LogFunction,
} from "./token-capture.js";

const TEAMS_URL = "https://teams.cloud.microsoft/";
const INTERACTIVE_LOGIN_TIMEOUT = 5 * 60 * 1_000;
/** Long enough for a saved session to load Teams, too short to wait for a person. */
const NON_INTERACTIVE_LOGIN_TIMEOUT = 45 * 1_000;

/**
 * Acquire a Teams skype token via interactive browser login.
 *
 * Opens a visible Chromium window (Playwright's bundled browser),
 * navigates to Teams, and waits for the user to complete login
 * manually. Once Teams loads, the skype token is captured via
 * CDP Fetch interception.
 *
 * Works on all platforms (macOS, Windows, Linux). No FIDO2 passkey
 * or system Chrome required.
 */
export async function acquireTokenViaInteractiveLogin(
  options?: InteractiveLoginOptions,
): Promise<TeamsToken> {
  const { chromium } = await import("playwright");
  const log: LogFunction =
    options?.log ?? (options?.verbose ? console.error.bind(console) : () => {});
  const nonInteractive = isNonInteractiveMode();
  const loginTimeout = nonInteractive
    ? NON_INTERACTIVE_LOGIN_TIMEOUT
    : INTERACTIVE_LOGIN_TIMEOUT;

  log("Launching browser for interactive login...");
  const profileDir = getDefaultBrowserProfileDir();
  log(`Using browser profile: ${profileDir}`);
  const context = await launchInteractiveBrowserContext(
    chromium,
    log,
    profileDir,
  );

  const existingPages = context.pages();
  const page = existingPages[0] ?? (await context.newPage());
  for (const stalePage of existingPages.slice(1)) {
    await stalePage.close();
  }

  try {
    log("Navigating to Teams...");
    await page.goto(TEAMS_URL, {
      waitUntil: "domcontentloaded",
      timeout: loginTimeout,
    });

    // Pre-fill email if provided (nobody is there to use it when non-interactive)
    if (options?.email && !nonInteractive) {
      try {
        await page.waitForURL(
          /login\.microsoftonline\.com|login\.microsoft\.com/,
          { timeout: 30_000 },
        );

        const emailInput = await page
          .waitForSelector(
            '#i0116, input[name="loginfmt"], input[type="email"]',
            { timeout: 10_000 },
          )
          .catch(() => null);

        if (emailInput) {
          await emailInput.fill(options.email);
          log(`Pre-filled email: ${options.email}`);
        }
      } catch {
        log("Could not pre-fill email (login page may have changed)");
      }
    }

    // Wait for user to complete login, polling page state for diagnostics
    log(
      nonInteractive
        ? "Waiting for Teams to load from the saved browser session..."
        : "Waiting for you to complete login in the browser window...",
    );
    const loginDeadline = Date.now() + loginTimeout;
    while (Date.now() < loginDeadline) {
      const state = await diagnosePageState(page);
      if (state.phase === "teams-loaded" || state.phase === "teams-loading") {
        log(`Reached Teams: ${state.detail}`);
        break;
      }
      log(`Login progress: ${state.detail}`);
      await new Promise((resolve) =>
        setTimeout(resolve, PAGE_STATE_POLL_INTERVAL),
      );
    }

    if (
      !page.url().includes("teams.cloud.microsoft") &&
      !page.url().includes("teams.microsoft.com")
    ) {
      const state = await diagnosePageState(page);
      if (nonInteractive) {
        throw new AuthRequiredError(
          `Teams needs an interactive sign-in: the saved browser session did not reach Teams ` +
            `within ${loginTimeout / 1_000}s (${state.detail}). ` +
            `Run "teams-api auth" where you can see the browser window, then retry.`,
        );
      }
      throw new Error(
        `Login did not complete within ${loginTimeout / 1_000}s. ` +
          `Current state: ${state.phase} — ${state.detail} (URL: ${redactUrl(state.url)})`,
      );
    }

    log("Login detected, capturing token...");

    const {
      skypeToken,
      chatBearerToken,
      region,
      bearerToken,
      substrateToken,
      amsToken,
      sharePointToken,
      sharePointHost,
    } = await captureTokensFromPage(page, log, TOKEN_INTERCEPT_TIMEOUT);

    return {
      skypeToken,
      chatBearerToken,
      region: resolveTeamsRegion(options?.region, region),
      bearerToken,
      substrateToken,
      amsToken,
      sharePointToken,
      sharePointHost,
    };
  } finally {
    await context.close();
  }
}
