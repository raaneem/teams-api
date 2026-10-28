/**
 * Unit tests for the background sign-in refresher (src/token-refresher.ts).
 * The clock, cache, and refresh are all fakes.
 */

import { describe, it, expect, vi } from "vitest";
import {
  RETRY_AFTER_FAILURE_MS,
  isRefreshDue,
  startTokenRefresher,
} from "../../src/token-refresher.js";

const minute = 60 * 1_000;
const hour = 60 * minute;
const margin = 45 * minute;
const expiresAt = Date.parse("2026-10-29T15:00:00Z");

describe("isRefreshDue", () => {
  const base = { expiresAt, marginMs: margin, lastAttemptFailed: false };

  it.each([
    ["far from expiry", expiresAt - 2 * hour, undefined, false],
    [
      "inside the margin, never tried",
      expiresAt - 30 * minute,
      undefined,
      true,
    ],
    [
      "inside the margin, already tried there",
      expiresAt - 10 * minute,
      expiresAt - 30 * minute,
      false,
    ],
    [
      "inside the margin, last try was before the window",
      expiresAt - 10 * minute,
      expiresAt - 2 * hour,
      true,
    ],
    [
      "just expired, last try was before expiry",
      expiresAt + minute,
      expiresAt - 30 * minute,
      true,
    ],
    [
      "expired, already tried after expiry",
      expiresAt + 30 * minute,
      expiresAt + minute,
      false,
    ],
    [
      "expired, last try 2 hours ago",
      expiresAt + 3 * hour,
      expiresAt + minute,
      true,
    ],
  ])("%s", (_label, now, lastAttemptAt, expected) => {
    expect(isRefreshDue({ ...base, now, lastAttemptAt })).toBe(expected);
  });

  it("waits after a failure, even inside the margin", () => {
    const lastAttemptAt = expiresAt - 40 * minute;
    expect(
      isRefreshDue({
        ...base,
        now: expiresAt - 10 * minute,
        lastAttemptAt,
        lastAttemptFailed: true,
      }),
    ).toBe(false);
    expect(
      isRefreshDue({
        ...base,
        now: lastAttemptAt + RETRY_AFTER_FAILURE_MS,
        lastAttemptAt,
        lastAttemptFailed: true,
      }),
    ).toBe(true);
  });
});

async function createHarness(options: {
  expiry: number | null;
  refresh?: () => Promise<void>;
}) {
  let now = expiresAt - 2 * hour;
  let currentExpiry = options.expiry;
  const logs: string[] = [];
  const refresh = vi.fn(options.refresh ?? (async () => {}));
  const refresher = startTokenRefresher({
    marginMs: margin,
    intervalMs: 24 * hour,
    now: () => now,
    getExpiry: async () => currentExpiry,
    refresh,
    log: (message) => logs.push(message),
  });
  await refresher.ready;
  return {
    refresher,
    refresh,
    logs,
    setNow: (value: number) => {
      now = value;
    },
    setExpiry: (value: number | null) => {
      currentExpiry = value;
    },
  };
}

describe("startTokenRefresher", () => {
  it("does nothing while nobody has signed in", async () => {
    const harness = await createHarness({ expiry: null });
    harness.setNow(expiresAt);

    await harness.refresher.check();
    harness.refresher.stop();

    expect(harness.refresh).not.toHaveBeenCalled();
  });

  it("refreshes once inside the margin and reports the new expiry", async () => {
    const harness = await createHarness({ expiry: expiresAt });
    harness.refresh.mockImplementation(async () => {
      harness.setExpiry(expiresAt + 24 * hour);
    });

    await harness.refresher.check();
    expect(harness.refresh).not.toHaveBeenCalled();

    harness.setNow(expiresAt - 30 * minute);
    await harness.refresher.check();
    harness.refresher.stop();

    expect(harness.refresh).toHaveBeenCalledTimes(1);
    expect(harness.logs.join("\n")).toContain(
      "valid until 2026-10-30T15:00:00.000Z",
    );
  });

  it("tries again right after expiry when Teams keeps the old sign-in", async () => {
    const harness = await createHarness({ expiry: expiresAt });

    harness.setNow(expiresAt - 30 * minute);
    await harness.refresher.check();
    harness.setNow(expiresAt - 10 * minute);
    await harness.refresher.check();
    expect(harness.refresh).toHaveBeenCalledTimes(1);
    expect(harness.logs.join("\n")).toContain("kept the current sign-in");

    harness.setNow(expiresAt + minute);
    await harness.refresher.check();
    harness.refresher.stop();

    expect(harness.refresh).toHaveBeenCalledTimes(2);
  });

  it("waits 2 hours after a failed refresh instead of retrying", async () => {
    const harness = await createHarness({
      expiry: expiresAt,
      refresh: async () => {
        throw new Error("sign-in required");
      },
    });

    harness.setNow(expiresAt - 30 * minute);
    await harness.refresher.check();
    harness.setNow(expiresAt + minute);
    await harness.refresher.check();
    expect(harness.refresh).toHaveBeenCalledTimes(1);
    expect(harness.logs.join("\n")).toContain("Trying again in 2 hours");

    harness.setNow(expiresAt - 30 * minute + RETRY_AFTER_FAILURE_MS);
    await harness.refresher.check();
    harness.refresher.stop();

    expect(harness.refresh).toHaveBeenCalledTimes(2);
  });

  it("never runs two checks at once", async () => {
    let finish: () => void = () => {};
    const harness = await createHarness({
      expiry: expiresAt,
      refresh: () =>
        new Promise<void>((resolve) => {
          finish = resolve;
        }),
    });
    harness.setNow(expiresAt - 30 * minute);

    const first = harness.refresher.check();
    await harness.refresher.check();
    await Promise.resolve();
    finish();
    await first;
    harness.refresher.stop();

    expect(harness.refresh).toHaveBeenCalledTimes(1);
  });
});
