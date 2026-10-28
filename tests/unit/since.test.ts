/**
 * Unit tests for "since" parsing (src/since.ts).
 */

import { describe, it, expect } from "vitest";
import { parseDurationMs, parseSinceTimestamp } from "../../src/since.js";
import { InvalidArgumentError } from "../../src/errors.js";

const now = Date.parse("2026-10-28T12:00:00.000Z");

describe("parseSinceTimestamp", () => {
  it.each([
    ["2026-10-28T09:30:00Z", "2026-10-28T09:30:00.000Z"],
    ["2026-10-28T09:30:00.123Z", "2026-10-28T09:30:00.123Z"],
    ["2026-10-28T14:30:00+05:00", "2026-10-28T09:30:00.000Z"],
    ["2026-10-28T14:30+05:00", "2026-10-28T09:30:00.000Z"],
    ["2026-10-28", "2026-10-28T00:00:00.000Z"],
    ["90m", "2026-10-28T10:30:00.000Z"],
    ["24h", "2026-10-27T12:00:00.000Z"],
    ["7d", "2026-10-21T12:00:00.000Z"],
    ["2w", "2026-10-14T12:00:00.000Z"],
    [" 24H ", "2026-10-27T12:00:00.000Z"],
  ])("normalizes %s to %s", (input, expected) => {
    expect(parseSinceTimestamp(input, now)).toBe(expected);
  });

  it.each([
    "yesterday",
    "",
    "2026-02-30",
    "2026-10-28T09:30:00",
    "October 28 2026",
    "24 hours",
    "-1h",
  ])("rejects %j instead of guessing", (input) => {
    expect(() => parseSinceTimestamp(input, now)).toThrow(InvalidArgumentError);
  });
});

describe("parseDurationMs", () => {
  it.each([
    ["45m", 45 * 60 * 1_000],
    ["2h", 2 * 60 * 60 * 1_000],
    ["1d", 24 * 60 * 60 * 1_000],
  ])("parses %s", (input, expected) => {
    expect(parseDurationMs(input)).toBe(expected);
  });

  it.each(["", "45", "45 minutes", "-1h"])("rejects %j", (input) => {
    expect(() => parseDurationMs(input)).toThrow(InvalidArgumentError);
  });
});
