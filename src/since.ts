/**
 * Parsing for "since" cutoffs given on the command line or by MCP tools.
 *
 * Accepted forms:
 *   - ISO-8601 date-time with `Z` or an offset: `2026-10-28T09:30:00Z`,
 *     `2026-10-28T14:30:00+05:00`
 *   - ISO-8601 date: `2026-10-28` (midnight UTC)
 *   - Duration before now: `90m`, `24h`, `7d`, `2w`
 *
 * Anything else is rejected rather than guessed, so a typo can never be
 * mistaken for "no new messages".
 */

import { InvalidArgumentError } from "./errors.js";

const DURATION_PATTERN = /^(\d+)\s*([mhdw])$/i;
const DURATION_MILLISECONDS = {
  m: 60 * 1_000,
  h: 60 * 60 * 1_000,
  d: 24 * 60 * 60 * 1_000,
  w: 7 * 24 * 60 * 60 * 1_000,
} as const;
const ISO_DATE_PATTERN = /^\d{4}-\d{2}-\d{2}$/;
const ISO_DATE_TIME_PATTERN =
  /^\d{4}-\d{2}-\d{2}T\d{2}:\d{2}(?::\d{2}(?:\.\d{1,9})?)?(?:Z|[+-]\d{2}:?\d{2})$/i;

/** Normalize a "since" value to an ISO-8601 UTC timestamp. */
export function parseSinceTimestamp(
  value: string,
  now: number = Date.now(),
): string {
  const trimmed = value.trim();

  const duration = DURATION_PATTERN.exec(trimmed);
  if (duration) {
    const unit =
      duration[2].toLowerCase() as keyof typeof DURATION_MILLISECONDS;
    return new Date(
      now - Number(duration[1]) * DURATION_MILLISECONDS[unit],
    ).toISOString();
  }

  if (ISO_DATE_PATTERN.test(trimmed)) {
    const time = Date.parse(`${trimmed}T00:00:00Z`);
    if (
      !Number.isNaN(time) &&
      new Date(time).toISOString().slice(0, 10) === trimmed
    ) {
      return new Date(time).toISOString();
    }
  } else if (ISO_DATE_TIME_PATTERN.test(trimmed)) {
    const time = Date.parse(trimmed);
    if (!Number.isNaN(time)) {
      return new Date(time).toISOString();
    }
  }

  throw new InvalidArgumentError(
    `Invalid "since" value "${value}". Use an ISO-8601 time with a timezone ` +
      `(2026-10-28T09:30:00Z), a date (2026-10-28, midnight UTC), ` +
      `or a duration such as 90m, 24h, or 7d.`,
  );
}
