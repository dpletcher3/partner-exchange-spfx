// Time/date formatting helpers for Phillips World Clock. Both functions
// route through Intl.DateTimeFormat with the configured IANA timezone so
// DST transitions are handled by the browser without any local table.

export interface ITimeParts {
  // "11:03" — hour:minute without the AM/PM marker.
  time: string;
  // "AM" / "PM" — split out so the tile can render it at a smaller size.
  // Always uppercase; the tile's SCSS lowercases it if a design pass calls
  // for that treatment.
  period: string;
}

// Returns "11:03" + "AM"/"PM" so the tile can render the period in a
// smaller font alongside the time. We use format() + a regex split here
// rather than formatToParts because formatToParts requires the
// `es2018.intl` TS lib, which the SPFx 1.23 build rig doesn't include
// and overriding the rig's tsconfig would be a bigger commitment than
// this one helper warrants.
export function formatTime(timeZone: string, now: Date): ITimeParts {
  let formatted: string;
  try {
    formatted = new Intl.DateTimeFormat('en-US', {
      timeZone,
      hour: 'numeric',
      minute: '2-digit',
      hour12: true
    }).format(now);
  } catch {
    // An invalid IANA string slipped past property-pane validation
    // somehow. Return empty so the tile renders blank rather than
    // crashing the whole grid.
    return { time: '', period: '' };
  }

  // en-US produces "H:MM AM"/"H:MM PM". Newer Chrome/Firefox use a narrow
  // no-break space (U+202F) between the time and the period — `\s+` matches
  // both regular and narrow no-break spaces, so this regex works on both
  // older and newer engines.
  const match = formatted.match(/^(\d{1,2}:\d{2})\s+(AM|PM)$/i);
  if (match) {
    return { time: match[1], period: match[2].toUpperCase() };
  }
  // Defensive fallback if en-US output ever changes shape — render the
  // raw string in the time slot rather than failing silently.
  return { time: formatted, period: '' };
}

// Returns "Sun, 5/31" — abbreviated weekday + month/day in en-US. The
// locale handles the ", " delimiter between the weekday and the date,
// so we don't have to concatenate. timeZone is honored, so a tile near
// the day boundary (e.g. Mumbai late on a US Sunday) reads as "Mon"
// once it's tipped past midnight in its own zone — not "Sun" because
// the user happens to be in the US.
export function formatDate(timeZone: string, now: Date): string {
  try {
    return new Intl.DateTimeFormat('en-US', {
      timeZone,
      weekday: 'short',
      month: 'numeric',
      day: 'numeric'
    }).format(now);
  } catch {
    return '';
  }
}

// Validates an IANA timezone string by attempting to construct a
// DateTimeFormat with it. Browsers throw a RangeError for unknown zones
// — we catch that and return false so the property pane can surface a
// clean error message instead of crashing the editor.
export function isValidTimezone(value: string): boolean {
  if (!value) {
    return false;
  }
  try {
    new Intl.DateTimeFormat('en-US', { timeZone: value });
    return true;
  } catch {
    return false;
  }
}
