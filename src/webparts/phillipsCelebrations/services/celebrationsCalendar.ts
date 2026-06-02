// Pure celebrations date logic — NO SharePoint/SPFx deps, so it unit-tests in
// isolation. Callers pass `today` (the web part passes new Date(); tests pass
// fixed dates) and people with raw ISO date strings.
//
// Rules (spec §3):
//   - Window: current week + next week (14 days), anchored to the configured
//     week start; includes already-passed days earlier this week.
//   - Birthdays/anniversaries recur annually — match on month/day, ignore the
//     stored year (the occurrence year comes from the window day).
//   - Feb 29 is observed on Feb 28 in non-leap years (both birthdays & anniversaries).
//   - Anniversary years = window-occurrence year − hire year; year 0 (hired this
//     year) is excluded.
//   - Within a day, sort by name.

export type WeekStart = 'sunday' | 'monday';
export type EventType = 'birthday' | 'anniversary';

export interface IPersonInput {
  id: number;
  name: string;
  birthDate?: string; // ISO; month/day used, year ignored
  hireDate?: string; // ISO; month/day used; year drives years-of-service
}

export interface ICelebrationEvent {
  personId: number;
  name: string;
  type: EventType;
  date: Date; // the observed day within the window (local midnight)
  years?: number; // anniversaries only (>= 1)
}

export interface ICelebrationsResult {
  window: Date[]; // 14 local-midnight days
  eventsByDay: ICelebrationEvent[][]; // aligned to `window`
}

function isLeapYear(year: number): boolean {
  return (year % 4 === 0 && year % 100 !== 0) || year % 400 === 0;
}

// Start-of-week (local midnight) for `today` given the configured week start.
export function startOfWeek(today: Date, weekStart: WeekStart): Date {
  const d = new Date(today.getFullYear(), today.getMonth(), today.getDate());
  const startDow = weekStart === 'monday' ? 1 : 0;
  let diff = d.getDay() - startDow;
  if (diff < 0) {
    diff += 7;
  }
  d.setDate(d.getDate() - diff);
  return d;
}

// The 14-day window: current week + next week, from the week start.
export function computeWindow(today: Date, weekStart: WeekStart): Date[] {
  const start = startOfWeek(today, weekStart);
  const days: Date[] = [];
  for (let i = 0; i < 14; i++) {
    days.push(new Date(start.getFullYear(), start.getMonth(), start.getDate() + i));
  }
  return days;
}

// Parse the calendar date parts from an ISO string by slicing the YYYY-MM-DD
// prefix — tz-proof for SharePoint date-only fields (stored as UTC-midnight of
// the intended date). Returns 0-based month.
export function parseDateParts(
  iso: string | undefined
): { year: number; month: number; day: number } | undefined {
  if (!iso) {
    return undefined;
  }
  const m = /^(\d{4})-(\d{2})-(\d{2})/.exec(iso);
  if (!m) {
    return undefined;
  }
  return { year: Number(m[1]), month: Number(m[2]) - 1, day: Number(m[3]) };
}

// The observed [month, day] of a recurring (srcMonth, srcDay) in a given year:
// Feb 29 → Feb 28 in non-leap years; everything else unchanged.
function observedMonthDay(srcMonth: number, srcDay: number, year: number): [number, number] {
  if (srcMonth === 1 && srcDay === 29 && !isLeapYear(year)) {
    return [1, 28];
  }
  return [srcMonth, srcDay];
}

// Window index whose day matches the recurring (srcMonth, srcDay) for that day's
// own year (Feb-29 aware), or -1. Per-day year handles the Dec→Jan boundary.
function findWindowIndex(window: Date[], srcMonth: number, srcDay: number): number {
  for (let i = 0; i < window.length; i++) {
    const wd = window[i];
    const [om, od] = observedMonthDay(srcMonth, srcDay, wd.getFullYear());
    if (wd.getMonth() === om && wd.getDate() === od) {
      return i;
    }
  }
  return -1;
}

export function buildCelebrations(
  people: IPersonInput[],
  today: Date,
  weekStart: WeekStart
): ICelebrationsResult {
  const window = computeWindow(today, weekStart);
  const eventsByDay: ICelebrationEvent[][] = window.map(() => []);

  for (const person of people) {
    const bd = parseDateParts(person.birthDate);
    if (bd) {
      const idx = findWindowIndex(window, bd.month, bd.day);
      if (idx >= 0) {
        eventsByDay[idx].push({
          personId: person.id,
          name: person.name,
          type: 'birthday',
          date: window[idx]
        });
      }
    }

    const hd = parseDateParts(person.hireDate);
    if (hd) {
      const idx = findWindowIndex(window, hd.month, hd.day);
      if (idx >= 0) {
        const years = window[idx].getFullYear() - hd.year;
        if (years >= 1) {
          // Exclude year 0 (hired this year → not an anniversary yet).
          eventsByDay[idx].push({
            personId: person.id,
            name: person.name,
            type: 'anniversary',
            date: window[idx],
            years
          });
        }
      }
    }
  }

  for (const day of eventsByDay) {
    day.sort((a, b) => a.name.localeCompare(b.name));
  }

  return { window, eventsByDay };
}
