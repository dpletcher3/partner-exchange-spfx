import {
  buildCelebrations,
  computeWindow,
  startOfWeek,
  ICelebrationEvent,
  ICelebrationsResult,
  IPersonInput
} from '../celebrationsCalendar';

const pad = (n: number): string => (n < 10 ? `0${n}` : `${n}`);
const ymd = (d: Date): [number, number, number] => [d.getFullYear(), d.getMonth(), d.getDate()];
const flatten = (r: ICelebrationsResult): ICelebrationEvent[] =>
  r.eventsByDay.reduce<ICelebrationEvent[]>((acc, day) => acc.concat(day), []);
const find = (r: ICelebrationsResult, personId: number): ICelebrationEvent[] =>
  flatten(r).filter((e) => e.personId === personId);
// Build an ISO date string for a given year + a window day's month/day.
const isoFor = (year: number, d: Date): string => `${year}-${pad(d.getMonth() + 1)}-${pad(d.getDate())}`;

describe('startOfWeek / computeWindow', () => {
  const today = new Date(2026, 5, 3); // arbitrary; assertions are DOW-independent

  it('Sunday week start lands on a Sunday and the window is 14 days', () => {
    const start = startOfWeek(today, 'sunday');
    expect(start.getDay()).toBe(0);
    const win = computeWindow(today, 'sunday');
    expect(win.length).toBe(14);
    expect(ymd(win[0])).toEqual(ymd(start));
    expect(start.getTime()).toBeLessThanOrEqual(new Date(2026, 5, 3).getTime());
  });

  it('Monday week start lands on a Monday and differs from Sunday start', () => {
    const sun = startOfWeek(today, 'sunday');
    const mon = startOfWeek(today, 'monday');
    expect(mon.getDay()).toBe(1);
    expect(mon.getTime()).not.toBe(sun.getTime());
  });
});

describe('buildCelebrations — birthdays in the window', () => {
  const today = new Date(2026, 5, 3);
  const win = computeWindow(today, 'sunday');

  it('mid-window birthday lands on its day, typed birthday', () => {
    const mid = win[5];
    const r = buildCelebrations([{ id: 1, name: 'Mid', birthDate: isoFor(1990, mid) }], today, 'sunday');
    const ev = find(r, 1);
    expect(ev).toHaveLength(1);
    expect(ev[0].type).toBe('birthday');
    expect(ymd(ev[0].date)).toEqual(ymd(mid));
  });

  it('a birthday ON today is included', () => {
    const r = buildCelebrations([{ id: 2, name: 'Today', birthDate: isoFor(1985, today) }], today, 'sunday');
    expect(ymd(find(r, 2)[0].date)).toEqual(ymd(new Date(2026, 5, 3)));
  });

  it('an already-passed day earlier this week is still shown (window[0])', () => {
    const passed = win[0]; // week start — on/before today
    const r = buildCelebrations([{ id: 3, name: 'Passed', birthDate: isoFor(1979, passed) }], today, 'sunday');
    expect(ymd(find(r, 3)[0].date)).toEqual(ymd(passed));
  });

  it('a birthday outside the window produces no event', () => {
    const outside = new Date(win[13].getFullYear(), win[13].getMonth(), win[13].getDate() + 10);
    const r = buildCelebrations([{ id: 4, name: 'Out', birthDate: isoFor(1990, outside) }], today, 'sunday');
    expect(find(r, 4)).toHaveLength(0);
  });

  it('within a day, people sort by name', () => {
    const day = win[4];
    const people: IPersonInput[] = [
      { id: 10, name: 'Zoe', birthDate: isoFor(1990, day) },
      { id: 11, name: 'Aaron', birthDate: isoFor(1991, day) }
    ];
    const r = buildCelebrations(people, today, 'sunday');
    const names = flatten(r).map((e) => e.name);
    expect(names).toEqual(['Aaron', 'Zoe']);
  });
});

describe('buildCelebrations — month and year boundaries', () => {
  it('window crossing a month boundary matches days in both months', () => {
    const today = new Date(2026, 4, 28); // late May → window spans into June
    const win = computeWindow(today, 'sunday');
    const may = win.find((d) => d.getMonth() === 4)!;
    const jun = win.find((d) => d.getMonth() === 5)!;
    expect(may).toBeDefined();
    expect(jun).toBeDefined();
    const r = buildCelebrations(
      [
        { id: 1, name: 'May', birthDate: isoFor(1990, may) },
        { id: 2, name: 'Jun', birthDate: isoFor(1990, jun) }
      ],
      today,
      'sunday'
    );
    expect(ymd(find(r, 1)[0].date)).toEqual(ymd(may));
    expect(ymd(find(r, 2)[0].date)).toEqual(ymd(jun));
  });

  it('window crossing Dec→Jan uses the occurrence year (incl years-of-service)', () => {
    const today = new Date(2026, 11, 30); // Dec 30 2026 → window reaches Jan 2027
    const win = computeWindow(today, 'sunday');
    const jan = win.find((d) => d.getMonth() === 0)!; // a January 2027 day
    expect(jan).toBeDefined();
    expect(jan.getFullYear()).toBe(2027);
    const r = buildCelebrations(
      [{ id: 1, name: 'NY', birthDate: isoFor(1990, jan), hireDate: isoFor(2020, jan) }],
      today,
      'sunday'
    );
    const evs = find(r, 1);
    const bday = evs.find((e) => e.type === 'birthday')!;
    const anniv = evs.find((e) => e.type === 'anniversary')!;
    expect(ymd(bday.date)).toEqual(ymd(jan));
    expect(anniv.years).toBe(7); // 2027 − 2020
  });
});

describe('buildCelebrations — Feb 29 rule', () => {
  it('leap year: Feb-29 birthday observed on Feb 29', () => {
    const today = new Date(2028, 1, 29); // Feb 29 2028 (leap)
    const win = computeWindow(today, 'sunday');
    const feb29 = win.find((d) => d.getMonth() === 1 && d.getDate() === 29)!;
    expect(feb29).toBeDefined();
    const r = buildCelebrations([{ id: 1, name: 'Leaper', birthDate: '2000-02-29' }], today, 'sunday');
    const ev = find(r, 1);
    expect(ev).toHaveLength(1);
    expect(ymd(ev[0].date)).toEqual(ymd(feb29));
  });

  it('non-leap year: Feb-29 birthday observed on Feb 28', () => {
    const today = new Date(2027, 1, 28); // Feb 28 2027 (non-leap)
    const win = computeWindow(today, 'sunday');
    const feb28 = win.find((d) => d.getMonth() === 1 && d.getDate() === 28)!;
    expect(feb28).toBeDefined();
    expect(win.find((d) => d.getMonth() === 1 && d.getDate() === 29)).toBeUndefined(); // no Feb 29 in 2027
    const r = buildCelebrations([{ id: 1, name: 'Leaper', birthDate: '2000-02-29' }], today, 'sunday');
    const ev = find(r, 1);
    expect(ev).toHaveLength(1);
    expect(ymd(ev[0].date)).toEqual(ymd(feb28));
  });
});

describe('buildCelebrations — years of service', () => {
  const today = new Date(2026, 5, 3);
  const win = computeWindow(today, 'sunday');
  const day = win[6];

  it('hired this year (year 0) is excluded — no anniversary', () => {
    const r = buildCelebrations(
      [{ id: 1, name: 'Newbie', hireDate: isoFor(day.getFullYear(), day) }],
      today,
      'sunday'
    );
    expect(find(r, 1)).toHaveLength(0);
  });

  it('hired N years before the occurrence shows N years', () => {
    const r = buildCelebrations(
      [{ id: 2, name: 'Veteran', hireDate: isoFor(day.getFullYear() - 7, day) }],
      today,
      'sunday'
    );
    const ev = find(r, 2);
    expect(ev).toHaveLength(1);
    expect(ev[0].type).toBe('anniversary');
    expect(ev[0].years).toBe(7);
  });
});
