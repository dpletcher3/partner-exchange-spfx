import { computeGreeting, extractFirstName } from '../timeOfDay';

// Deterministic dates (fixed Y/M/D) so getHours() drives the result, not now().
function at(hour: number, minute = 0): Date {
  return new Date(2026, 4, 26, hour, minute, 0);
}

describe('computeGreeting', () => {
  it('00:00 -> Good evening', () => {
    expect(computeGreeting(at(0, 0))).toBe('Good evening');
  });
  it('04:59 -> Good evening', () => {
    expect(computeGreeting(at(4, 59))).toBe('Good evening');
  });
  it('05:00 -> Good morning', () => {
    expect(computeGreeting(at(5, 0))).toBe('Good morning');
  });
  it('11:59 -> Good morning', () => {
    expect(computeGreeting(at(11, 59))).toBe('Good morning');
  });
  it('12:00 -> Good afternoon', () => {
    expect(computeGreeting(at(12, 0))).toBe('Good afternoon');
  });
  it('17:59 -> Good afternoon', () => {
    expect(computeGreeting(at(17, 59))).toBe('Good afternoon');
  });
  it('18:00 -> Good evening', () => {
    expect(computeGreeting(at(18, 0))).toBe('Good evening');
  });
  it('22:59 -> Good evening', () => {
    expect(computeGreeting(at(22, 59))).toBe('Good evening');
  });
  it('23:00 -> Good evening', () => {
    expect(computeGreeting(at(23, 0))).toBe('Good evening');
  });
});

describe('extractFirstName', () => {
  it('"Dan Pletcher" -> "Dan"', () => {
    expect(extractFirstName('Dan Pletcher')).toBe('Dan');
  });
  it('"Dan" -> "Dan"', () => {
    expect(extractFirstName('Dan')).toBe('Dan');
  });
  it('empty string -> undefined', () => {
    expect(extractFirstName('')).toBeUndefined();
  });
  it('null -> undefined', () => {
    expect(extractFirstName(null)).toBeUndefined();
  });
  it('undefined -> undefined', () => {
    expect(extractFirstName(undefined)).toBeUndefined();
  });
  it('extra whitespace "  Dan  Pletcher  " -> "Dan"', () => {
    expect(extractFirstName('  Dan  Pletcher  ')).toBe('Dan');
  });
  it('hyphenated "Dan-Pletcher" stays whole', () => {
    expect(extractFirstName('Dan-Pletcher')).toBe('Dan-Pletcher');
  });
});
