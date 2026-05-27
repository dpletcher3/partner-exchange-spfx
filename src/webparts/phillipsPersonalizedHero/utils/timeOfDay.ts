// Pure helpers for the Personalized Hero greeting. No SPFx dependency, so they
// are unit-testable in isolation.

// Time-of-day greeting from browser-local hour. Boundaries (per spec):
//   05:00-11:59 -> Good morning
//   12:00-17:59 -> Good afternoon
//   18:00-04:59 -> Good evening
export function computeGreeting(date: Date): string {
  const hour = date.getHours();
  if (hour >= 5 && hour < 12) {
    return 'Good morning';
  }
  if (hour >= 12 && hour < 18) {
    return 'Good afternoon';
  }
  return 'Good evening';
}

// First name = first whitespace-delimited token of displayName. Hyphenated names
// stay whole (a single first-name unit). Returns undefined when no usable name
// can be derived, so the greeting falls back to the bare time-of-day phrase.
//
// Returns `undefined` rather than the spec's literal `null` to comply with the
// rig's @rushstack/no-new-null rule; semantics are identical. Accepts `unknown`
// so a null/non-string pageContext value is handled without a null type.
export function extractFirstName(displayName: unknown): string | undefined {
  if (typeof displayName !== 'string') {
    return undefined;
  }
  const first = displayName.trim().split(/\s+/)[0];
  return first || undefined;
}
