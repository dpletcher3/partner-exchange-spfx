// Pure audience-matching predicate (no SPHttpClient dependency, unit-testable in
// isolation). This is the security-relevant gate, so every ambiguous case is
// FAIL-CLOSED (returns false → tile hidden):
//   - viewer has no resolved Division (non-partner / no profile row) → false
//   - tile allows no divisions (empty set) → false (shown to no one)
//   - otherwise: visible iff the viewer's single Division is in the allowed set.
export function isTileVisible(
  viewerDivision: string | undefined,
  allowedDivisions: string[]
): boolean {
  if (!viewerDivision) {
    return false;
  }
  if (!allowedDivisions || allowedDivisions.length === 0) {
    return false;
  }
  return allowedDivisions.indexOf(viewerDivision) !== -1;
}
