/**
 * Deterministic clock for the real-browser harness (S0.3).
 *
 * The default clock reads performance.now() (real measurements stay real).
 * Tests and deterministic scenarios inject a fake now() so marker math is
 * reproducible without weakening the real-browser path.
 */
export interface DeterministicClock {
  now(): number
}

export function createDeterministicClock(now?: () => number): DeterministicClock {
  return { now: now ?? (() => performance.now()) }
}
