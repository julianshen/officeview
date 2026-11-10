/** §13.2 decision thresholds by declared candidate-table size M. A trace over
 * `count` candidates uses the next declared M at or above `count`. */
export const TAU_TABLE: ReadonlyArray<{ m: number; abs: number; margin: number; source: string }> = Object.freeze([
  { m: 2 ** 16, abs: 7.99, margin: 3.85, source: 'stage-0 provisional' },
  // Gaussian-tail extrapolation from 2^16: +2 ln(2^20 / 2^16) under the square.
  { m: 2 ** 20, abs: Math.sqrt(7.99 ** 2 + 2 * Math.log(16)), margin: 3.85, source: 'stage-0 provisional' },
  // The fitted-tail value (Stage 0's τ ≈ 8.66 at 2^24).
  { m: 2 ** 24, abs: 8.66, margin: 3.85, source: 'stage-0 provisional' },
].map((t) => Object.freeze(t)))

/** The most candidates one decision scores (the detector's matrix bound); above it the tool refuses. */
export const MAX_CANDIDATES = 2 ** 22

export function tauFor(count: number): { abs: number; margin: number; m: number } {
  if (!Number.isSafeInteger(count) || count < 0) throw new RangeError('Invalid candidate count')
  const t = TAU_TABLE.find((x) => x.m >= count)
  if (!t) throw new RangeError(`No declared threshold for ${count} candidates`)
  return { abs: t.abs, margin: t.margin, m: t.m }
}
