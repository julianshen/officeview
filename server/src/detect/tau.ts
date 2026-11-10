import { CANDIDATE_MATRIX_MAX_BYTES } from '../mark/detect'

/** §13.2 decision thresholds by declared candidate-table size M. A trace over
 * `count` candidates uses the next declared M at or above `count`. */
export const TAU_TABLE: ReadonlyArray<{ m: number; abs: number; margin: number; source: string }> = Object.freeze([
  // Fitted (Stage-0 `calibrate`, upper-decile Gaussian tail at 1e-6) over 2160 null trace() decisions at 2^16.
  { m: 2 ** 16, abs: 7.93, margin: 3.45, source: 'stage-2b calibrate 6e2129af' },
  // Stage-0 union rule from the fitted 2^16 value: sqrt(τ² + 2 ln(M / 2^16)).
  { m: 2 ** 20, abs: 8.27, margin: 3.45, source: 'stage-2b calibrate 6e2129af' },
  { m: 2 ** 24, abs: 8.6, margin: 3.45, source: 'stage-2b calibrate 6e2129af' },
].map((t) => Object.freeze(t)))

/** The most candidates one decision scores, derived from the detector's matrix bound
 * (⌊(2³¹ − 1) / 512⌋ = 4,194,303, so a single matrix at the maximum still fits); above it the tool refuses. */
export const MAX_CANDIDATES = Math.floor(CANDIDATE_MATRIX_MAX_BYTES / 512)

export function tauFor(count: number): { abs: number; margin: number; m: number } {
  if (!Number.isSafeInteger(count) || count < 0) throw new RangeError('Invalid candidate count')
  const t = TAU_TABLE.find((x) => x.m >= count)
  if (!t) throw new RangeError(`No declared threshold for ${count} candidates`)
  return { abs: t.abs, margin: t.margin, m: t.m }
}
