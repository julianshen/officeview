/**
 * Time manipulations (slideshow spec 6.2; SMIL timing manipulations, as
 * PresentationML `accel`, `decel` and `autoRev`).
 */

/**
 * Maps a position `p` (0–1) through one iteration to eased local progress.
 *
 * With `autoRev`, `p` covers the doubled iteration: the first half plays
 * forward and the second half mirrors it, so the ease applies to the forward
 * simple duration and the reverse retraces it.
 *
 * The rate ramps linearly from 0 over `accel`, holds at
 * `r = 1 / (1 - accel/2 - decel/2)`, then ramps back to 0 over `decel`, so
 * progress is continuous and ends at 1. Hostile values are tamed: non-finite
 * or negative fractions count as 0, and a sum above 1 is scaled down to 1.
 */
export function easeProgress(p: number, accel: number, decel: number, autoRev: boolean): number {
  let q = p > 0 ? (p < 1 ? p : 1) : 0 // also maps NaN to 0
  if (autoRev) q = q <= 0.5 ? 2 * q : 2 - 2 * q
  let a = accel > 0 && Number.isFinite(accel) ? accel : 0
  let b = decel > 0 && Number.isFinite(decel) ? decel : 0
  if (a + b > 1) {
    const s = a + b
    a /= s
    b /= s
  }
  if (a === 0 && b === 0) return q
  const r = 1 / (1 - a / 2 - b / 2)
  if (q < a) return (r * q * q) / (2 * a)
  if (q <= 1 - b) return r * (q - a / 2)
  const u = 1 - q
  return 1 - (r * u * u) / (2 * b)
}
