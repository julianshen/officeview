/**
 * Pure 2-D affine helpers (piece 4 D17): the slide-px matrix algebra the hit
 * test, the Morph renderer and the package-free playback modules share.
 * Package-free and DOM-free: types only.
 */
import type { Affine, TargetState } from './types'

const IDENTITY: Affine = [1, 0, 0, 1, 0, 0]

/** `m1 ∘ m2`: apply `m2`, then `m1`. */
export function mul(m1: Affine, m2: Affine): Affine {
  const [a, b, c, d, e, f] = m1, [A, B, C, D, E, F] = m2
  return [a * A + c * B, b * A + d * B, a * C + c * D, b * C + d * D, a * E + c * F + e, b * E + d * F + f]
}

export function invert(m: Affine): Affine | undefined {
  const [a, b, c, d, e, f] = m
  const det = a * d - b * c
  if (!Number.isFinite(det) || Math.abs(det) < 1e-12) return undefined
  return [d / det, -b / det, -c / det, a / det, (c * f - d * e) / det, (b * e - a * f) / det]
}

export function apply(m: Affine, x: number, y: number): [number, number] {
  return [m[0] * x + m[2] * y + m[4], m[1] * x + m[3] * y + m[5]]
}

/** A target's draw matrix in slide px: `translate(motion) ∘ transform` (the compositor contract). */
export function stateMatrix(s: TargetState | undefined): Affine {
  if (!s) return IDENTITY
  const t = s.transform.every(Number.isFinite) ? s.transform : IDENTITY
  const dx = s.motion && Number.isFinite(s.motion.dx) ? s.motion.dx : 0
  const dy = s.motion && Number.isFinite(s.motion.dy) ? s.motion.dy : 0
  return dx || dy ? mul([1, 0, 0, 1, dx, dy], t) : t
}
