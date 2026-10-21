/**
 * The WebGL2 ↔ Canvas 2D frame comparator (slideshow plan S1-14), shared by
 * the GPU harness page (src/debug/slideshow-gpu.ts) and its Node test. Dev
 * only: src/debug is outside the library build.
 *
 * Tolerance (`Tolerance`, fractions of 255, spec 6.4/10):
 * - `meanAbs`: the mean absolute difference over every channel of every pixel;
 * - `rowMax`: the worst row's mean absolute difference over its channels.
 */
import type { Tolerance } from '../slideshow/types'

export function frameDiff(a: ArrayLike<number>, b: ArrayLike<number>, w: number, h: number): Tolerance {
  if (a.length !== w * h * 4 || b.length !== w * h * 4) throw new RangeError('frameDiff: frame size mismatch')
  let total = 0
  let worst = 0
  for (let y = 0; y < h; y++) {
    let row = 0
    for (let i = y * w * 4, end = i + w * 4; i < end; i++) row += Math.abs(a[i] - b[i])
    total += row
    worst = Math.max(worst, row / (w * 4))
  }
  return { meanAbs: total / (w * h * 4) / 255, rowMax: worst / 255 }
}

export function withinTolerance(d: Tolerance, t: Tolerance): boolean {
  return d.meanAbs <= t.meanAbs && d.rowMax <= t.rowMax
}
