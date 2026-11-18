/**
 * Frame comparison shared by the slideshow compare harness (`scripts/slideshow-compare.ts`, which re-exports
 * it) and the stream server's show-plan mark validation (piece 4 D9). DOM-free.
 *
 * Units: every difference and tolerance is a fraction of 255 (0..1), the unit of the shared `Tolerance` type.
 */
import type { Tolerance } from './types'

export interface CompareResult {
  pass: boolean
  /** Mean over all pixels of the pixel difference (fraction of 255). */
  meanAbs: number
  /** The largest row mean of the pixel difference (fraction of 255). */
  rowMax: number
  /** The row holding `rowMax`; -1 when the frames could not be compared. */
  worstRow: number
  /** Set when the frames could not be compared at all; such a result never passes. */
  reason?: 'size-mismatch' | 'empty' | 'malformed'
}

const refused = (reason: NonNullable<CompareResult['reason']>): CompareResult => ({ pass: false, meanAbs: 1, rowMax: 1, worstRow: -1, reason })

/**
 * Compare two RGBA frames. The pixel difference is the largest absolute
 * difference over the R, G, B and A channels, divided by 255, so a
 * single-channel or alpha error is not averaged away. `meanAbs` is its mean
 * over the frame; `rowMax` is the largest per-row mean (the spec's "threshold
 * per row": a defect confined to one row is caught even when the frame mean
 * hides it; a small block in a wide row is diluted by the row width).
 *
 * Frames must have identical, non-zero dimensions: the reference resolution is
 * fixed (tests/fixtures/slideshow-refs/README.md), so nothing is ever resized.
 * A size mismatch, an empty frame or a buffer that does not match its
 * dimensions fails with a `reason`.
 */
export function compareFrames(a: ImageData, b: ImageData, tol: Tolerance): CompareResult {
  for (const f of [a, b]) if (f.data.length !== f.width * f.height * 4) return refused('malformed')
  if (a.width !== b.width || a.height !== b.height) return refused('size-mismatch')
  if (a.width === 0 || a.height === 0) return refused('empty')
  const { width, height } = a
  const da = a.data, db = b.data
  let total = 0, rowMax = 0, worstRow = 0
  for (let y = 0; y < height; y++) {
    let row = 0
    for (let i = y * width * 4, end = i + width * 4; i < end; i += 4) {
      row += Math.max(Math.abs(da[i] - db[i]), Math.abs(da[i + 1] - db[i + 1]), Math.abs(da[i + 2] - db[i + 2]), Math.abs(da[i + 3] - db[i + 3]))
    }
    total += row
    const mean = row / width / 255
    if (mean > rowMax) { rowMax = mean; worstRow = y }
  }
  const meanAbs = total / (width * height) / 255
  return { pass: meanAbs <= tol.meanAbs && rowMax <= tol.rowMax, meanAbs, rowMax, worstRow }
}
