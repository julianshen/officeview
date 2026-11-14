/**
 * Marks on show-plan assets (piece 4 spec §7, plan D7, D8, D9).
 *
 * - **Visible mark.** Every layer asset carries the page's visible mark, burned `source-atop`: drawn only over
 *   the layer's own ink, so its transparent padding stays alpha 0 and its alpha is unchanged. The mark is the
 *   PAGE's mark, translated so that it lines up with the page when the layer sits at its rest (owner decision
 *   1): a mover's mark travels with it during the motion and is exact at every rest point.
 * - **Forensic mark.** Layer-1 is applied per session at send time (`markLayerRaster`, the page basis cropped
 *   to the layer's rest bounds).
 * - **Validation (§7, D9).** At every step boundary the layer composite must equal a reference composed once and
 *   marked once, within `MARK_FIDELITY`, and the validation code must be attributed on it (D7). A slide that
 *   fails plays every step `composited`.
 */
import { canvasBackend } from '../../../src/core/canvas-backend'
import { compareFrames, type CompareResult } from '../../../src/slideshow/compare'
import type { LayerBitmap, Tolerance } from '../../../src/slideshow/types'
import type { Rect } from '../../../src/stream/show-plan'
import { markOptions, paintMark, type MarkSpec } from '../mark/visible'

/** D9: the layer composite against the once-marked reference, in fractions of 255. */
export const MARK_FIDELITY: Readonly<Tolerance> = Object.freeze({ meanAbs: 0.004, rowMax: 0.02 })

/**
 * `layer` with the page's visible mark (`page` slide px) burned `source-atop`, the page translated by `-rest`
 * (`rest`: the layer's bounds at its rest position, slide px). A new bitmap; the input is not touched. With no
 * mark (`off`) the layer itself is returned.
 */
export function burnVisibleMark(layer: LayerBitmap, rest: Rect, mark: MarkSpec, page: { w: number; h: number }): LayerBitmap {
  const opts = markOptions(mark)
  if (!opts) return layer
  const w = layer.canvas.width, h = layer.canvas.height
  const canvas = canvasBackend().createCanvas(w, h) as unknown as HTMLCanvasElement
  const ctx = canvas.getContext('2d') as unknown as CanvasRenderingContext2D
  ctx.drawImage(layer.canvas as CanvasImageSource, 0, 0)
  ctx.globalCompositeOperation = 'source-atop'
  // Layer px per slide px of the rest box on each axis (a layer resting scaled covers more slide px); the layer's
  // px (0, 0) is the rest box's top-left.
  const kx = w / rest.w, ky = h / rest.h
  ctx.setTransform(kx, 0, 0, ky, -rest.x * kx, -rest.y * ky)
  paintMark(ctx, { widthPx: page.w, heightPx: page.h }, opts)
  return { canvas, bounds: { ...layer.bounds }, scale: layer.scale, bytes: w * h * 4 }
}

/** What `validateMarks` checks over: the bundle's plan (its step count). */
export interface ValidatedPlan { plan: { steps: readonly unknown[] } }

/** The visible mark's own contribution at boundary k: each side composed with and without its visible mark. */
export interface MarkContribution {
  composeMarked(step: number): ImageData
  composeUnmarked(step: number): ImageData
  referenceMarked(step: number): ImageData
  referenceUnmarked(step: number): ImageData
}

/**
 * Task 6 follow-up (controller ruling): the visible mark's own contribution, `composite marked − composite
 * unmarked` against `reference marked − reference unmarked`, compared over the mark's ink region only (every
 * pixel either contribution touches), with the `MARK_FIDELITY` values. A whole-frame comparison dilutes a
 * misplaced mark over the page (a mover's mark 240 px off passes it); over the ink region it does not.
 */
export function compareMarkContribution(cm: ImageData, cu: ImageData, rm: ImageData, ru: ImageData): CompareResult {
  const frames = [cm, cu, rm, ru]
  if (frames.some(f => f.data.length !== f.width * f.height * 4)) return { pass: false, meanAbs: 1, rowMax: 1, worstRow: -1, reason: 'malformed' }
  if (frames.some(f => f.width !== cm.width || f.height !== cm.height)) return { pass: false, meanAbs: 1, rowMax: 1, worstRow: -1, reason: 'size-mismatch' }
  const { width, height } = cm
  let total = 0, count = 0, rowMax = 0, worstRow = 0
  for (let y = 0; y < height; y++) {
    let row = 0, n = 0
    for (let i = y * width * 4, end = i + width * 4; i < end; i += 4) {
      let ink = false, d = 0
      for (let c = 0; c < 4; c++) {
        const a = cm.data[i + c]! - cu.data[i + c]!, b = rm.data[i + c]! - ru.data[i + c]!
        if (a !== 0 || b !== 0) ink = true
        d = Math.max(d, Math.abs(a - b))
      }
      if (!ink) continue
      row += d; n++
    }
    if (!n) continue
    total += row; count += n
    const mean = row / n / 255
    if (mean > rowMax) { rowMax = mean; worstRow = y }
  }
  const meanAbs = count ? total / count / 255 : 0
  return { pass: meanAbs <= MARK_FIDELITY.meanAbs && rowMax <= MARK_FIDELITY.rowMax, meanAbs, rowMax, worstRow }
}

/**
 * §7 validation over every step boundary `0..steps`:
 * - `compose(k)` (the shipped layers composed at boundary k) against `reference(k)` (the slide composed once and
 *   marked once) within `MARK_FIDELITY`, over the whole frame;
 * - with `contribution`, the visible mark's own contribution over its ink region (`compareMarkContribution`);
 * - with `attributes(k, composite)`, the validation code is attributed on the composite (D7, `detect`).
 * `worst` is the boundary with the largest mean difference, or the first failure.
 */
export function validateMarks(plan: ValidatedPlan, reference: (step: number) => ImageData, compose: (step: number) => ImageData, opts: { attributes?: (step: number, composite: ImageData) => boolean; contribution?: MarkContribution } = {}): { ok: boolean; worst: CompareResult } {
  let worst: CompareResult = { pass: true, meanAbs: 0, rowMax: 0, worstRow: 0 }
  for (let k = 0; k <= plan.plan.steps.length; k++) {
    const composite = compose(k)
    const r = compareFrames(composite, reference(k), MARK_FIDELITY)
    if (!r.pass) return { ok: false, worst: r }
    if (r.meanAbs >= worst.meanAbs) worst = r
    if (opts.contribution) {
      const m = opts.contribution
      const c = compareMarkContribution(m.composeMarked(k), m.composeUnmarked(k), m.referenceMarked(k), m.referenceUnmarked(k))
      if (!c.pass) return { ok: false, worst: c }
    }
    if (opts.attributes && !opts.attributes(k, composite)) return { ok: false, worst: { ...r, pass: false } }
  }
  return { ok: true, worst }
}
