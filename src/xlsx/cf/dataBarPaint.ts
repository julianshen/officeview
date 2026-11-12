/**
 * Data-bar painting (Stage 3, Task 22): one shared painter for the static, interactive (resident and paged), print and
 * export views. Lives in the lazy `cf` chunk; eager renderers reach it only through `CfPaintBatch.barAt`/`paintBars`.
 *
 * - The geometry is Excel's measured shape at 1 px/pt (copy picture, docs/validation/xlsx-cf.md §11), pinned in
 *   `CF_BAR_RENDER_MODEL` and gated by its `bar:*` capabilities (Rulings S3-M, S3-N). A value the captures did not
 *   measure (gradient stops, the minimum visible width) is absent from the model, and the branch that needs it withholds.
 * - Other zooms and DPRs scale the measured point geometry and snap each edge to the sheet-global device grid, as
 *   sparklines do (`sparklines/geometry.ts`), so a tile and the full render carry identical rects (OD-8: fidelity there
 *   is unverified, the fractions stay exact).
 * - Pure layout plus a small canvas replay: no DOM, no clock, no global state beyond the evidence lookup.
 */
import type { CfBarDescriptor } from './model'
import { hasEvidence as defaultHasEvidence } from './evidence'
import type { Rect } from '../view/textLayout'

/** 1 pt in logical (CSS) px: the viewer's convention (`sparklines/geometry.ts`). At zoom × DPR = 3/4 a point is one
 * device px, which is the scale Excel's captures were measured at. */
const PT_PX = 4 / 3

export interface CfBarRenderModel {
  /** `bar:geometry-insets`: the bar spans [left, right) × [top, bottom) of the host rect inset by these points (the
   * rect's first column and row are its own gridline; right/bottom count from the next gridline). */
  readonly insetsPt: Readonly<{ left: number; right: number; top: number; bottom: number }>
  /** `bar:geometry-snapping`: lengths are `fraction × span` rounded half up to whole device px. */
  readonly rounding: 'half-up'
  /** `bar:render-border`: the border is a frame of this width inside the bar rect. */
  readonly borderPt: number
  /** The axis is one column this wide over the host's full height (addendum §2; Excel's dash is not reproduced). */
  readonly axisPt: number
  /** `bar:geometry-min-visible-width`: unmeasured (non-discriminating), so absent; a non-zero length that rounds to
   * nothing withholds. */
  readonly minVisiblePt?: number
  /** `bar:render-gradient-stops`: unmeasured, so absent; a gradient bar withholds (Ruling S3-M). */
  readonly gradientStops?: readonly Readonly<{ at: number; argb: number }>[]
}

/** Values from the Stage 3 captures (refs/stage3-bar-geometry.json, refs/stage3-bar-rendering.json; pinned by
 * tests/xlsx-cf-bar-render.test.ts): 2 pt in from the left gridline, 1 pt short of the next one, 2 pt below the top
 * gridline and above the bottom one; a 1 pt border frame; a 1 pt axis column. */
export const CF_BAR_RENDER_MODEL: CfBarRenderModel = Object.freeze({
  insetsPt: Object.freeze({ left: 2, right: 1, top: 2, bottom: 2 }),
  rounding: 'half-up',
  borderPt: 1,
  axisPt: 1,
})

/** The cell-independent paint gates: the measured inset model, its rounding and its clip to narrow hosts. "Narrow"
 * has no measured threshold, so `bar:geometry-narrow-clip` gates every bar rather than a guessed width class. */
export const CF_BAR_PAINT_CAPABILITIES: readonly string[] = Object.freeze(['bar:geometry-insets', 'bar:geometry-snapping', 'bar:geometry-narrow-clip'])

/** Whether a descriptor can paint at all, whatever the cell: the paint gates hold and the model covers its fill.
 * `barAt` of every batch answers through this, so a withheld bar never suppresses text nor announces itself. */
export function barPaintable(bar: CfBarDescriptor | undefined, model: CfBarRenderModel = CF_BAR_RENDER_MODEL, check: (cap: string) => boolean = defaultHasEvidence): bar is CfBarDescriptor {
  if (!bar) return false
  for (const cap of CF_BAR_PAINT_CAPABILITIES) if (!check(cap)) return false
  return !bar.isGradient || model.gradientStops !== undefined
}

/** S3-O: text is hidden only under a painted bar whose `showValue` is false, and only with evidence. */
export function barHidesText(bar: CfBarDescriptor | undefined, check: (cap: string) => boolean = defaultHasEvidence): boolean {
  return barPaintable(bar, CF_BAR_RENDER_MODEL, check) && bar.showValue === false && check('bar:showvalue-canvas-suppression')
}

/** One axis-aligned rect in output units; colours stay ARGB. */
export type CfBarRect = Readonly<{ x: number; y: number; w: number; h: number; argb: number }>
/** What one bar paints, in paint order: the fill, its border frame, then the axis. Empty when the host is too small. */
export type CfBarPrimitives = Readonly<{ fill?: CfBarRect; border?: readonly CfBarRect[]; axis?: CfBarRect }>
export type CfBarScale = Readonly<{ zoom: number; dpr: number; output?: 'logical' | 'device' }>

const roundHalfUp = (x: number): number => Math.floor(x + 0.5)

/**
 * Lays one bar out in its host rect (logical px). Output is logical px snapped to the device grid, or device px on the
 * sheet-global grid with `scale.output === 'device'`. Undefined when a model value it needs is unmeasured or a gate
 * lacks evidence (withheld, never approximated).
 *
 * Lengths follow the measured shape: with no axis drawn the bar grows from the leading edge (or, all-negative, from
 * the trailing one) as `round(share × span)`; with an axis, the axis takes `axisPt` at `round(axisFraction × span)`
 * and each side is its own span: the positive side `span − axis`, the negative side `axis − axisPt` (Ruling S3-B:
 * axis layouts are `lengthBasis: 'side'`). Right-to-left mirrors the whole layout inside the span.
 */
export function layoutDataBar(
  bar: CfBarDescriptor,
  cell: Rect,
  scale: CfBarScale,
  model: CfBarRenderModel = CF_BAR_RENDER_MODEL,
  check: (cap: string) => boolean = defaultHasEvidence,
): CfBarPrimitives | undefined {
  if (!barPaintable(bar, model, check)) return undefined
  if (bar.borderColor !== undefined && !check('bar:render-border')) return undefined
  for (const n of [bar.axisFraction, bar.startFraction, bar.endFraction]) if (!Number.isFinite(n) || n < 0 || n > 1) return undefined
  const zoom = scale.zoom > 0 && Number.isFinite(scale.zoom) ? scale.zoom : 1
  const dpr = scale.dpr > 0 && Number.isFinite(scale.dpr) ? scale.dpr : 1
  const devPerPx = zoom * dpr, devPerPt = devPerPx * PT_PX
  const dev = (pt: number): number => roundHalfUp(pt * devPerPt)
  // Host edges on the sheet-global device grid; the measured insets in whole device px.
  const X0 = roundHalfUp(cell.x * devPerPx), X1 = roundHalfUp((cell.x + cell.width) * devPerPx)
  const Y0 = roundHalfUp(cell.y * devPerPx), Y1 = roundHalfUp((cell.y + cell.height) * devPerPx)
  const L = X0 + dev(model.insetsPt.left), R = X1 - dev(model.insetsPt.right)
  const T = Y0 + dev(model.insetsPt.top), B = Y1 - dev(model.insetsPt.bottom)
  const S = R - L
  if (!(S > 0) || !(B > T) || !(X1 > X0)) return {}
  // A non-zero share that rounds to no pixel is the unmeasured minimum-visible-width branch.
  const lengthOf = (share: number, span: number): number | undefined => {
    const n = roundHalfUp(share * span)
    if (n > 0 || !(share > 0)) return n
    return model.minVisiblePt === undefined ? undefined : Math.min(span, Math.max(1, dev(model.minVisiblePt)))
  }
  let a: number, b: number
  let axis: [number, number] | undefined
  if (bar.axisColor === undefined) {
    // No axis line: a leading-edge bar (axisFraction 0) or the all-negative trailing-edge bar (axisFraction 1).
    if (bar.axisFraction !== 0 && bar.axisFraction !== 1) return undefined
    const len = lengthOf(bar.endFraction - bar.startFraction, S)
    if (len === undefined) return undefined
    ;[a, b] = bar.axisFraction === 0 ? [0, len] : [S - len, S]
  } else {
    const axisW = Math.max(1, dev(model.axisPt))
    const sn = Math.min(S, Math.max(axisW, roundHalfUp(bar.axisFraction * S)))
    axis = [sn - axisW, sn]
    if (!bar.isNegative) {
      const side = S - sn, share = bar.axisFraction < 1 ? (bar.endFraction - bar.axisFraction) / (1 - bar.axisFraction) : 0
      const len = lengthOf(share, side)
      if (len === undefined) return undefined
      ;[a, b] = [sn, sn + len]
    } else {
      const side = sn - axisW, share = bar.axisFraction > 0 ? (bar.axisFraction - bar.startFraction) / bar.axisFraction : 0
      const len = lengthOf(share, side)
      if (len === undefined) return undefined
      ;[a, b] = [side - len, side]
    }
  }
  if (bar.direction === 'rightToLeft') {
    ;[a, b] = [S - b, S - a]
    if (axis) axis = [S - axis[1], S - axis[0]]
  }
  const out = scale.output === 'device' ? (v: number) => v : (v: number) => v / devPerPx
  const rect = (x: number, y: number, w: number, h: number, argb: number): CfBarRect => ({ x: out(x), y: out(y), w: out(w), h: out(h), argb })
  const primitives: { fill?: CfBarRect; border?: CfBarRect[]; axis?: CfBarRect } = {}
  if (b > a) {
    const x = L + a, w = b - a, h = B - T
    primitives.fill = rect(x, T, w, h, bar.fillColor)
    if (bar.borderColor !== undefined) {
      const bw = Math.max(1, dev(model.borderPt))
      const side = Math.min(bw, w), cap = Math.min(bw, h)
      primitives.border = [
        rect(x, T, w, cap, bar.borderColor), rect(x, B - cap, w, cap, bar.borderColor),
        rect(x, T, side, h, bar.borderColor), rect(x + w - side, T, side, h, bar.borderColor),
      ]
    }
  }
  if (axis && bar.axisColor !== undefined) primitives.axis = rect(L + axis[0], Y0, axis[1] - axis[0], Y1 - Y0, bar.axisColor)
  return primitives
}

const css = (argb: number): string => `#${(argb & 0xffffff).toString(16).padStart(6, '0')}`

/** The bars a batch holds, keyed `row * 16384 + col` in source coordinates. */
export type CfBarCells = ReadonlyMap<number, CfBarDescriptor>

/**
 * Paints every paintable bar of `bars` inside `visible` whose host `cellRect` yields, clipped to that host. Runs in the
 * fill phase, before text. Bounded work: it walks the smaller of the bar set and the visible window. Device output
 * snaps on the sheet-global device grid and applies the canvas translation afterwards, so tiles agree with the full
 * render. Returns true (there is no failure mode once loaded).
 */
export function paintDataBars(
  ctx: CanvasRenderingContext2D,
  bars: CfBarCells,
  barAt: (row: number, col: number) => CfBarDescriptor | undefined,
  cellRect: (row: number, col: number) => Rect | undefined,
  visible: { r0: number; c0: number; r1: number; c1: number },
  scale: CfBarScale,
  check: (cap: string) => boolean = defaultHasEvidence,
): boolean {
  if (bars.size === 0) return true
  const paintOne = (row: number, col: number): void => {
    const bar = barAt(row, col)
    if (!bar) return
    const host = cellRect(row, col)
    if (!host || !(host.width > 0 && host.height > 0)) return
    const deviceSpace = scale.output === 'device'
    const primitives = layoutDataBar(bar, host, scale, CF_BAR_RENDER_MODEL, check)
    if (!primitives) return
    const transform = deviceSpace ? ctx.getTransform() : undefined
    ctx.save()
    try {
      ctx.beginPath(); ctx.rect(host.x, host.y, host.width, host.height); ctx.clip()
      if (transform) ctx.setTransform(1, 0, 0, 1, transform.e, transform.f)
      for (const r of [primitives.fill, ...(primitives.border ?? []), primitives.axis]) {
        if (!r) continue
        ctx.fillStyle = css(r.argb)
        ctx.fillRect(r.x, r.y, r.w, r.h)
      }
    } finally { ctx.restore() }
  }
  const { r0, c0, r1, c1 } = visible
  const area = (r1 - r0 + 1) * (c1 - c0 + 1)
  if (Number.isFinite(area) && area > 0 && area <= bars.size) {
    for (let row = r0; row <= r1; row++) for (let col = c0; col <= c1; col++) if (bars.has(row * 16384 + col)) paintOne(row, col)
  } else {
    for (const key of bars.keys()) {
      const row = Math.floor(key / 16384), col = key % 16384
      if (row >= r0 && row <= r1 && col >= c0 && col <= c1) paintOne(row, col)
    }
  }
  return true
}
