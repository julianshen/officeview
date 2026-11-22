/**
 * The visible mark (§7.6): its spec, its styling and the one way the server paints it. Moved out of
 * `render.ts` (which re-exports it) so the show-plan builder can burn it into layers without a cycle.
 */
import { defaultMarkColor, fontSizePx, isWatermarkMuted, normalizeWatermark, paintWatermark, setWatermarkMeasurer, watermarkStamps, type WatermarkOptions } from '../../../src/core/watermark'

export const FIXED_MARK_TEXT = 'officeview · protected'
export type MarkSpec = { kind: 'fixed' } | { kind: 'viewer'; text: string } | { kind: 'off' }

/** Visible-mark styling [call]: mid-grey reads on light and dark pages alike. */
export function markOptions(mark: MarkSpec): WatermarkOptions | undefined {
  if (mark.kind === 'off') return undefined
  return { text: mark.kind === 'fixed' ? FIXED_MARK_TEXT : mark.text, color: '#808080', opacity: 0.18 }
}

/** Paints the visible mark from a clean watermark module state.
 *
 * `src/core/watermark.ts` keeps a process-global text-width cache keyed on the
 * ROUNDED point size, but (with no measurer registered, as here) stores a
 * width computed from the UNROUNDED pixel size. Two pages whose sizes round to
 * the same pt (a 960-px deck and a 955-px sheet are both 50 pt) then share one
 * entry, and whichever painted first sets the tile step of the other: the
 * canonical a long-lived worker serves would depend on what it rendered
 * before, and a fresh trace process would refuse the reference (F1, Stage 2b).
 * `setWatermarkMeasurer(null)` is the module's only reset (measurer stays
 * null, cache cleared), so every mark is laid out from its own page alone. */
export function paintMark(ctx: CanvasRenderingContext2D, page: { widthPx: number; heightPx: number }, mark: WatermarkOptions): void {
  setWatermarkMeasurer(null)
  paintWatermark(ctx, page, mark)
}

/** [call] The periodic mark's font, when the mark names no size: the size a 512-CSS-px page gets (7 % of its width,
 * as `paintWatermark` sizes a page), i.e. the Task 1 spike's tile at scale 1. A sheet is unbounded, so its own width
 * cannot size the mark. */
export const MARK_REGION_PAGE_CSS = 512
/** Stamp device positions are snapped to 1/64 px, so a stamp's offset inside any region is exact in float32 and two
 * regions paint identical pixels where they meet. */
const SNAP = 64
const snap = (v: number) => Math.round(v * SNAP) / SNAP

/**
 * Stage 4a (S4 §6.5): the visible mark of a sheet, laid out in global sheet space and painted into one region of it.
 * The layout is the `'tile'` placement's periodic lattice (its steps from `watermarkStamps`, whatever the mark's
 * placement), anchored at the sheet's origin, in logical CSS px scaled by `region.scale` (zoom × dpr). `region` is in
 * the sheet's device px; the context's device origin is the region's top-left, and nothing outside `w × h` is painted.
 * A tile is the 512² region at `(col·512, row·512)`, so neighbouring tiles show adjoining parts of one mark. Stamps
 * paint in one fixed order (row, then column), whatever the region, and each stamp's device offset is snapped (SNAP).
 */
export function paintMarkRegion(ctx: CanvasRenderingContext2D, mark: WatermarkOptions, region: { x: number; y: number; w: number; h: number; scale: number }): void {
  const resolved = normalizeWatermark(mark)
  if (!resolved || isWatermarkMuted()) return
  const s = region.scale
  if (!(s > 0) || !Number.isFinite(s) || !(region.w > 0) || !(region.h > 0)) return
  const m = { ...resolved, placement: 'tile' as const, fontSizePt: resolved.fontSizePt ?? Math.round((fontSizePx({ widthPx: MARK_REGION_PAGE_CSS, heightPx: MARK_REGION_PAGE_CSS }, resolved) * 72) / 96) }
  // A clean width cache, as `paintMark` (the steps below measure the text through it).
  setWatermarkMeasurer(null)
  const fontPx = fontSizePx({ widthPx: 0, heightPx: 0 }, m)
  // The lattice steps, read off `watermarkStamps` (3 × 3 stamps on an empty page: x = -stepX, 0, stepX; y likewise).
  const probe = watermarkStamps({ widthPx: 0, heightPx: 0 }, m)
  const stepX = probe[1]!.xPx - probe[0]!.xPx, stepY = probe[3]!.yPx - probe[0]!.yPx
  if (!(stepX > 0) || !(stepY > 0)) return
  // A stamp reaches at most its text's half-diagonal from its centre; `reach` bounds that generously.
  const reach = stepX + fontPx
  const lx0 = region.x / s - reach, lx1 = (region.x + region.w) / s + reach
  const ly0 = region.y / s - reach, ly1 = (region.y + region.h) / s + reach
  const j0 = Math.floor(ly0 / stepY), j1 = Math.ceil(ly1 / stepY)
  const i0 = Math.floor(lx0 / stepX) - 1, i1 = Math.ceil(lx1 / stepX)
  const rad = (m.rotate * Math.PI) / 180, cos = Math.cos(rad), sin = Math.sin(rad)
  ctx.save()
  try {
    ctx.setTransform(1, 0, 0, 1, 0, 0)
    ctx.beginPath()
    ctx.rect(0, 0, region.w, region.h)
    ctx.clip()
    ctx.globalAlpha = m.opacity
    // An invalid colour is ignored by the canvas, leaving the fallback assigned first.
    ctx.fillStyle = defaultMarkColor()
    if (m.color) ctx.fillStyle = m.color
    ctx.font = `${m.fontSizePt}pt ${quoteFamily(m.fontFamily)}`
    ctx.textAlign = 'center'
    ctx.textBaseline = 'middle'
    for (let j = j0; j <= j1; j++) {
      const odd = ((j % 2) + 2) % 2 === 1
      for (let i = i0; i <= i1; i++) {
        const dx = snap(s * (i + (odd ? 0.5 : 0)) * stepX) - region.x, dy = snap(s * j * stepY) - region.y
        ctx.setTransform(s * cos, s * sin, -s * sin, s * cos, dx, dy)
        ctx.fillText(m.text, 0, 0)
      }
    }
  } finally { ctx.restore() }
}

/** A font family list for the CSS font shorthand, quoted as `paintWatermark` quotes it. */
function quoteFamily(family: string): string {
  return family.split(',').map((f) => f.trim()).filter((f) => f.length > 0).map((f) => (/^[A-Za-z][A-Za-z0-9 ]*$/.test(f) ? f : `"${f}"`)).join(', ')
}
