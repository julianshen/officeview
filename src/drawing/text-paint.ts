import type { FontResolver } from '../core/fonts/register'
import { createLocalTextMeasurer } from '../core/text-metrics'
import { graphemes, RECORD_TEXT, type LogicalTextRange, type LogicalTextSource, type TextRecordingContext } from '../core/text-recording'
import type { ThemeContext } from './style'
import { layoutTextBody, resolveTextFamily, type MeasureText, type TextLayout } from './text-layout'
import type { DrawingTextBody as PptxTextBody, DrawingTextStyle as PptxTextStyle } from './text'
import { buildWarpMapping, buildWarpLattice, type WarpMapping } from './text-warp'
import { scratchSurface, type PaintSurface } from './paint'
import { getPatternTile, evictPatternTile } from './pattern'
export { paintPatternTile } from './pattern'

/** Without native spacing, preserve contextual shaping rather than drawing
 * isolated Arabic/Indic clusters. Requested source tracking remains in the model. */
export function needsContextualShaping(text: string): boolean {
  return /[\p{Script=Arabic}\p{Script=Syriac}\p{Script=Mongolian}\p{Script=Devanagari}\p{Script=Bengali}\p{Script=Gujarati}\p{Script=Gurmukhi}\p{Script=Kannada}\p{Script=Malayalam}\p{Script=Oriya}\p{Script=Tamil}\p{Script=Telugu}\p{Script=Sinhala}\p{Script=Khmer}\p{Script=Myanmar}\p{Script=Thai}]/u.test(text)
}

/** Blink suppresses spacing for default-ignorable/zero-width formatting
 * characters after Canvas whitespace preparation (CR/FF become spaces).
 * Attached marks/joiners remain part of their visible cluster. */
export function trackingEligible(cluster: string): boolean {
  return !/^[\p{Default_Ignorable_Code_Point}\ufffc]+$/u.test(cluster)
}

// WHATWG Canvas text preparation normalizes ASCII whitespace before shaping.
// Keep this paint string separate from the source recorded for search/copy.
const canvasText = (text: string): string => text.replace(/[\t\n\f\r]/g, ' ')

export interface AppearanceBox { x: number; y: number; width: number; height: number }

function resolveTextFill(
  ctx: CanvasRenderingContext2D,
  style: PptxTextStyle,
  box: AppearanceBox,
  patternCache: Map<string, CanvasPattern | null>,
): string | CanvasGradient | CanvasPattern {
  const fill = style.textFill
  if (fill?.kind === 'gradient' && fill.stops.length >= 2) {
    const angle = Number.isFinite(fill.angle) ? fill.angle : 0
    const dx = Math.cos(angle), dy = Math.sin(angle)
    const cx = box.x + box.width / 2, cy = box.y + box.height / 2
    const center = cx * dx + cy * dy
    const dots = [box.x, box.x + box.width].flatMap(px => [box.y, box.y + box.height].map(py => px * dx + py * dy))
    const lo = Math.min(...dots) - center, hi = Math.max(...dots) - center
    const x0 = Number.isFinite(cx + dx * lo) ? cx + dx * lo : box.x
    const y0 = Number.isFinite(cy + dy * lo) ? cy + dy * lo : box.y
    const x1 = Number.isFinite(cx + dx * hi) ? cx + dx * hi : box.x + box.width
    const y1 = Number.isFinite(cy + dy * hi) ? cy + dy * hi : box.y
    const g = ctx.createLinearGradient(x0, y0, x1, y1)
    for (const s of fill.stops) g.addColorStop(Math.min(1, Math.max(0, s.position)), s.color)
    return g
  }
  if (fill?.kind === 'pattern') {
    const key = `${fill.preset}\n${fill.fg}\n${fill.bg}`
    let pat = patternCache.get(key)
    if (pat === undefined) {
      const tile = getPatternTile(ctx, fill.preset, fill.fg, fill.bg)
      try {
        pat = tile ? ctx.createPattern(tile, 'repeat') : null
      } catch {
        pat = null
        evictPatternTile(fill.preset, fill.fg, fill.bg)
      }
      patternCache.set(key, pat)
    }
    if (pat) return pat
    return fill.fg
  }
  return style.color ?? '#000000'
}
/**
 * Shadow state is assigned on every painted segment (transparent default),
 * so a shadowed run can never leak into its neighbors. Offsets follow canvas
 * CTM semantics in rotated frames; Word-parity of shadow direction there is a
 * validation item, not asserted here.
 */
export const MAX_SHADOW_BLUR_PX = 100
export const MAX_SHADOW_OFFSET_PX = 200
export const MAX_OUTLINE_WIDTH_PX = 100

function applyTextShadow(ctx: CanvasRenderingContext2D, style: PptxTextStyle): void {
  const sh = style.textShadow
  if (!sh) {
    ctx.shadowColor = 'rgba(0,0,0,0)'
    ctx.shadowBlur = 0
    ctx.shadowOffsetX = 0
    ctx.shadowOffsetY = 0
    return
  }
  ctx.shadowColor = sh.color
  const blur = Number.isFinite(sh.blurPx) ? sh.blurPx : 0
  const offX = Number.isFinite(sh.offsetX) ? sh.offsetX : 0
  const offY = Number.isFinite(sh.offsetY) ? sh.offsetY : 0
  ctx.shadowBlur = Math.min(Math.max(0, blur), MAX_SHADOW_BLUR_PX)
  ctx.shadowOffsetX = Math.min(Math.max(-MAX_SHADOW_OFFSET_PX, offX), MAX_SHADOW_OFFSET_PX)
  ctx.shadowOffsetY = Math.min(Math.max(-MAX_SHADOW_OFFSET_PX, offY), MAX_SHADOW_OFFSET_PX)
}

/** The same resolved face and tracking settings are used for measuring and painting.
 * Resolvers return either a bare family (quoted here) or a full CSS stack such
 * as `"Liter", sans-serif` (used verbatim, never quoted as one family). */
export function fontShorthand(
  style: { italic?: boolean; bold?: boolean; fontSizePt?: number; fontFamily?: string },
  resolveFont: FontResolver,
): string {
  const raw = resolveFont(style.fontFamily ?? 'Calibri')
  const family = /^".*",/.test(raw) ? raw : `"${raw.replace(/["\\]/g, '\\$&')}"`
  return `${style.italic ? 'italic ' : ''}${style.bold ? 'bold ' : ''}${style.fontSizePt ?? 12}pt ${family}`
}
function settings(ctx: CanvasRenderingContext2D, style: PptxTextStyle, resolveFont: FontResolver): boolean {
  ctx.font = fontShorthand(style, resolveFont)
  if ('lang' in ctx) (ctx as unknown as { lang: string }).lang = style.language ?? 'inherit'
  const nativeTracking = 'letterSpacing' in ctx
  if (nativeTracking) ctx.letterSpacing = `${(style.characterSpacingPt ?? 0) * 96 / 72}px`
  return nativeTracking
}
export function createTextBodyMeasurer(ctx: CanvasRenderingContext2D, resolveFont: FontResolver): MeasureText {
    const measureLocal = createLocalTextMeasurer(ctx)
    return (text, style) => {
      const nativeTracking = settings(ctx, style, resolveFont), m = measureLocal(canvasText(text))
      const tracking = nativeTracking || needsContextualShaping(text) ? 0 : (style.characterSpacingPt ?? 0) * 96 / 72
      // Canvas advance includes spacing after the last eligible glyph (WPT
      // 2d.text.drawing.style.letterSpacing.measure). Mirror that in fallback;
      // styled segment boundaries therefore retain the previous run's gap.
      return { width: Math.max(0, m.width + tracking * graphemes(text).filter(g => trackingEligible(g.text)).length),
        ascent: m.actualBoundingBoxAscent, descent: m.actualBoundingBoxDescent,
        normalHeight: (m.fontBoundingBoxAscent ?? (m as TextMetrics & { emHeightAscent?: number }).emHeightAscent ?? m.actualBoundingBoxAscent) + (m.fontBoundingBoxDescent ?? (m as TextMetrics & { emHeightDescent?: number }).emHeightDescent ?? m.actualBoundingBoxDescent) }
    }
}


/**
 * Scratch raster pixel budget for the shaped-run warp painter. A band whose
 * padded raster area would exceed this paints unwarped via the ordinary path
 * (declared policy: bounded scratch memory; policy documentation lives in
 * REPORT.md — no silent unbounded allocation for file-controlled sizes). */
export const MAX_WARP_SURFACE_PIXELS = 4_194_304
/** Hard cap for any single raster dimension. */
export const MAX_WARP_SURFACE_SIDE = 8_192

function affineForTriangle(
  s0: [number, number], s1: [number, number], s2: [number, number],
  d0: [number, number], d1: [number, number], d2: [number, number],
): [number, number, number, number, number, number] | undefined {
  // dst = M·src via Cramer's rule on the 2x3 affine system.
  const [x0, y0] = s0, [x1, y1] = s1, [x2, y2] = s2
  const det = (x1 - x0) * (y2 - y0) - (x2 - x0) * (y1 - y0)
  if (!Number.isFinite(det) || Math.abs(det) < 1e-12) return undefined
  const a = ((d1[0] - d0[0]) * (y2 - y0) - (d2[0] - d0[0]) * (y1 - y0)) / det
  const c = ((d2[0] - d0[0]) * (x1 - x0) - (d1[0] - d0[0]) * (x2 - x0)) / det
  const b = ((d1[1] - d0[1]) * (y2 - y0) - (d2[1] - d0[1]) * (y1 - y0)) / det
  const d = ((d2[1] - d0[1]) * (x1 - x0) - (d1[1] - d0[1]) * (x2 - x0)) / det
  const e = d0[0] - (a * x0 + c * y0)
  const f = d0[1] - (b * x0 + d * y0)
  if (![a, b, c, d, e, f].every(Number.isFinite)) return undefined
  return [a, b, c, d, e, f]
}

/** Clip mesh triangles EXACTLY at shared lattice edges. Adjacent cells share
 * identical edge coordinates, so translucent ink composites once per pixel and
 * authored alpha is preserved (no overlap accumulation). Abutting antialiased
 * edges may leave sub-pixel seams on opaque backgrounds; that trade-off is
 * documented in REPORT.md and measured by the transparent-target alpha test. */
function drawWarpedTriangle(
  ctx: CanvasRenderingContext2D, surface: PaintSurface,
  src: Array<[number, number]>, dst: Array<[number, number]>,
): void {
  // Zero-area destinations (e.g. a collapsed envelope boundary column) cover
  // no pixels: skip before touching canvas state. Clipping an empty path can
  // otherwise empty the clip region for every later triangle on contexts
  // that do not recover it on restore, blanking the whole surface.
  const dx1 = dst[1][0] - dst[0][0], dy1 = dst[1][1] - dst[0][1]
  const dx2 = dst[2][0] - dst[0][0], dy2 = dst[2][1] - dst[0][1]
  if (Math.abs(dx1 * dy2 - dx2 * dy1) <= 1e-9) return
  ctx.save()
  ctx.beginPath()
  dst.forEach(([px, py], index) => {
    if (index === 0) ctx.moveTo(px, py)
    else ctx.lineTo(px, py)
  })
  ctx.closePath()
  ctx.clip()
  const m = affineForTriangle(src[0], src[1], src[2], dst[0], dst[1], dst[2])
  if (m) {
    ctx.transform(m[0], m[1], m[2], m[3], m[4], m[5])
    ctx.drawImage(surface as unknown as CanvasImageSource, 0, 0)
  }
  ctx.restore()
}

interface WarpRunSegmentLike {
  text: string
  x: number
  width: number
  transform?: { a: number; b: number; c: number; d: number; e: number; f: number }
  style: PptxTextStyle
}

function styleFingerprint(style: PptxTextStyle, resolveFont: FontResolver): string {
  return JSON.stringify([
    resolveFont(style.fontFamily ?? 'Calibri'), style.fontSizePt, style.bold, style.italic,
    style.characterSpacingPt, style.language, style.fontFamilyEastAsia, style.fontFamilyComplexScript,
    style.noFill, style.color, style.textFill, style.textOutline, style.textShadow,
  ])
}

interface WarpRunGroup {
  start: number
  end: number
  /** 'transformed' groups are single-member (vertical writing frames); the
   * raster internalizes each member transform inside the body-frame surface. */
  frame: 'normal' | 'transformed'
  style: PptxTextStyle
}

/** consecutive visible segments sharing style + writing frame */
function buildWarpRunGroups(segments: WarpRunSegmentLike[], resolveFont: FontResolver): WarpRunGroup[] {
  const groups: WarpRunGroup[] = []
  let start = -1
  let key = ''
  let frame: 'normal' | 'transformed' = 'normal'
  let frameKey = ''
  for (let i = 0; i <= segments.length; i++) {
    const segment = segments[i]
    const isPaintable = !!segment && segment.text.length > 0 && !/^[\n\t]$/.test(segment.text)
    const nextKey = isPaintable ? styleFingerprint(segment.style, resolveFont) : ''
    const nextFrame = segment?.transform ? 'transformed' : 'normal'
    const nextFrameKey = segment?.transform ? JSON.stringify(segment.transform) : ''
    if (start >= 0 && (!isPaintable || nextKey !== key || nextFrame !== frame || nextFrameKey !== frameKey)) {
      groups.push({ start, end: i - 1, frame, style: segments[start].style })
      start = -1
    }
    if (isPaintable && start < 0) { start = i; key = nextKey; frame = nextFrame; frameKey = nextFrameKey }
  }
  return groups
}

/** Local writing-frame mapping for a transformed (e.g. vertical) segment.
 *
 * The member's (t.e, t.f) is a PHYSICAL body translation, not local-body
 * coordinates. The warp runs in the local body frame reached through the
 * whole physical↔local body affine: for a clockwise quarter turn (vert,
 * b=+1) the member origin is (t.f, box.w−t.e); for counter-clockwise
 * (vert270, b=−1) it is (box.h−t.f, t.e). Dimensions swap for quarter turns.
 * Every mapped local point composes back through that same whole-body affine
 * exactly once, plus the caller origin — whole-body positions are never
 * rotated a second time about the glyph origin. Pure-translation members
 * (upright glyphs in a vertical column) keep their offset in the coinciding
 * frames. This mirrors the independent local-frame reference (local warp,
 * then the ordinary outer quarter-turn). */
function localWarpFrame(
  mapping: WarpMapping,
  box: { x: number; y: number; w: number; h: number },
  t: { a: number; b: number; c: number; d: number; e: number; f: number },
): { local: WarpMapping; localW: number; localH: number; toBody: (lx: number, ly: number) => [number, number]; origin: { x: number; y: number } } | undefined {
  const det = t.a * t.d - t.b * t.c
  if (!Number.isFinite(det) || det === 0) return undefined
  const quarter = Math.abs(t.b) > Math.abs(t.a)
  const localW = quarter ? box.h : box.w
  const localH = quarter ? box.w : box.h
  if (!(localW > 0 && localH > 0)) return undefined
  let local: WarpMapping
  try {
    local = buildWarpMapping({ preset: mapping.preset, adjustments: { ...mapping.resolution.effective } }, localW, localH)
  } catch {
    return undefined
  }
  if (!quarter) {
    return {
      local, localW, localH,
      origin: { x: t.e, y: t.f },
      toBody: (lx, ly) => [box.x + lx, box.y + ly],
    }
  }
  const clockwise = t.b > 0
  return {
    local, localW, localH,
    origin: clockwise ? { x: t.f, y: box.w - t.e } : { x: box.h - t.f, y: t.e },
    toBody: clockwise
      ? (lx, ly) => [box.x + (box.w - ly), box.y + lx]
      : (lx, ly) => [box.x + ly, box.y + (box.h - lx)],
  }
}

/** Dense samples along a mapped band edge for non-convex curved coverage. */
const WARP_POLYGON_SAMPLES = 12

/** Segment visual geometry for curved hit/selection support: same logical
 * record, optional mapped polygon + per-cluster mapped centers (in paintTextBody
 * box coordinates; search projects with the record transform). The polygon
 * samples the mapped band edges densely so curved (non-convex) ink is covered;
 * transformed (vertical) segments warp in their local writing frame. */
/** Readonly tuples: cells/tris are built once and only read thereafter, matching
 * the readonly visual contracts in text-recording/search and the reviewed
 * cell-order regression, which passes recorded cells straight in. */
export type WarpVisualCell = readonly [topLeft: readonly [number, number], topRight: readonly [number, number], bottomRight: readonly [number, number], bottomLeft: readonly [number, number]]
/** Mapped mesh triangles covering the band. Envelope bands can twist (upper and
 * lower boundary x-progression may run opposite), so even single quads can be
 * bowties — triangles are always convex and their union covers folded ink. */
export type WarpVisualTri = readonly [readonly [number, number], readonly [number, number], readonly [number, number]]

function segmentWarpGeometry(
  mapping: WarpMapping,
  box: { x: number; y: number; w: number; h: number },
  segment: WarpRunSegmentLike,
  baselineY: number,
  measure: MeasureText,
): { polygon: Array<[number, number]>; clusters: Array<[number, number]>; cells: WarpVisualCell[]; tris: WarpVisualTri[]; bands: Array<[number, number, number, number]> } {
  const metrics = measure(canvasText(segment.text), segment.style)
  const fontSize = segment.style.fontSizePt ?? 12
  const asc = Math.max(0, metrics.ascent ?? 0) > 0 ? metrics.ascent as number : fontSize * 0.8
  const desc = Math.max(0, metrics.descent ?? 0) > 0 ? metrics.descent as number : fontSize * 0.2
  // Pixel-center coverage: rasterization and resampling spread opaque ink
  // beyond the analytic ascent/descent, and magnifying mappings scale that
  // edge outward. Growing the band here in SOURCE space (transversely, before
  // mapping) scales correctly through any mapping and preserves exact
  // parameter correspondence of the two edges (e.g. Slant x-correspondence);
  // no post-mapping perturbation is applied. Cells, triangles, the polygon
  // and the highlight bands all describe this same covered region (within the
  // 2px tolerance the ordinary strict-hit path already applies); the logical
  // source and the cluster centers are untouched.
  const ascBand = asc + 2, descBand = desc + 2
  // Map a body-frame point through the body mapping.
  const bodyPoint = (cx: number, cy: number): [number, number] => {
    const mapped = mapping.map((cx - box.x) / box.w, (cy - box.y) / box.h)
    return [box.x + mapped.x, box.y + mapped.y]
  }
  // Shared parameter grid for a top/bottom edge pair, adaptively subdivided
  // where either mapped edge's chord nips the true curve. Containment uses
  // chords between consecutive grid points, so this directly bounds the
  // triangle-union-vs-ink error; both edges share the grid so cells keep
  // pairing corresponding samples.
  const refineGrid = (
    topInput: (t: number) => [number, number],
    bottomInput: (t: number) => [number, number],
    topOutput: (mx: number, my: number) => [number, number],
    bottomOutput: (mx: number, my: number) => [number, number],
  ): number[] => {
    let grid: number[] = []
    for (let i = 0; i <= WARP_POLYGON_SAMPLES; i++) grid.push(i / WARP_POLYGON_SAMPLES)
    const TOL = 0.2, MAX_INTERVALS = 256
    for (let pass = 0; pass < 8 && grid.length - 1 < MAX_INTERVALS; pass++) {
      const next: number[] = [grid[0]]
      let split = false
      for (let k = 0; k < grid.length - 1; k++) {
        const a = grid[k], b = grid[k + 1], m = (a + b) / 2
        let dev = 0
        for (const [pin, pout] of [[topInput, topOutput], [bottomInput, bottomOutput]] as const) {
          const pa = pout(...pin(a)), pb = pout(...pin(b)), pm = pout(...pin(m))
          dev = Math.max(dev, Math.hypot(pm[0] - (pa[0] + pb[0]) / 2, pm[1] - (pa[1] + pb[1]) / 2))
        }
        if (dev > TOL && b - a > 1e-9) { next.push(m); split = true }
        next.push(b)
      }
      grid = next
      if (!split) break
    }
    return grid
  }
  const mapEdgeOnGrid = (
    grid: number[],
    mapInput: (t: number) => [number, number],
    mapOutput: (mx: number, my: number) => [number, number],
  ): Array<[number, number]> => grid.map(t => mapOutput(...mapInput(t)))
  // Per-cluster mapped band bounds from the edge arrays: interpolate the
  // edges at the cluster's own source-interval endpoints and bound those with
  // the interior grid samples. Same order as the cluster centers; describes
  // the cluster's actual mapped ink for highlights.
  const bandsForIntervals = (
    intervals: Array<[number, number]>,
    grid: number[],
    topPts: Array<[number, number]>,
    botPts: Array<[number, number]>,
  ): Array<[number, number, number, number]> => {
    const at = (pts: Array<[number, number]>, t: number): [number, number] => {
      if (t <= grid[0]) return pts[0]
      for (let i = 0; i < grid.length - 1; i++) {
        if (t <= grid[i + 1]) {
          const f = (t - grid[i]) / Math.max(1e-12, grid[i + 1] - grid[i])
          return [pts[i][0] + (pts[i + 1][0] - pts[i][0]) * f, pts[i][1] + (pts[i + 1][1] - pts[i][1]) * f]
        }
      }
      return pts[pts.length - 1]
    }
    return intervals.map(([t0, t1]) => {
      const pts: Array<[number, number]> = [at(topPts, t0), at(botPts, t0), at(topPts, t1), at(botPts, t1)]
      for (let i = 0; i < grid.length; i++) if (grid[i] > t0 && grid[i] < t1) { pts.push(topPts[i], botPts[i]) }
      const xs = pts.map(p => p[0]), ys = pts.map(p => p[1])
      const x = Math.min(...xs), y = Math.min(...ys)
      return [x, y, Math.max(...xs) - x, Math.max(...ys) - y]
    })
  }
  // Grapheme source intervals as fractions of the segment advance, aligned
  // with the cluster-center loop of each branch below.
  const clusterIntervals = (): Array<[number, number]> => {
    const out: Array<[number, number]> = []
    const total = segment.width > 0 ? segment.width : 1
    let prefix = ''
    for (const g of graphemes(segment.text)) {
      const through = prefix + g.text
      const throughWidth = measure(canvasText(through), segment.style).width
      const glyphWidth = measure(canvasText(g.text), segment.style).width
      out.push([Math.max(0, (throughWidth - glyphWidth) / total), Math.min(1, throughWidth / total)])
      prefix = through
    }
    return out
  }
  let polygon: Array<[number, number]>
  let clusterCenters: Array<[number, number]>
  let clusterBands: Array<[number, number, number, number]>
  // Small mesh triangles between consecutive edge samples. Global polygon
  // winding can degrade on folded/twisted bands, so hit testing uses this
  // triangle union (always convex) instead of the single polygon. Quads are
  // retained for compatibility; triangles are authoritative for containment.
  const cells: WarpVisualCell[] = []
  const tris: WarpVisualTri[] = []
  const emitCells = (top: Array<[number, number]>, bottom: Array<[number, number]>): void => {
    for (let i = 0; i < top.length - 1 && i < bottom.length - 1; i++) {
      cells.push([top[i], top[i + 1], bottom[i + 1], bottom[i]])
      tris.push([top[i], top[i + 1], bottom[i + 1]], [top[i], bottom[i + 1], bottom[i]])
    }
  }
  if (!segment.transform) {
    const sx = box.x + segment.x
    const topIn = (t: number): [number, number] => [sx + segment.width * t, baselineY - ascBand]
    const bottomIn = (t: number): [number, number] => [sx + segment.width * t, baselineY + descBand]
    const grid = refineGrid(topIn, bottomIn, bodyPoint, bodyPoint)
    const top = mapEdgeOnGrid(grid, topIn, bodyPoint)
    const bottom = mapEdgeOnGrid(grid, bottomIn, bodyPoint)
    // Non-mutating copy: emitCells below pairs top[i] with the corresponding
    // bottom[i]; Array.reverse() in place would cross the band into chords.
    polygon = [...top, ...[...bottom].reverse()]
    emitCells(top, bottom)
    clusterCenters = []
    clusterBands = bandsForIntervals(clusterIntervals(), grid, top, bottom)
    {
      const gm = graphemes(segment.text)
      let prefix = ''
      for (const g of gm) {
        const through = prefix + g.text
        const throughWidth = measure(canvasText(through), segment.style).width
        const glyphWidth = measure(canvasText(g.text), segment.style).width
        const centerLocal = throughWidth - glyphWidth / 2
        clusterCenters.push(bodyPoint(box.x + segment.x + centerLocal, baselineY))
        prefix = through
      }
    }
  } else {
    const frame = localWarpFrame(mapping, box, segment.transform)
    if (!frame || frame.local.identity) {
      // Identity-local (or unresolvable) warp: ordinary unwarped band geometry
      // so hits behave like the ordinary painter.
      const t = segment.transform
      const apply = (px: number, py: number): [number, number] =>
        [box.x + t.e + t.a * px + t.c * py, box.y + t.f + t.b * px + t.d * py]
      const topInId = (tt: number): [number, number] => [segment.width * tt, -ascBand]
      const bottomInId = (tt: number): [number, number] => [segment.width * tt, descBand]
      const outId = (px: number, py: number): [number, number] => apply(px, py)
      const gridId = refineGrid(topInId, bottomInId, outId, outId)
      const top = mapEdgeOnGrid(gridId, topInId, outId)
      const bottom = mapEdgeOnGrid(gridId, bottomInId, outId)
      // Non-mutating copy (see above): cells keep original corresponding order.
      polygon = [...top, ...[...bottom].reverse()]
      emitCells(top, bottom)
      clusterCenters = []
      clusterBands = bandsForIntervals(clusterIntervals(), gridId, top, bottom)
      const gm = graphemes(segment.text)
      let prefix = ''
      for (const g of gm) {
        const through = prefix + g.text
        const throughWidth = measure(canvasText(through), segment.style).width
        const glyphWidth = measure(canvasText(g.text), segment.style).width
        clusterCenters.push(apply(throughWidth - glyphWidth / 2, 0))
        prefix = through
      }
      return { polygon, clusters: clusterCenters, cells, tris, bands: clusterBands }
    }
    const { local, localW, localH, origin, toBody } = frame
    const localPoint = (lx: number, ly: number): [number, number] => {
      const mapped = local.map(lx / localW, ly / localH)
      return toBody(mapped.x, mapped.y)
    }
    const topIn = (t: number): [number, number] => [origin.x + segment.width * t, origin.y - ascBand]
    const bottomIn = (t: number): [number, number] => [origin.x + segment.width * t, origin.y + descBand]
    const grid = refineGrid(topIn, bottomIn, localPoint, localPoint)
    const top = mapEdgeOnGrid(grid, topIn, localPoint)
    const bottom = mapEdgeOnGrid(grid, bottomIn, localPoint)
    // Non-mutating copy (see above): cells keep original corresponding order.
    polygon = [...top, ...[...bottom].reverse()]
    emitCells(top, bottom)
    clusterCenters = []
    clusterBands = bandsForIntervals(clusterIntervals(), grid, top, bottom)
    {
      const gm = graphemes(segment.text)
      let prefix = ''
      for (const g of gm) {
        const through = prefix + g.text
        const throughWidth = measure(canvasText(through), segment.style).width
        const glyphWidth = measure(canvasText(g.text), segment.style).width
        clusterCenters.push(localPoint(origin.x + throughWidth - glyphWidth / 2, origin.y))
        prefix = through
      }
    }
  }
  return { polygon, clusters: clusterCenters, cells, tris, bands: clusterBands }
}

interface WarpedRunGroupInput {
  mapping: WarpMapping
  /** target context (surface creation + final composite) */
  ctx: CanvasRenderingContext2D
  box: { x: number; y: number; w: number; h: number }
  group: WarpRunGroup
  segments: WarpRunSegmentLike[]
  lineTop: number
  lineBottom: number
  lineBaseline: number
  resolveFont: FontResolver
  patternCache: Map<string, CanvasPattern | null>
  measure: MeasureText
  /** Called per member when the scratch budget is exceeded — paints the
   * unwarped ordinary path (bounded-memory policy, no silent budget). */
  paintFallback: (segment: WarpRunSegmentLike) => void
  /** Called once when this group takes the diagnosed fallback path (budget or
   * surface failure) so paint and record sides share one honest diagnostic. */
  diagnoseFallback?: () => void
  /** Called once when scratch-surface allocation itself fails (host/resource),
   * so the diagnostic names the allocation failure instead of the budget. */
  diagnoseAllocFallback?: () => void
}

/** Pure scratch-budget check shared by the paint and record paths: a group
 * whose integer-aligned scratch window would exceed the bounded surface
 * policy must render AND record through the ordinary fallback (never warped
 * paint with unwarped records, or vice versa). Takes the exact integer window
 * dimensions from the shared plan. */
export function warpedGroupExceedsBudget(rasterW: number, rasterH: number): boolean {
  return rasterW > MAX_WARP_SURFACE_SIDE || rasterH > MAX_WARP_SURFACE_SIDE || rasterW * rasterH > MAX_WARP_SURFACE_PIXELS
}

/** Shared warp-rendering plan: paint and record paths resolve the SAME mode
 * (warped raster vs ordinary fallback) from pure layout/budget math, so a
 * budget-exceeded run never paints unwarped while recording warped geometry
 * (or vice versa). Live surface-allocation failure stays paint-time-only and
 * is reported through `diagnoseFallback`. */
interface WarpedGroupPlan {
  members: WarpRunSegmentLike[]
  style: PptxTextStyle
  asc: number
  desc: number
  frameMapping: WarpMapping
  frameW: number
  frameH: number
  toCaller: (fx: number, fy: number) => [number, number]
  frameOrigin: { x: number; y: number }
  useLocalFrame: boolean
  frameBaseline: number
  x0: number
  y0: number
  x1: number
  y1: number
  pad: number
  /** Integer-aligned source window origin (floor of the padded band). The
   * raster is painted at the inverse integer translation so ordinary glyph
   * pixel phase is preserved; rasterW/H are the exact ceil/floor extents. */
  ix0: number
  iy0: number
  rasterW: number
  rasterH: number
  /** True → render AND record through the ordinary path (no visual payload). */
  fallback: boolean
  /** True → the fallback is a diagnosed resource/policy outcome. */
  diagnose: boolean
}

function planWarpedGroup(
  mapping: WarpMapping,
  box: { x: number; y: number; w: number; h: number },
  group: WarpRunGroup,
  segments: WarpRunSegmentLike[],
  lineBaseline: number,
  measure: MeasureText,
): WarpedGroupPlan | undefined {
  const members = segments.slice(group.start, group.end + 1)
  const runText = members.map(m => canvasText(m.text)).join('')
  if (!runText.length) return undefined
  const style = group.style
  const outline = style.textOutline
  const shadow = style.textShadow
  const metrics = measure(runText, style)
  const fontSize = style.fontSizePt ?? 12
  const asc = Math.max(0, metrics.ascent ?? 0) > 0 ? metrics.ascent as number : fontSize * 0.8
  const desc = Math.max(0, metrics.descent ?? 0) > 0 ? metrics.descent as number : fontSize * 0.2
  const firstTransform = members.find(m => m.transform)?.transform
  const useLocalFrame = group.frame === 'transformed' && !!firstTransform
  let frameMapping: WarpMapping = mapping
  let frameW = box.w, frameH = box.h
  let toCaller: (fx: number, fy: number) => [number, number] = (fx, fy) => [box.x + fx, box.y + fy]
  let frameOrigin = { x: 0, y: 0 }
  if (useLocalFrame) {
    const frame = localWarpFrame(mapping, box, firstTransform!)
    if (!frame || frame.local.identity) return {
      members, style, asc, desc, frameMapping: mapping, frameW: box.w, frameH: box.h,
      toCaller: (fx, fy) => [box.x + fx, box.y + fy] as [number, number],
      frameOrigin, useLocalFrame: false, frameBaseline: lineBaseline - box.y,
      x0: 0, y0: 0, x1: 0, y1: 0, pad: 0, ix0: 0, iy0: 0, rasterW: 0, rasterH: 0,
      fallback: true, diagnose: false,
    }
    frameMapping = frame.local
    frameW = frame.localW; frameH = frame.localH
    frameOrigin = frame.origin
    toCaller = (fx, fy) => frame.toBody(fx, fy)
  }
  let x0 = Infinity, x1 = -Infinity, y0 = Infinity, y1 = -Infinity
  for (const member of members) {
    const metrics2 = measure(canvasText(member.text), member.style)
    const a2 = Math.max(0, metrics2.ascent ?? 0) > 0 ? metrics2.ascent as number : fontSize * 0.8
    const d2 = Math.max(0, metrics2.descent ?? 0) > 0 ? metrics2.descent as number : fontSize * 0.2
    if (!useLocalFrame && !member.transform) {
      x0 = Math.min(x0, member.x); x1 = Math.max(x1, member.x + member.width)
      const top = lineBaseline - box.y - a2, bottom = lineBaseline - box.y + d2
      y0 = Math.min(y0, top); y1 = Math.max(y1, bottom)
    } else if (!useLocalFrame) {
      const t = member.transform!
      for (const [px, py] of [[0, -a2], [member.width, -a2], [member.width, d2], [0, d2]] as Array<[number, number]>) {
        const cx = t.e + t.a * px + t.c * py
        const cy = t.f + t.b * px + t.d * py
        x0 = Math.min(x0, cx); x1 = Math.max(x1, cx); y0 = Math.min(y0, cy); y1 = Math.max(y1, cy)
      }
    } else {
      x0 = Math.min(x0, frameOrigin.x); x1 = Math.max(x1, frameOrigin.x + member.width)
      y0 = Math.min(y0, frameOrigin.y - a2); y1 = Math.max(y1, frameOrigin.y + d2)
    }
  }
  const frameBaseline = useLocalFrame ? frameOrigin.y : lineBaseline - box.y
  if (![x0, x1, y0, y1].every(Number.isFinite) || x1 <= x0 || y1 <= y0) {
    return {
      members, style, asc, desc, frameMapping, frameW, frameH, toCaller, frameOrigin,
      useLocalFrame, frameBaseline, x0: 0, y0: 0, x1: 0, y1: 0, pad: 0, ix0: 0, iy0: 0, rasterW: 0, rasterH: 0,
      fallback: true, diagnose: false,
    }
  }
  if (![x0, x1, y0, y1].every(Number.isFinite) || x1 <= x0 || y1 <= y0) {
    return {
      members, style, asc, desc, frameMapping, frameW, frameH, toCaller, frameOrigin,
      useLocalFrame, frameBaseline, x0: 0, y0: 0, x1: 0, y1: 0, pad: 0, ix0: 0, iy0: 0, rasterW: 0, rasterH: 0,
      fallback: true, diagnose: false,
    }
  }
  const pad = Math.min(64, 4 + (outline ? outline.widthPx / 2 : 0) + (shadow ? Math.max(Math.abs(shadow.offsetX ?? 0), Math.abs(shadow.offsetY ?? 0)) : 0))
  // Integer-aligned source window: floor left/top, ceil right/bottom. Glyphs
  // rasterize at the inverse INTEGER translation, preserving ordinary pixel
  // phase before the exact mapping. This same window origin/span drives
  // allocation, budget, lattice region and affine correspondence below.
  const ix0 = Math.floor(x0 - pad), iy0 = Math.floor(y0 - pad)
  const rasterW = Math.ceil(x1 + pad) - ix0, rasterH = Math.ceil(y1 + pad) - iy0
  if (warpedGroupExceedsBudget(rasterW, rasterH)) {
    return {
      members, style, asc, desc, frameMapping, frameW, frameH, toCaller, frameOrigin,
      useLocalFrame, frameBaseline, x0, y0, x1, y1, pad, ix0, iy0, rasterW, rasterH,
      fallback: true, diagnose: true,
    }
  }
  return {
    members, style, asc, desc, frameMapping, frameW, frameH, toCaller, frameOrigin,
    useLocalFrame, frameBaseline, x0, y0, x1, y1, pad, ix0, iy0, rasterW, rasterH,
    fallback: false, diagnose: false,
  }
}

/** Combined source-raster→caller affine for the whole-raster fastpath.
 *
 * Source-raster point (sx, sy) IS the frame point (bandX+sx, bandY+sy) (the
 * raster is painted at unit scale from the integer window origin), and its
 * caller destination is
 * toCaller(frameMapping.map(...)). When that combined map is affine — true
 * for affine envelopes such as SlantUp/Down, including under the quarter-turn
 * frame composition — the returned canvas 2x3 draws the complete shaped raster
 * once, exactly. Genuinely curved meshes fail the dense-grid affinity check
 * (corners+midpoint alone admit S-shaped envelopes) and stay on the mesh path. */
function combinedAffineTransform(
  frameMapping: WarpMapping,
  frameW: number, frameH: number,
  toCaller: (fx: number, fy: number) => [number, number],
  bandX: number, bandY: number, physW: number, physH: number,
): [number, number, number, number, number, number] | undefined {
  if (!(physW > 0 && physH > 0 && frameW > 0 && frameH > 0)) return undefined
  const dest = (sx: number, sy: number): [number, number] => {
    const mapped = frameMapping.map((bandX + sx) / frameW, (bandY + sy) / frameH)
    return toCaller(mapped.x, mapped.y)
  }
  const p00 = dest(0, 0), p10 = dest(physW, 0), p01 = dest(0, physH)
  if (![p00, p10, p01].every(p => p.every(Number.isFinite))) return undefined
  const m: [number, number, number, number, number, number] = [(p10[0] - p00[0]) / physW, (p10[1] - p00[1]) / physW, (p01[0] - p00[0]) / physH, (p01[1] - p00[1]) / physH, p00[0], p00[1]]
  // Affinity check on the COMBINED source→caller map over a dense grid:
  // corners+midpoint alone admit S-shaped envelopes that match an affine at
  // those five points yet bulge pixels interiorly, so every grid point must
  // agree. Only truly affine envelopes (e.g. SlantUp/Down) take the fastpath;
  // genuinely curved meshes stay on the mesh path with exact per-cell paint.
  const tol = 1e-6 * Math.max(1, physW, physH)
  const N = 8
  for (let j = 0; j <= N; j++) {
    for (let i = 0; i <= N; i++) {
      if ((i === 0 || i === N) && (j === 0 || j === N)) continue
      const sx = (physW * i) / N, sy = (physH * j) / N
      const p = dest(sx, sy)
      if (!p.every(Number.isFinite)) return undefined
      if (Math.hypot(m[0] * sx + m[2] * sy + m[4] - p[0], m[1] * sx + m[3] * sy + m[5] - p[1]) > tol) return undefined
    }
  }
  return m
}

/** Single-drawImage composite of the whole shaped source under an exact affine.
 * Affine maps send corners to corners, so the destination bounds come from the
 * four source corners. Shadowed runs compose into one intermediate surface
 * first so the shadow stays coherent under the deformed ink. */
function paintAffineSource(
  ctx: CanvasRenderingContext2D, surface: PaintSurface,
  m: [number, number, number, number, number, number],
  physW: number, physH: number,
  style: PptxTextStyle | undefined,
): void {
  const at = (sx: number, sy: number): [number, number] => [m[0] * sx + m[2] * sy + m[4], m[1] * sx + m[3] * sy + m[5]]
  if (!style?.textShadow) {
    ctx.save()
    try {
      ctx.transform(m[0], m[1], m[2], m[3], m[4], m[5])
      ctx.drawImage(surface as unknown as CanvasImageSource, 0, 0)
    } finally { ctx.restore() }
    return
  }
  const corners = [at(0, 0), at(physW, 0), at(0, physH), at(physW, physH)]
  const minX = Math.min(...corners.map(p => p[0])), minY = Math.min(...corners.map(p => p[1]))
  const maxX = Math.max(...corners.map(p => p[0])), maxY = Math.max(...corners.map(p => p[1]))
  const warpW = Math.ceil(maxX - minX), warpH = Math.ceil(maxY - minY)
  // A composed-surface failure here must not change the render mode: direct
  // warped ink still agrees with the warped records (unlike a main-surface
  // failure, which falls back to ordinary on both sides). The 2D context is
  // acquired inside the same guard — a null or throwing context continues on
  // the direct path exactly like a missing surface.
  let composed: PaintSurface | undefined
  let composedCtx: CanvasRenderingContext2D | undefined
  try {
    composed = warpW > 0 && warpH > 0 && warpW <= MAX_WARP_SURFACE_SIDE && warpH <= MAX_WARP_SURFACE_SIDE && warpW * warpH <= MAX_WARP_SURFACE_PIXELS ? scratchSurface(ctx, warpW, warpH) ?? undefined : undefined
    composedCtx = composed?.getContext('2d') as CanvasRenderingContext2D | undefined ?? undefined
  } catch {
    composed = undefined
    composedCtx = undefined
  }
  if (composed && composedCtx) {
    composedCtx.setTransform(1, 0, 0, 1, -minX, -minY)
    composedCtx.transform(m[0], m[1], m[2], m[3], m[4], m[5])
    composedCtx.drawImage(surface as unknown as CanvasImageSource, 0, 0)
    applyTextShadow(ctx, style)
    ctx.drawImage(composed as unknown as CanvasImageSource, minX, minY)
    applyTextShadow(ctx, {})
    return
  }
  applyTextShadow(ctx, style)
  try {
    ctx.save()
    try {
      ctx.transform(m[0], m[1], m[2], m[3], m[4], m[5])
      ctx.drawImage(surface as unknown as CanvasImageSource, 0, 0)
    } finally { ctx.restore() }
  } finally { applyTextShadow(ctx, {}) }
}

/** Paint one compatible run as a shaped source raster, deformed through the
 * warp mapping as a coherent surface (interior envelope deformation — not a
 * per-glyph approximation). Gradient/pattern brush is resolved once per run in
 * the run's own source coordinate box, so warped color tracks the mapped
 * source position continuously. Shadow composites as ONE final drawImage so
 * the shadow stays coherent under the deformed ink. */
function paintWarpedRunGroup(input: WarpedRunGroupInput): void {
  const { mapping, ctx, box, group, segments, resolveFont, patternCache, measure } = input
  const plan = planWarpedGroup(mapping, box, group, segments, input.lineBaseline, measure)
  if (!plan) return
  if (plan.fallback) {
    if (plan.diagnose) input.diagnoseFallback?.()
    for (const m of plan.members) input.paintFallback(m)
    return
  }
  const { members, style, asc, desc, frameMapping, frameW, frameH, toCaller, frameOrigin, frameBaseline } = plan
  const { x0, x1, ix0, iy0, rasterW, rasterH } = plan
  const runText = members.map(m => canvasText(m.text)).join('')
  const outline = style.textOutline
  const shadow = style.textShadow

  // Source window origin/span: the INTEGER-aligned window from the shared
  // plan. Raster pixel (i,j) IS frame point (ix0+i, iy0+j); destinations span
  // that same window, so the source is never compressed/remapped.
  const physW = rasterW, physH = rasterH
  // Whole-raster exact-affine fastpath (see combinedAffineTransform): affine
  // envelopes draw once, with no mesh seams at all.
  const affine = combinedAffineTransform(frameMapping, frameW, frameH, toCaller, ix0, iy0, physW, physH)
  // Live host allocation failure takes the SAME ordinary fallback as the
  // budget path (frame-aware member paint, no visual payload on the record
  // side via the record-flow probe below) but reports its own truthful
  // diagnostic — never the budget message.
  let surface: PaintSurface | undefined
  let raster: CanvasRenderingContext2D | undefined
  try {
    surface = scratchSurface(ctx, rasterW, rasterH) ?? undefined
    raster = surface?.getContext('2d') as CanvasRenderingContext2D | undefined
  } catch {
    surface = undefined
    raster = undefined
  }
  if (!surface || !raster) {
    if (input.diagnoseAllocFallback) input.diagnoseAllocFallback()
    else input.diagnoseFallback?.()
    for (const member of members) input.paintFallback(member)
    return
  }
  raster.setTransform(1, 0, 0, 1, -ix0, -iy0)
  raster.textAlign = 'left'; raster.textBaseline = 'alphabetic'
  applyTextShadow(raster, {})  // ink only; shadow composites at the final drawImage
  const paintFill = !style.noFill
  const paintOutline = !!outline
  // All raster paint uses FRAME coordinates (the raster transform above maps
  // frame space into raster pixels); frameBaseline comes from the shared plan.
  // Brush resolved ONCE over the run's own box (continuous per source segment).
  let brush: string | CanvasGradient | CanvasPattern | undefined
  try { brush = resolveTextFill(raster, style, { x: x0, y: frameBaseline - asc, width: x1 - x0, height: asc + desc }, patternCache) } catch { brush = undefined }
  const paintOne = (text: string, px: number, py: number, wdt: number, ascent: number, descent: number): void => {
    void wdt
    if (paintFill) raster.fillText(text, px, py)
    if (paintOutline && outline && outline.widthPx > 0) {
      raster.strokeStyle = outline.color
      raster.lineWidth = Math.min(outline.widthPx, MAX_OUTLINE_WIDTH_PX)
      raster.lineJoin = 'round'
      raster.strokeText(text, px, py)
    }
    void ascent; void descent
  }
  const tracking = (style.characterSpacingPt ?? 0) * 96 / 72
  const nativeTracking = settings(raster, style, resolveFont)
  const needsShaping = !nativeTracking && needsContextualShaping(runText)
  const runTextWidth = members.reduce((sum, m) => sum + m.width, 0)
  const isContiguous = members.length === 1 || members.every((m, idx) =>
    idx === 0 || Math.abs(members[idx - 1].x + members[idx - 1].width - m.x) < 0.05
  )
  if (group.frame === 'normal' && isContiguous && (tracking === 0 || needsShaping || nativeTracking)) {
    // A compatible 'normal'-frame run is painted AS ONE SOURCE CALL: the whole
    // run text reaches the raster through a single fillText so contextual
    // joins and cross-run ligatures survive the warp (P2 evidence). Group
    // membership already guarantees one style fingerprint, and with zero
    // tracking the members occupy exactly the joined advance chain, so one
    // call at the first member's position reproduces the same ink.
    if (paintFill && brush !== undefined) raster.fillStyle = brush
    else if (paintFill) raster.fillStyle = style.color ?? '#000000'
    {
      const gm = raster.measureText(runText)
      paintOne(runText, members[0].x, frameBaseline, runTextWidth, gm.actualBoundingBoxAscent, gm.actualBoundingBoxDescent)
    }
  } else if (group.frame === 'normal' && (tracking === 0 || needsShaping || nativeTracking)) {
    if (paintFill && brush !== undefined) raster.fillStyle = brush
    else if (paintFill) raster.fillStyle = style.color ?? '#000000'
    for (const member of members) {
      const mText = canvasText(member.text)
      const gm = raster.measureText(mText)
      paintOne(mText, member.x, frameBaseline, member.width, gm.actualBoundingBoxAscent, gm.actualBoundingBoxDescent)
    }
  } else if (group.frame === 'normal') {
    let prefix = '', index = 0
    if (paintFill && brush !== undefined) raster.fillStyle = brush
    else if (paintFill) raster.fillStyle = style.color ?? '#000000'
    for (const g of graphemes(runText)) {
      const glyph = g.text, through = prefix + glyph
      const gm = raster.measureText(through), gw = raster.measureText(glyph)
      const advance = gm.width - gw.width
      if (trackingEligible(g.text)) {
        paintOne(glyph, members[0].x + advance + index * tracking, frameBaseline, gw.width, 0, 0)
        index++
      }
      prefix = through
    }
  } else {
    // Local-frame groups (e.g. vertical writing frames): members paint WITHOUT
    // their body transform at local-body coordinates — the warp mapping itself
    // is local, and destinations map back through the member affine.
    for (const member of members) {
      const memberTracking = (member.style.characterSpacingPt ?? 0) * 96 / 72
      let memberBrush: string | CanvasGradient | CanvasPattern | undefined = brush
      try {
        memberBrush = resolveTextFill(raster, member.style, { x: frameOrigin.x, y: frameOrigin.y - asc, width: Math.max(member.width, 0.5), height: asc + desc }, patternCache)
      } catch { memberBrush = undefined }
      if (paintFill && memberBrush !== undefined) raster.fillStyle = memberBrush
      else if (paintFill) raster.fillStyle = member.style.color ?? '#000000'
      const memberNative = settings(raster, member.style, resolveFont)
      const memberText = canvasText(member.text)
      const memberNeedsShaping = !memberNative && needsContextualShaping(memberText)
      if (memberTracking === 0 || memberNative || memberNeedsShaping) {
        const gm = raster.measureText(memberText)
        paintOne(memberText, frameOrigin.x, frameOrigin.y, member.width, gm.actualBoundingBoxAscent, gm.actualBoundingBoxDescent)
      } else {
        let prefix = '', index = 0
        for (const g of graphemes(memberText)) {
          const glyph = g.text, through = prefix + glyph
          const gm = raster.measureText(through), gw = raster.measureText(glyph)
          const advance = gm.width - gw.width
          if (trackingEligible(g.text)) {
            paintOne(glyph, frameOrigin.x + advance + index * memberTracking, frameOrigin.y, gw.width, 0, 0)
            index++
          }
          prefix = through
        }
      }
    }
  }
  if (affine) {
    paintAffineSource(ctx, surface, affine, physW, physH, shadow ? style : undefined)
    return
  }
  // Non-affine mesh path. Lattice over the INTEGER source window;
  // destinations carry the caller origin exactly once via toCaller (frame
  // mappings are origin-free).
  const region = {
    u0: ix0 / frameW, v0: iy0 / frameH,
    u1: (ix0 + rasterW) / frameW, v1: (iy0 + rasterH) / frameH,
  }
  const lattice = buildWarpLattice(frameMapping, region)
  const destOf = (p: { x: number; y: number }): [number, number] => toCaller(p.x, p.y)
  let minX = Infinity, minY = Infinity, maxX = -Infinity, maxY = -Infinity
  for (const row of lattice.points) for (const p of row) {
    const [dx, dy] = destOf(p)
    minX = Math.min(minX, dx); minY = Math.min(minY, dy); maxX = Math.max(maxX, dx); maxY = Math.max(maxY, dy)
  }
  // Mesh source/destination correspondence: the source raster spans the exact
  // INTEGER window [ix0, ix0+rasterW] (raster pixel (i,j) is source point
  // (i,j)), and each lattice cell maps that same window interval — unpadded
  // ink is never stretched into padded destinations.
  const sourcePoint = (row: number, col: number): [number, number] => [
    (physW * col) / lattice.cols,
    (physH * row) / lattice.rows,
  ]
  if (shadow) {
    // Coherent shadow: compose the full warped run into one surface, draw once.
    // A composed-surface failure keeps the warped mode (direct mesh below),
    // which still agrees with the warped records. Context acquisition shares
    // the guard: null/throwing contexts continue direct, never ordinary.
    const warpW = Math.ceil(maxX - minX), warpH = Math.ceil(maxY - minY)
    let composed: PaintSurface | undefined
    let composedCtx: CanvasRenderingContext2D | undefined
    try {
      composed = warpW > 0 && warpH > 0 && warpW <= MAX_WARP_SURFACE_SIDE && warpH <= MAX_WARP_SURFACE_SIDE && warpW * warpH <= MAX_WARP_SURFACE_PIXELS ? scratchSurface(ctx, warpW, warpH) ?? undefined : undefined
      composedCtx = composed?.getContext('2d') as CanvasRenderingContext2D | undefined ?? undefined
    } catch {
      composed = undefined
      composedCtx = undefined
    }
    if (composed && composedCtx) {
      composedCtx.setTransform(1, 0, 0, 1, -minX, -minY)
      for (let j = 0; j < lattice.rows; j++) for (let i = 0; i < lattice.cols; i++) {
        const dst0 = destOf(lattice.points[j][i]), dst1 = destOf(lattice.points[j][i + 1]), dst2 = destOf(lattice.points[j + 1][i + 1]), dst3 = destOf(lattice.points[j + 1][i])
        drawWarpedTriangle(composedCtx, surface, [sourcePoint(j, i), sourcePoint(j, i + 1), sourcePoint(j + 1, i + 1)], [dst0, dst1, dst2])
        drawWarpedTriangle(composedCtx, surface, [sourcePoint(j, i), sourcePoint(j + 1, i + 1), sourcePoint(j + 1, i)], [dst0, dst2, dst3])
      }
      applyTextShadow(ctx, style)
      ctx.drawImage(composed as unknown as CanvasImageSource, minX, minY)
      applyTextShadow(ctx, {})
      return
    }
  }
  for (let j = 0; j < lattice.rows; j++) for (let i = 0; i < lattice.cols; i++) {
    const dst0 = destOf(lattice.points[j][i]), dst1 = destOf(lattice.points[j][i + 1]), dst2 = destOf(lattice.points[j + 1][i + 1]), dst3 = destOf(lattice.points[j + 1][i])
    drawWarpedTriangle(ctx, surface, [sourcePoint(j, i), sourcePoint(j, i + 1), sourcePoint(j + 1, i + 1)], [dst0, dst1, dst2])
    drawWarpedTriangle(ctx, surface, [sourcePoint(j, i), sourcePoint(j + 1, i + 1), sourcePoint(j + 1, i)], [dst0, dst2, dst3])
  }
}

export function paintTextBody(body: PptxTextBody, ctx: CanvasRenderingContext2D, x: number, y: number, w: number, h: number, resolveFont: FontResolver, theme?: ThemeContext,
  options?: { layout?: TextLayout; clip?: LogicalTextRange['clip']; clipToLineBox?: boolean }): void {
  if (w - (body.insetLeftEmu + body.insetRightEmu) / 9525 <= 0 || h - (body.insetTopEmu + body.insetBottomEmu) / 9525 <= 0) return
  ctx.save()
  try {
    const viewport = options?.clip ?? (body.direction && body.direction !== 'horz' ? { x, y, width: w, height: h } : undefined)
    const matrix = ctx.getTransform()
    const clip = viewport ? { ...viewport, transform: viewport.transform ?? { a: matrix.a, b: matrix.b, c: matrix.c, d: matrix.d, e: matrix.e, f: matrix.f } } : undefined
    if (viewport) {
      ctx.beginPath(); ctx.rect(viewport.x, viewport.y, viewport.width, viewport.height); ctx.clip()
    }
    const measureLocal = createLocalTextMeasurer(ctx)
    const measure = createTextBodyMeasurer(ctx, resolveFont)
    const patternCache = new Map<string, CanvasPattern | null>()
    const layout = options?.layout ?? layoutTextBody(body, w, h, measure, theme)
    // A fresh identity for each invocation prevents reused legacy paragraphs or
    // equal-height table cells from merging into one source text scope.
    const scope = {}
    const sources: LogicalTextSource[] = body.paragraphs.map((p, order) => ({ text: p.runs.map(r => r.text).join(''), scope, order }))
    ctx.textAlign = 'left'; ctx.textBaseline = 'alphabetic'
    const record = (ctx as TextRecordingContext)[RECORD_TEXT]
    const warp = body.textWarp && body.textWarp.preset !== 'textNoShape' && body.textWarp.preset !== 'textPlain' ? body.textWarp : undefined

    // Warp: one mapping per body box (resolved once, not per glyph). Identity
    // mappings fall through to the ordinary painter so P1 zero-warp shaping,
    // gradients, and records stay byte-identical; record-replay indexing never
    // paints at all. Recording contexts receive unwarped canonical bands plus
    // optional mapped visual geometry (no duplicate logical records).
    let warpMapping: WarpMapping | undefined
    if (warp) {
      try {
        const mapping = buildWarpMapping(warp, w, h)
        // Identity mappings ride the ordinary painter in paint flows and are
        // recorded with no visual payload; nonzero warps map both the record
        // payloads AND the painted surface through the same mapping.
        if (!mapping.identity) warpMapping = mapping
      } catch (error) {
        // Fail-safe: unresolvable warp geometries keep the previously
        // diagnosed behavior (unwarped text), no throw reaches callers.
        const diagnostics = (body as { diagnostics?: Array<{ kind: string; feature: string; message: string }> }).diagnostics
        if (diagnostics) diagnostics.push({ kind: 'unsupported-text-warp', feature: 'warp-resolve', message: `Text warp could not be resolved: ${error instanceof Error ? error.message : String(error)}` })
      }
    }
    const boxForMapping = { x, y, w, h }
    type LineT = (typeof layout.lines)[number]
    type SegmentT = LineT['segments'][number]

    /** Ordinary per-segment paint — byte-identical to the legacy painter so
     * identity warps and non-warp callers keep exact pixels. */
    function paintSegmentInk(segment: SegmentT, line: LineT, nativeTracking: boolean): void {
      const fillIt = !segment.style.noFill
      const outline = segment.style.textOutline
      if (!fillIt && !outline) return
      if (/^[\n\t]$/.test(segment.text)) return
      applyTextShadow(ctx, segment.style)
      const inFrame = !!segment.transform
      const sx = inFrame ? 0 : x + segment.x
      const sy = inFrame ? 0 : y + line.baseline
      const size = segment.style.fontSizePt ?? 12
  const paintOne = (text: string, px: number, py: number, wdt: number, ascent: number, descent: number): void => {
        const asc = ascent > 0 ? ascent : size * 0.8, desc = descent > 0 ? descent : size * 0.2
        if (fillIt) {
          ctx.fillStyle = resolveTextFill(ctx, segment.style,
            { x: px, y: py - asc, width: Math.max(wdt, 0.5), height: asc + desc }, patternCache)
          ctx.fillText(text, px, py)
        }
        if (outline) {
          ctx.strokeStyle = outline.color
          ctx.lineWidth = Math.min(outline.widthPx, MAX_OUTLINE_WIDTH_PX)
          ctx.lineJoin = 'round'
          ctx.strokeText(text, px, py)
        }
      }
      const tracking = !nativeTracking && needsContextualShaping(segment.text) ? 0 : (segment.style.characterSpacingPt ?? 0) * 96 / 72
      if (nativeTracking || tracking === 0) {
        const gm = measureLocal(canvasText(segment.text))
        paintOne(canvasText(segment.text), sx, sy, segment.width, gm.actualBoundingBoxAscent, gm.actualBoundingBoxDescent)
      } else {
        let prefix = '', index = 0
        for (const g of graphemes(segment.text)) {
          const glyph = canvasText(g.text), through = prefix + glyph
          const gm = measureLocal(through), gw = measureLocal(glyph)
          const advance = gm.width - gw.width
          if (trackingEligible(g.text)) {
            paintOne(glyph, sx + advance + index * tracking, sy, gw.width, gw.actualBoundingBoxAscent, gw.actualBoundingBoxDescent)
            index++
          }
          prefix = through
        }
      }
    }

    for (const line of layout.lines) {
      if (options?.clipToLineBox) {
        const lineWidth = line.segments.reduce((acc, s) => Math.max(acc, s.x + s.width - line.x), 0)
        ctx.save()
        ctx.beginPath()
        ctx.rect(x + line.x, y + line.y, Math.max(lineWidth, 1), Math.max(line.height, 1))
        ctx.clip()
      }
      try {
        const visibleSegment = line.segments.find(segment => !segment.style.noFill && segment.text !== '\n')
        const paragraph = body.paragraphs[line.paragraphIndex]
        const bulletStyle = visibleSegment?.style ?? { ...paragraph.defaultProperties, ...paragraph.endProperties }
        if (line.bullet && !bulletStyle.noFill && (visibleSegment || !line.segments.length)) {
          const style = bulletStyle
          settings(ctx, style, resolveFont); ctx.fillStyle = style.color ?? '#000000'
          // Bullets carry no shadow model: reset state so a shadowed run on an
          // earlier line can never leak into them.
          applyTextShadow(ctx, {})
          // Bullet glyphs are presentation, not part of the source text string.
          if (!record) ctx.fillText(line.bullet, x + line.x - measureLocal(line.bullet).width - 4, y + line.baseline)
        }
        if (!line.segments.length && record) {
          const style = { ...body.paragraphs[line.paragraphIndex].defaultProperties, ...body.paragraphs[line.paragraphIndex].endProperties }
          settings(ctx, { ...style, fontFamily: resolveTextFamily({ ...style, text: '' }, theme) }, resolveFont)
          const sourceOffset = line.inlineSlots?.[0]?.sourceOffset ?? sources[line.paragraphIndex].text.length
          record('', x + line.x, y + line.baseline, 0, { source: sources[line.paragraphIndex], start: sourceOffset,
            end: sourceOffset, clip, ...(line.logicalLineIndex === undefined ? {} : { line: line.logicalLineIndex, flow: 'vertical' as const }) })
        }
        // Warp run groups: consecutive compatible segments share one shaped
        // raster (join-preserving); the record loop below still visits every
        // member exactly once with its canonical UTF-16 identity.
        const segments = line.segments
        const warpGroups = warpMapping ? buildWarpRunGroups(segments as unknown as WarpRunSegmentLike[], resolveFont) : []
        const groupStart = new Map<number, WarpRunGroup>()
        const skipInk = new Set<number>()
        // Resolve the SHARED render mode per group BEFORE recording or
        // painting: fallback groups render AND record through the ordinary
        // path (no visual payload), so the two sides can never disagree.
        // The record flow additionally probes scratch allocation per warped
        // group: live host failure resolves to the same ordinary mode on both
        // sides however paint and index are ordered (the probe surface is
        // discarded; paint allocates its own). Paint flow discovers the same
        // failure itself and reports the same specific diagnostic.
        const fallbackGroups = new Set<WarpRunGroup>()
        const allocFailedGroups = new Set<WarpRunGroup>()
        const diagnoseGroupFallback = (): void => {
          const holder = body as { diagnostics?: Array<{ kind: string; feature: string; message: string }> }
          holder.diagnostics ??= []
          if (!holder.diagnostics.some(d => d.feature === 'warp-budget')) {
            holder.diagnostics.push({ kind: 'unsupported-text-warp', feature: 'warp-budget', message: 'Warp raster budget exceeded for a run; rendering unwarped fallback' })
          }
        }
        const diagnoseAllocGroupFallback = (): void => {
          const holder = body as { diagnostics?: Array<{ kind: string; feature: string; message: string }> }
          holder.diagnostics ??= []
          if (!holder.diagnostics.some(d => d.feature === 'warp-alloc-failure')) {
            holder.diagnostics.push({ kind: 'unsupported-text-warp', feature: 'warp-alloc-failure', message: 'Warp scratch surface allocation failed for a run; rendering unwarped fallback' })
          }
        }
        if (warpMapping) {
          for (const group of warpGroups) {
            const plan = planWarpedGroup(warpMapping, boxForMapping, group, segments as unknown as WarpRunSegmentLike[], y + line.baseline, measure)
            if (!plan || plan.fallback) {
              fallbackGroups.add(group)
              continue
            }
            if (record) {
              try {
                const s = scratchSurface(ctx, plan.rasterW, plan.rasterH)
                const r2d = s && typeof (s as any).getContext === 'function' ? (s as any).getContext('2d') : null
                if (!s || !r2d) {
                  allocFailedGroups.add(group)
                }
              } catch {
                allocFailedGroups.add(group)
              }
            }
          }
          if (fallbackGroups.size > 0) diagnoseGroupFallback()
          if (allocFailedGroups.size > 0) diagnoseAllocGroupFallback()
        }
        for (const group of warpGroups) {
          if (group.end > group.start) for (let k = group.start + 1; k <= group.end; k++) skipInk.add(k)
          groupStart.set(group.start, group)
        }
        for (let si = 0; si < segments.length; si++) {
          const segment = segments[si]
          // A source break has no glyph. On a nonempty line its recording band
          // follows visible text, so endParaRPr cannot enlarge hit/selection bands.
          const recordedStyle = segment.text === '\n' ? segments.find(s => s.text !== '\n')?.style ?? segment.style : segment.style
          // Warped runs keep the body frame; their mapped geometry rides the
          // record's visual field, and the group raster internalizes frames.
          // A transformed member whose group takes the RESOLVED ordinary
          // fallback (budget plan or failed allocation probe) records through
          // the frame-aware ordinary path — never unwarped bands for
          // transformed paint.
          const ownerOf = warpGroups.find(g => si >= g.start && si <= g.end)
          const inOrdinaryFallback = !warpMapping || (ownerOf !== undefined && (fallbackGroups.has(ownerOf) || allocFailedGroups.has(ownerOf)))
          const needsFrame = !!segment.transform && inOrdinaryFallback
          const nativeTracking = settings(ctx, recordedStyle, resolveFont)
          // The outer member transform serves recording (placement + visual
          // projection read the context transform) and ordinary paint. Group
          // paint owns its per-member transforms inside paintMemberFallback,
          // so the outer frame must not wrap it a second time.
          const applyOuterFrame = needsFrame && (!!record || !groupStart.has(si))
          if (applyOuterFrame) {
            const t = segment.transform!
            ctx.save()
            ctx.transform(t.a, t.b, t.c, t.d, x + t.e, y + t.f)
          }
          try {
            const sx = needsFrame ? 0 : x + segment.x, sy = needsFrame ? 0 : y + line.baseline
            // Optional mapped visual geometry for curved hit/selection: same
            // logical record, no extra records, canonical UTF-16 untouched.
            // Resolved-fallback groups (shared budget plan or failed
            // allocation probe) record NO visual payload so the replay index
            // matches the actually painted ordinary ink.
            let visual: { polygon: Array<[number, number]>; clusters: Array<[number, number]>; bands?: Array<[number, number, number, number]> } | undefined
            if (warpMapping && segment.text && !/^[\n\t]$/.test(segment.text)) {
              const owner = warpGroups.find(g => si >= g.start && si <= g.end)
              if (owner && !fallbackGroups.has(owner) && !allocFailedGroups.has(owner)) {
                try {
                  visual = segmentWarpGeometry(warpMapping, boxForMapping, segment as unknown as WarpRunSegmentLike, y + line.baseline, measure)
                } catch { /* geometry problems keep the canonical record band */ }
              }
            }
            const logical = { source: sources[line.paragraphIndex], start: segment.sourceStart, end: segment.sourceEnd,
              clip,
              run: segment.runIndex, graphemeBoundaries: segment.graphemeBoundaries,
              ...(visual ? { visual } : {}),
              ...(line.logicalLineIndex === undefined ? {} : { line: line.logicalLineIndex, flow: 'vertical' as const }) }
            if (record) { record(segment.text, sx, sy, segment.width, logical); continue }
            if (skipInk.has(si)) continue
            const group = groupStart.get(si)
            if (group) {
              // One shaped raster for the whole compatible run — contextual
              // Arabic joins and cross-glyph ligatures survive the warp.
              const hasInk = !group.style.noFill || !!group.style.textOutline
              if (hasInk) {
                const paintMemberFallback = (member: WarpRunSegmentLike): void => {
                  const memberNative = settings(ctx, member.style, resolveFont)
                  if (member.transform) {
                    const t = member.transform
                    ctx.save()
                    ctx.transform(t.a, t.b, t.c, t.d, x + t.e, y + t.f)
                  }
                  try { paintSegmentInk(member as unknown as SegmentT, line, memberNative) }
                  finally { if (member.transform) ctx.restore() }
                }
                try {
                  paintWarpedRunGroup({
                    mapping: warpMapping!, ctx, box: boxForMapping, group,
                    segments: segments as unknown as WarpRunSegmentLike[],
                    lineTop: y + line.y, lineBottom: y + line.y + line.height, lineBaseline: y + line.baseline,
                    resolveFont, patternCache, measure,
                    paintFallback: paintMemberFallback,
                    diagnoseFallback: diagnoseGroupFallback,
                    diagnoseAllocFallback: diagnoseAllocGroupFallback,
                  })
                } catch {
                  // Fail-safe: an unexpected warp paint failure falls back
                  // through the FRAME-AWARE ordinary path (member transforms
                  // preserved, so vertical ink stays transformed); never a
                  // throw to callers. Reported as its own outcome, not budget.
                  const holder = body as { diagnostics?: Array<{ kind: string; feature: string; message: string }> }
                  holder.diagnostics ??= []
                  if (!holder.diagnostics.some(d => d.feature === 'warp-paint-failure')) {
                    holder.diagnostics.push({ kind: 'unsupported-text-warp', feature: 'warp-paint-failure', message: 'Warp run paint failed; rendering unwarped fallback' })
                  }
                  for (const member of segments.slice(group.start, group.end + 1)) {
                    paintMemberFallback(member as unknown as WarpRunSegmentLike)
                  }
                }
              }
              continue
            }
            paintSegmentInk(segment, line, nativeTracking)
          } finally { if (applyOuterFrame) ctx.restore() }
        }
      } finally {
        if (options?.clipToLineBox) ctx.restore()
      }
    }
  } finally { ctx.restore() }
}
