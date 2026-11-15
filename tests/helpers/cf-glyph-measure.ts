/**
 * Task 28 (OD-9(b)): measures our original vector glyphs against Excel 16.115's glyph-key capture at the key's own
 * geometry, with the harness's own ink rule. The key crops are comparison fixtures only: nothing here feeds pixels back
 * into the source.
 *
 * - The key geometry: 72 pt wide, 20 pt tall, 11 pt Calibri, bottom-aligned (`scripts/cf-refs/cases.ts`
 *   stage4-icon-glyph-key; sidecar `colWidthsPt`/`rowHeightsPt`), laid out at the capture scale (1 device px per pt, as
 *   the sidecar's `pxPerPt` 1). The key's own cells are right-aligned numbers under `;;;`, which production still
 *   withholds (`rightInsetPt` is unmeasured); `KEY_FONT` lays the glyph out as left-aligned text, which the frame
 *   admits and which does not move the icon (geometry-2 I5/J5).
 * - Ink: a pixel whose max channel delta from the cell's own background (white) reaches 64 (`ICON_STRONG_DELTA`,
 *   S4-A), inside the cell's inner box (gridlines excluded), as `measureIconGeometry` reads it.
 * - A region: the pixels one fill owns (coverage >= 1/2 of its own path, < 1/2 of every later fill's path). Its
 *   dominant colour in the crop is the per-channel median of the crop over the region.
 */
import { createCanvas } from 'canvas'
import { drawIconGlyph, type CfGlyphCache, type CfIconGlyph } from '../../src/xlsx/cf/iconGlyphs'
import { ICON_CLASS_TOLERANCE, ICON_STRONG_DELTA, type DecodedPng } from '../../scripts/cf-refs/verify'

export const KEY_CELL_W = 72
export const KEY_CELL_H = 20
export const AT_CAPTURE = { zoom: 0.75, dpr: 1, output: 'device' as const }
const PT_PX = 4 / 3
/** The key cell as the viewer lays it out (logical px): the number 10, general alignment (right), 11 pt Calibri. */
export const keyHost = (overrides: Record<string, unknown> = {}) => ({
  x: 0, y: 0, width: KEY_CELL_W * PT_PX, height: KEY_CELL_H * PT_PX, textWidth: 0,
  cell: { value: 10, style: { numFmtId: 0, fontSizePt: 11, fontName: 'Calibri' } }, ...overrides,
})
export const KEY_FONT = { sizePt: 11, vertical: 'bottom' as const, horizontal: 'left' as const }
/** The frame at the glyph-key cell (`CF_ICON_RENDER_MODEL.keyFrame`: 13 × 12, 2 px in; the classes' ink rows 7..18 at
 * 1 px/pt). The approximation check draws there directly: it measures the artwork, not where production admits it
 * (S4-ROW-EXACT withholds the key's own 20 pt row, whose viewer px differ from Excel's). */
export const KEY_BOX = Object.freeze({ x: 2, y: 7, w: 13, h: 12 })
const admitAll = (glyph: CfIconGlyph): CfGlyphCache => ({ classes: new Set([glyph.glyphClass]), bytes: 0, release() {} })

/** The glyph drawn into a white key cell (`KEY_CELL_W` × `KEY_CELL_H` device px) in the key frame (`drawIconGlyph`). */
export function renderAtKey(glyph: CfIconGlyph, background = '#ffffff'): Uint8ClampedArray {
  const canvas = createCanvas(KEY_CELL_W, KEY_CELL_H)
  const ctx = canvas.getContext('2d') as unknown as CanvasRenderingContext2D
  ctx.fillStyle = background
  ctx.fillRect(0, 0, KEY_CELL_W, KEY_CELL_H)
  drawIconGlyph(ctx, glyph, KEY_BOX, AT_CAPTURE, admitAll(glyph))
  return ctx.getImageData(0, 0, KEY_CELL_W, KEY_CELL_H).data
}

/** Coverage (0..1) of one fill's path alone, per key-cell pixel. */
function coverageOf(glyph: CfIconGlyph, ops: CfIconGlyph['path']): Float64Array {
  const canvas = createCanvas(KEY_CELL_W, KEY_CELL_H)
  const ctx = canvas.getContext('2d') as unknown as CanvasRenderingContext2D
  drawIconGlyph(ctx, { ...glyph, path: ops }, KEY_BOX, AT_CAPTURE, admitAll(glyph))
  const data = ctx.getImageData(0, 0, KEY_CELL_W, KEY_CELL_H).data
  const out = new Float64Array(KEY_CELL_W * KEY_CELL_H)
  for (let i = 0; i < out.length; i++) out[i] = data[i * 4 + 3]! / 255
  return out
}

export type GlyphRegion = { argb: number; pixels: Array<{ x: number; y: number }> }
/** Each fill's owned pixels (cell coordinates), in paint order. */
export function regionsOf(glyph: CfIconGlyph): GlyphRegion[] {
  const fills: Array<{ argb: number; ops: CfIconGlyph['path'] }> = []
  let current: CfIconGlyph['path'][number][] = []
  for (const op of glyph.path) {
    if (op[0] === 'F') { fills.push({ argb: op[1], ops: [...current, ['F', 0xff000000]] }); current = [] } else current.push(op)
  }
  const cov = fills.map(f => coverageOf(glyph, f.ops))
  return fills.map((f, i) => {
    const pixels: Array<{ x: number; y: number }> = []
    for (let y = 0; y < KEY_CELL_H; y++) for (let x = 0; x < KEY_CELL_W; x++) {
      const p = y * KEY_CELL_W + x
      if (cov[i]![p]! >= 0.5 && cov.slice(i + 1).every(c => c[p]! < 0.5)) pixels.push({ x, y })
    }
    return { argb: f.argb, pixels }
  })
}

export type InkBox = { left: number; top: number; right: number; bottom: number }
const delta = (r: number, g: number, b: number): number => Math.max(255 - r, 255 - g, 255 - b)
/** Inclusive ink box, cell-relative, of an RGBA cell image over its inner box, at max delta >= `threshold` from white. */
export function inkBoxOf(data: Uint8ClampedArray, width: number, height: number, threshold = ICON_STRONG_DELTA, at: (x: number, y: number) => number = (x, y) => (y * width + x) * 4): InkBox | undefined {
  let left = Infinity, top = Infinity, right = -1, bottom = -1
  for (let y = 1; y < height - 1; y++) for (let x = 1; x < width - 1; x++) {
    const k = at(x, y)
    if (delta(data[k]!, data[k + 1]!, data[k + 2]!) < threshold) continue
    left = Math.min(left, x); right = Math.max(right, x); top = Math.min(top, y); bottom = Math.max(bottom, y)
  }
  return right < 0 ? undefined : { left, top, right, bottom }
}

/** A key crop (the inner box's left 28 px: crop (0,0) is cell (1,1)) read at cell coordinates. */
export const cropAt = (crop: DecodedPng, x: number, y: number): [number, number, number] | undefined => {
  const cx = x - 1, cy = y - 1
  if (cx < 0 || cy < 0 || cx >= crop.width || cy >= crop.height) return undefined
  const k = (cy * crop.width + cx) * 4
  return [crop.data[k]!, crop.data[k + 1]!, crop.data[k + 2]!]
}
/** The crop's ink box in cell coordinates at `threshold` (a faint class is read at the S4-A tolerance). */
export function cropInkBox(crop: DecodedPng, threshold: number): InkBox | undefined {
  const box = inkBoxOf(crop.data, crop.width + 2, crop.height + 2, threshold, (x, y) => {
    const cx = Math.min(Math.max(x - 1, 0), crop.width - 1), cy = Math.min(Math.max(y - 1, 0), crop.height - 1)
    return (cy * crop.width + cx) * 4
  })
  return box
}

const median = (v: number[]): number => {
  const s = [...v].sort((a, b) => a - b)
  const m = s.length >> 1
  return s.length % 2 ? s[m]! : Math.round((s[m - 1]! + s[m]!) / 2)
}
/** The crop's dominant colour over a region: the per-channel median. */
export function cropMedian(crop: DecodedPng, pixels: ReadonlyArray<{ x: number; y: number }>): [number, number, number] | undefined {
  const ch: [number[], number[], number[]] = [[], [], []]
  for (const { x, y } of pixels) {
    const p = cropAt(crop, x, y)
    if (!p) continue
    ch[0].push(p[0]); ch[1].push(p[1]); ch[2].push(p[2])
  }
  return ch[0].length ? [median(ch[0]), median(ch[1]), median(ch[2])] : undefined
}

export type GlyphCheck = { ok: boolean; reasons: string[]; drawn?: InkBox; measured?: InkBox; regions: Array<{ argb: number; size: number; median?: [number, number, number] }> }
/**
 * The OD-9(b) approximation check of one glyph: its drawn ink box equals the measured one within ±1 px per edge, and
 * every region's fill equals the crop's dominant colour within ±8 per channel. `faint` (S4-FAINT) reads both boxes at
 * the S4-A tolerance and requires that the glyph draws no ink at the strong delta, as Excel's does not.
 */
export function checkGlyph(glyph: CfIconGlyph, crop: DecodedPng, measured: InkBox | undefined, faint: boolean): GlyphCheck {
  const reasons: string[] = []
  const img = renderAtKey(glyph)
  // Fix round (review P1): the reference's own regions and ink, independent of the candidate's paths.
  const iou = inkIoU(img, crop, faint ? ICON_CLASS_TOLERANCE + 1 : ICON_STRONG_DELTA)
  if (iou < INK_IOU_MIN) reasons.push(`ink IoU ${iou.toFixed(3)} < ${INK_IOU_MIN}`)
  for (const r of referenceRegions(crop)) {
    const a = regionAgreement(img, r, REGION_TOLERANCE)
    if (a < REGION_AGREEMENT_MIN) reasons.push(`reference region ${r.median.join(',')} (${r.pixels.length} px) agrees on ${Math.round(a * 100)}% < ${REGION_AGREEMENT_MIN * 100}%`)
  }
  const drawn = inkBoxOf(img, KEY_CELL_W, KEY_CELL_H, faint ? ICON_CLASS_TOLERANCE + 1 : ICON_STRONG_DELTA)
  if (faint && inkBoxOf(img, KEY_CELL_W, KEY_CELL_H, ICON_STRONG_DELTA)) reasons.push('a faint class draws strong ink')
  if (!drawn || !measured) reasons.push(`no ${!drawn ? 'drawn' : 'measured'} ink box`)
  else for (const edge of ['left', 'top', 'right', 'bottom'] as const) if (Math.abs(drawn[edge] - measured[edge]) > 1) reasons.push(`${edge} ${drawn[edge]} vs ${measured[edge]}`)
  const regions = regionsOf(glyph).map(r => {
    const med = cropMedian(crop, r.pixels)
    const rgb = [(r.argb >>> 16) & 255, (r.argb >>> 8) & 255, r.argb & 255]
    if (!med) reasons.push(`a fill ${r.argb.toString(16)} owns no pixel`)
    else if (rgb.some((c, i) => Math.abs(c - med[i]!) > 8)) reasons.push(`fill ${r.argb.toString(16)} vs crop median ${med.join(',')}`)
    return { argb: r.argb, size: r.pixels.length, ...(med ? { median: med } : {}) }
  })
  return { ok: reasons.length === 0, reasons, ...(drawn ? { drawn } : {}), ...(measured ? { measured } : {}), regions }
}

/**
 * Fix round (review P1): reference-derived semantic regions, frozen from the key crop alone and independent of any
 * candidate's paths. The crop's visible pixels (max delta > 8 from white) are grouped by colour: each pixel joins the
 * first region whose seed colour lies within `REGION_RADIUS` per channel, seeds taken in order of how many pixels share
 * a colour; a region of fewer than `REGION_MIN_PX` pixels is noise. Each region carries its median colour.
 */
export const REGION_RADIUS = 48
/**
 * The approximation thresholds, fixed from the negative controls before any real glyph was measured (fix round):
 * a solid rectangle in the crop's median colour over the measured ink box, substituted for 3TrafficLights1:0 and each
 * 3Symbols glyph, scores ink IoU 0.835 / 0.736 / 0.777 / 0.736 and leaves one reference region at 0 % agreement.
 * - `INK_IOU_MIN` 0.85: above every control (a box over a round glyph is about π/4 of it plus edges), so substituting a
 *   box for a round or symbol glyph fails on shape alone.
 * - `REGION_AGREEMENT_MIN` 0.5 at `REGION_TOLERANCE` ±48 per channel: each reference region must be drawn in its own
 *   colour on a majority of its pixels; every control has a region at 0 % (the rim, or the symbol, drawn in the wrong
 *   colour), so a vanished symbol or rim fails.
 */
export const INK_IOU_MIN = 0.85
export const REGION_TOLERANCE = 48
export const REGION_AGREEMENT_MIN = 0.5
export const REGION_MIN_PX = 3
export type ReferenceRegion = { pixels: Array<{ x: number; y: number }>; median: [number, number, number] }
export function referenceRegions(crop: DecodedPng): ReferenceRegion[] {
  const px: Array<{ x: number; y: number; c: [number, number, number] }> = []
  for (let cy = 0; cy < crop.height; cy++) for (let cx = 0; cx < crop.width; cx++) {
    const k = (cy * crop.width + cx) * 4
    const c: [number, number, number] = [crop.data[k]!, crop.data[k + 1]!, crop.data[k + 2]!]
    if (delta(...c) > ICON_CLASS_TOLERANCE) px.push({ x: cx + 1, y: cy + 1, c })
  }
  const freq = new Map<string, number>()
  for (const p of px) freq.set(p.c.join(), (freq.get(p.c.join()) ?? 0) + 1)
  const seeds = [...new Set(px.map(p => p.c.join()))].sort((a, b) => freq.get(b)! - freq.get(a)! || (a < b ? -1 : 1)).map(s => s.split(',').map(Number) as [number, number, number])
  const near = (a: readonly number[], b: readonly number[]) => a.every((v, i) => Math.abs(v - b[i]!) <= REGION_RADIUS)
  const groups: Array<{ seed: [number, number, number]; pixels: typeof px }> = []
  for (const p of px) {
    let g = groups.find(x => near(x.seed, p.c))
    if (!g) { const seed = seeds.find(s => near(s, p.c))!; g = groups.find(x => x.seed === seed) ?? (groups.push({ seed, pixels: [] }), groups[groups.length - 1]!) }
    g.pixels.push(p)
  }
  return groups.filter(g => g.pixels.length >= REGION_MIN_PX).map(g => ({
    pixels: g.pixels.map(({ x, y }) => ({ x, y })),
    median: [median(g.pixels.map(p => p.c[0])), median(g.pixels.map(p => p.c[1])), median(g.pixels.map(p => p.c[2]))],
  }))
}
/** The share of a region's pixels where the candidate image's colour lies within `tol` per channel of its median. */
export function regionAgreement(img: Uint8ClampedArray, region: ReferenceRegion, tol: number): number {
  let hit = 0
  for (const { x, y } of region.pixels) {
    const k = (y * KEY_CELL_W + x) * 4
    if ([img[k]!, img[k + 1]!, img[k + 2]!].every((v, i) => Math.abs(v - region.median[i]!) <= tol)) hit++
  }
  return hit / region.pixels.length
}
/** Intersection over union of the candidate's and the crop's ink masks (max delta >= `threshold`), cell coordinates. */
export function inkIoU(img: Uint8ClampedArray, crop: DecodedPng, threshold: number): number {
  let inter = 0, union = 0
  for (let y = 1; y < KEY_CELL_H - 1; y++) for (let x = 1; x < KEY_CELL_W - 1; x++) {
    const k = (y * KEY_CELL_W + x) * 4
    const a = delta(img[k]!, img[k + 1]!, img[k + 2]!) >= threshold
    const p = cropAt(crop, x, y)
    const b = p !== undefined && delta(...p) >= threshold
    if (a && b) inter++
    if (a || b) union++
  }
  return union ? inter / union : 1
}

/**
 * Round 8 (S4-NO-FIT): a glyph's fills, each with its ops and area (px² in frame points at the key scale), for the
 * ablation test. A "patch" is a fill of at most `PATCH_MAX_AREA` px².
 */
export const PATCH_MAX_AREA = 6
export function glyphFills(glyph: CfIconGlyph): Array<{ ops: CfIconGlyph['path']; area: number }> {
  const out: Array<{ ops: CfIconGlyph['path']; area: number }> = []
  let ops: CfIconGlyph['path'][number][] = []
  for (const op of glyph.path) {
    ops.push(op)
    if (op[0] !== 'F') continue
    let area = 0
    const pts: Array<[number, number]> = []
    for (const o of ops) {
      if (o[0] === 'M' || o[0] === 'L') pts.push([o[1], o[2]])
      if (o[0] === 'A') area += Math.PI * o[3] * o[3] * Math.min(1, Math.abs(o[5] - o[4]) / (2 * Math.PI))
    }
    for (let i = 0; i < pts.length; i++) { const a = pts[i]!, b = pts[(i + 1) % pts.length]!; area += (a[0] * b[1] - b[0] * a[1]) / 2 }
    out.push({ ops, area: Math.abs(area) })
    ops = []
  }
  return out
}
/** The glyph without fill `index` (its ops removed). */
export function withoutFill(glyph: CfIconGlyph, index: number): CfIconGlyph {
  return { ...glyph, path: glyphFills(glyph).filter((_, i) => i !== index).flatMap(f => f.ops) }
}
/**
 * A patch is fitted when removing it barely moves the ink IoU (< `FIT_IOU_DELTA`) yet flips a reference region from
 * passing (>= 50 %) to failing: it exists only to colour a region of the map, not to draw the shape.
 */
export const FIT_IOU_DELTA = 0.01
export function fittedPatches(glyph: CfIconGlyph, crop: DecodedPng, faint: boolean): number[] {
  const t = faint ? ICON_CLASS_TOLERANCE + 1 : ICON_STRONG_DELTA
  const regions = referenceRegions(crop)
  const base = renderAtKey(glyph)
  const baseIoU = inkIoU(base, crop, t)
  const basePass = regions.map(r => regionAgreement(base, r, REGION_TOLERANCE) >= REGION_AGREEMENT_MIN)
  const fitted: number[] = []
  glyphFills(glyph).forEach((f, i) => {
    if (f.area > PATCH_MAX_AREA) return
    const img = renderAtKey(withoutFill(glyph, i))
    const flips = regions.some((r, k) => basePass[k] && regionAgreement(img, r, REGION_TOLERANCE) < REGION_AGREEMENT_MIN)
    if (flips && Math.abs(inkIoU(img, crop, t) - baseIoU) < FIT_IOU_DELTA) fitted.push(i)
  })
  return fitted
}
