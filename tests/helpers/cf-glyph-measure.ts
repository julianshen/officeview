/**
 * Task 28 (OD-9(b)): measures our original vector glyphs against Excel 16.115's glyph-key capture at the key's own
 * geometry, with the harness's own ink rule. The key crops are comparison fixtures only: nothing here feeds pixels back
 * into the source.
 *
 * - The key cell: 72 pt wide, 20 pt tall, 11 pt Calibri, bottom-aligned, a right-aligned number under `;;;`
 *   (`scripts/cf-refs/cases.ts` stage4-icon-glyph-key; sidecar `colWidthsPt`/`rowHeightsPt`), laid out at the capture
 *   scale (1 device px per pt, as the sidecar's `pxPerPt` 1).
 * - Ink: a pixel whose max channel delta from the cell's own background (white) reaches 64 (`ICON_STRONG_DELTA`,
 *   S4-A), inside the cell's inner box (gridlines excluded), as `measureIconGeometry` reads it.
 * - A region: the pixels one fill owns (coverage >= 1/2 of its own path, < 1/2 of every later fill's path). Its
 *   dominant colour in the crop is the per-channel median of the crop over the region.
 */
import { createCanvas } from 'canvas'
import { CF_ICON_RENDER_MODEL, layoutIcon } from '../../src/xlsx/cf/iconPaint'
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
const ALL = (): boolean => true
const admitAll = (glyph: CfIconGlyph): CfGlyphCache => ({ classes: new Set([glyph.glyphClass]), bytes: 0, release() {} })

/** The glyph drawn into a white key cell (`KEY_CELL_W` × `KEY_CELL_H` device px) through `layoutIcon` + `drawIconGlyph`. */
export function renderAtKey(glyph: CfIconGlyph, background = '#ffffff'): Uint8ClampedArray {
  const canvas = createCanvas(KEY_CELL_W, KEY_CELL_H)
  const ctx = canvas.getContext('2d') as unknown as CanvasRenderingContext2D
  ctx.fillStyle = background
  ctx.fillRect(0, 0, KEY_CELL_W, KEY_CELL_H)
  const laid = layoutIcon({ ruleId: 1, set: '3Arrows', iconId: 0, bucket: 0, buckets: 3, showValue: true }, keyHost(), KEY_FONT, AT_CAPTURE, CF_ICON_RENDER_MODEL, ALL)
  if (!laid) throw new Error('the production model lays out no icon at the key geometry')
  drawIconGlyph(ctx, glyph, laid.box, AT_CAPTURE, admitAll(glyph))
  return ctx.getImageData(0, 0, KEY_CELL_W, KEY_CELL_H).data
}

/** Coverage (0..1) of one fill's path alone, per key-cell pixel. */
function coverageOf(glyph: CfIconGlyph, ops: CfIconGlyph['path']): Float64Array {
  const canvas = createCanvas(KEY_CELL_W, KEY_CELL_H)
  const ctx = canvas.getContext('2d') as unknown as CanvasRenderingContext2D
  const laid = layoutIcon({ ruleId: 1, set: '3Arrows', iconId: 0, bucket: 0, buckets: 3, showValue: true }, keyHost(), KEY_FONT, AT_CAPTURE, CF_ICON_RENDER_MODEL, ALL)!
  drawIconGlyph(ctx, { ...glyph, path: ops }, laid.box, AT_CAPTURE, admitAll(glyph))
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
