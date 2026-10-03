/**
 * Page watermarks.
 *
 * A watermark is painted into the page bitmap rather than laid over it as a DOM
 * layer, so it moves with the document and appears in anything derived from the
 * rendered canvas (a screenshot, a saved page image) rather than being trivially
 * removable by hiding an element.
 *
 * Two ways to supply one:
 *   - the `watermark` prop on <OfficeDoc> / <OfficeFile>, and
 *   - the `X-OfficeView-Watermark` response header on an HTTP download, which
 *     wins over the prop because the server is the authority on what a document
 *     is stamped with.
 *
 * Honest limits, same as copy/print prevention: this is a deterrent. The text
 * is pixels on a canvas, so a reader can still crop it out or capture it with a
 * screenshot. It marks a page; it does not protect a secret inside it.
 */

/** Where the mark is drawn on each page. */
export type WatermarkPlacement = 'center' | 'tile' | 'header' | 'footer'

export interface WatermarkOptions {
  /** The mark itself. Empty/whitespace-only text disables the watermark. */
  text: string
  /** default 'tile' */
  placement?: WatermarkPlacement
  /** Degrees clockwise about the mark's centre. default -45 (rises to the right). */
  rotate?: number
  /** 0..1, clamped. default 0.15 */
  opacity?: number
  /**
   * When set, always wins — including an explicit '#000000' on a dark page.
   * When unset, a colour is derived from the page background at paint time
   * (a light mark on dark pages) so the default stays visible everywhere.
   */
  color?: string
  /**
   * Size in points, matching the font convention used everywhere else here.
   * When omitted it scales with the page width (7% of it) so the mark reads the
   * same on Letter and A4.
   */
  fontSizePt?: number
  fontFamily?: string
}

/** Defaults applied when the caller leaves a field out. */
export const WATERMARK_DEFAULTS = {
  placement: 'tile' as WatermarkPlacement,
  rotate: -45,
  opacity: 0.15,
  color: '#000000',
  fontFamily: 'Helvetica, Arial, sans-serif',
  /** fraction of page width used when no explicit size is given */
  sizeRatioOfPageWidth: 0.07,
  /** extra spacing between tiled marks, as a multiple of the text height */
  tileGap: 1.6,
}

const PLACEMENTS: readonly WatermarkPlacement[] = ['center', 'tile', 'header', 'footer']

/** A resolved watermark: every field present and sane, except colour. */
export interface ResolvedWatermark {
  text: string
  placement: WatermarkPlacement
  rotate: number
  opacity: number
  /**
   * Present only when the caller specified one. An unset colour is resolved at
   * paint time against the page background (see defaultMarkColor) — baking a
   * default in here would make "unset" and "explicitly black" indistinguishable
   * and break black marks on dark slides.
   */
  color?: string
  fontSizePt?: number
  fontFamily: string
}

function clampNumber(v: number, lo: number, hi: number, fallback: number): number {
  if (!Number.isFinite(v)) return fallback
  return Math.min(hi, Math.max(lo, v))
}

/**
 * Fill in defaults and clamp out-of-range values. Returns undefined for a
 * watermark that should not be drawn at all, so callers can treat "off" and
 * "invalid" identically.
 */
export function normalizeWatermark(options: WatermarkOptions | undefined | null): ResolvedWatermark | undefined {
  if (!options) return undefined
  const text = typeof options.text === 'string' ? options.text.trim() : ''
  if (text.length === 0) return undefined
  const placement = PLACEMENTS.includes(options.placement as WatermarkPlacement)
    ? (options.placement as WatermarkPlacement)
    : WATERMARK_DEFAULTS.placement
  return {
    text,
    placement,
    rotate: Number.isFinite(options.rotate) ? Number(options.rotate) : WATERMARK_DEFAULTS.rotate,
    opacity: clampNumber(options.opacity ?? WATERMARK_DEFAULTS.opacity, 0, 1, WATERMARK_DEFAULTS.opacity),
    color: options.color && options.color.trim().length > 0 ? options.color.trim() : undefined,
    fontSizePt: options.fontSizePt !== undefined && options.fontSizePt > 0 ? options.fontSizePt : undefined,
    fontFamily: options.fontFamily && options.fontFamily.trim().length > 0
      ? options.fontFamily.trim()
      : WATERMARK_DEFAULTS.fontFamily,
  }
}

/**
 * Merge a prop watermark with one from the response header. The header wins:
 * a server that stamps a document as CONFIDENTIAL should not be overridable
 * by client props. Returns undefined only when neither side asks for a mark.
 */
export function mergeWatermark(
  fromProp: WatermarkOptions | undefined,
  fromHeader: WatermarkOptions | undefined,
): ResolvedWatermark | undefined {
  if (fromHeader) return normalizeWatermark(fromHeader)
  return normalizeWatermark(fromProp)
}

/** Header name carrying the watermark, matching the protection header's style. */
export const WATERMARK_HEADER = 'X-OfficeView-Watermark'

/**
 * Percent-decoding that cannot throw on malformed input. A bad header must not
 * be able to break a document load.
 */
function safeDecode(value: string): string {
  if (!value.includes('%')) return value
  try {
    return decodeURIComponent(value)
  } catch {
    // decodeURIComponent throws on a stray '%' or a truncated escape; fall back to
    // decoding the sequences that are well formed and keeping the rest.
    return value.replace(/%[0-9A-Fa-f]{2}/g, (seq) => {
      try {
        return decodeURIComponent(seq)
      } catch {
        return seq
      }
    })
  }
}

function parseNumber(value: string | undefined): number | undefined {
  if (value === undefined) return undefined
  const n = Number(value)
  return Number.isFinite(n) ? n : undefined
}

/**
 * Parse `X-OfficeView-Watermark: text=DRAFT; rotate=-45; opacity=0.15`.
 * Pairs are separated by `;` and values are percent-decoded so non-ASCII
 * survives a header.
 *
 * Only `;` separates pairs: treating `,` as one too silently truncated real
 * watermark text ("text=DRAFT, DO NOT COPY" became "DRAFT"). A literal comma
 * still works by encoding it as %2C, because decoding happens after splitting.
 *
 * Returns undefined when the header is absent or carries no text — the common
 * case, and the one that must never affect an existing render.
 */
export function watermarkFromHeaders(headers: Headers): WatermarkOptions | undefined {
  const raw = headers.get(WATERMARK_HEADER)
  if (!raw) return undefined
  const fields: Record<string, string> = {}
  for (const part of raw.split(';')) {
    const eq = part.indexOf('=')
    if (eq <= 0) continue
    const key = part.slice(0, eq).trim().toLowerCase()
    if (key.length > 0) fields[key] = safeDecode(part.slice(eq + 1).trim())
  }
  const text = fields.text
  if (!text || text.trim().length === 0) return undefined
  const placement = fields.placement as WatermarkPlacement | undefined
  return {
    text,
    ...(placement !== undefined ? { placement } : {}),
    ...(parseNumber(fields.rotate) !== undefined ? { rotate: parseNumber(fields.rotate) } : {}),
    ...(parseNumber(fields.opacity) !== undefined ? { opacity: parseNumber(fields.opacity) } : {}),
    ...(fields.color ? { color: fields.color } : {}),
    ...(parseNumber(fields.fontsize) !== undefined ? { fontSizePt: parseNumber(fields.fontsize) } : {}),
    ...(fields.font ? { fontFamily: fields.font } : {}),
  }
}

/** Geometry a painter needs from a page. */
export interface WatermarkPage {
  widthPx: number
  heightPx: number
  /** Page background for contrast derivation; absent means white. */
  background?: string
}

/** One placed mark, in page coordinates, before rotation about its own centre. */
export interface WatermarkStamp {
  xPx: number
  yPx: number
}

/**
 * Where the marks go, in page coordinates. Exported separately from the painter
 * so the layout is testable without a canvas and hosts can reuse it.
 */
export function watermarkStamps(page: WatermarkPage, mark: ResolvedWatermark): WatermarkStamp[] {
  const { widthPx: w, heightPx: h } = page
  const fontPx = fontSizePx(page, mark)
  switch (mark.placement) {
    case 'center':
      return [{ xPx: w / 2, yPx: h / 2 }]
    case 'header':
      return [{ xPx: w / 2, yPx: fontPx * 1.6 }]
    case 'footer':
      return [{ xPx: w / 2, yPx: h - fontPx * 1.2 }]
    case 'tile':
    default: {
      // Step by the mark's OWN measured width, not a guess from the font size.
      // A fixed multiplier overlaps as soon as the text is long enough
      // (CONFIDENTIAL measured ~419px against a ~263px step).
      const textW = markTextWidth(mark, fontPx)
      const stepY = fontPx * (1 + WATERMARK_DEFAULTS.tileGap)
      // +1 so neighbouring marks never share an edge; rotation is accounted for
      // by widening both steps rather than by letting them clip.
      const stepX = Math.max(textW, fontPx) + fontPx * WATERMARK_DEFAULTS.tileGap
      const stamps: WatermarkStamp[] = []
      // start off the top-left corner and step until past the far edges
      for (let y = -stepY; y <= h + stepY; y += stepY) {
        for (let x = -stepX; x <= w + stepX; x += stepX) {
          // offset every other row so the marks interleave like real watermarks
          const row = Math.round((y + stepY) / stepY)
          stamps.push({ xPx: x + (row % 2 === 0 ? 0 : stepX / 2), yPx: y })
        }
      }
      return stamps
    }
  }
}

/**
 * Set by the search-index replay so the watermark does not pollute the text
 * index. A module flag rather than a ctx property because the recorder is a
 * Proxy with no notion of intent, and the replay is synchronous so set/unset
 * around it in a `finally` is safe.
 */
let muted = false

/** Mute/unmute watermark painting. Intended for the search-index replay. */
export function setWatermarkMuted(value: boolean): void {
  muted = value
}

/** True while painting is muted for text-index capture. */
export function isWatermarkMuted(): boolean {
  return muted
}

/** Measured width of a mark's text in CSS px, cached per text+font. */
const textWidthCache = new Map<string, number>()
let measurerCtx: CanvasRenderingContext2D | null = null

/** Registered by the renderer so geometry can measure real text. */
export function setWatermarkMeasurer(ctx: CanvasRenderingContext2D | null): void {
  measurerCtx = ctx
  textWidthCache.clear()
}

function markTextWidth(mark: ResolvedWatermark, fontPx: number): number {
  // fontPx must be the size the CALLER is using: measuring with a different one
  // (a placeholder page) silently under-estimates the width and the tiles overlap.
  const pt = mark.fontSizePt ?? fontPxToPt(fontPx)
  const key = `${pt}|${mark.fontFamily}|${mark.text}`
  const hit = textWidthCache.get(key)
  if (hit !== undefined) return hit
  let w = mark.text.length * fontPx * 0.6 // reasonable fallback without a ctx
  if (measurerCtx) {
    const prev = measurerCtx.font
    measurerCtx.font = `${pt}pt ${quoteFamily(mark.fontFamily)}`
    w = measurerCtx.measureText(mark.text).width
    measurerCtx.font = prev
  }
  textWidthCache.set(key, w)
  return w
}

/** Font size in CSS px: points convert at 96/72, or scales with page width. */
export function fontSizePx(page: WatermarkPage, mark: ResolvedWatermark): number {
  if (mark.fontSizePt !== undefined) return (mark.fontSizePt * 96) / 72
  const scaled = page.widthPx * WATERMARK_DEFAULTS.sizeRatioOfPageWidth
  return Math.min(120, Math.max(12, scaled))
}

type Rgb = [number, number, number]

/** Parse #rgb / #rrggbb (with or without the hash); anything else is undefined. */
function parseHexColor(value: string | undefined): Rgb | undefined {
  if (!value) return undefined
  const hex = value.trim().replace(/^#/, '')
  const full = hex.length === 3 ? hex.split('').map((c) => c + c).join('') : hex
  if (!/^[0-9A-Fa-f]{6}$/.test(full)) return undefined
  return [parseInt(full.slice(0, 2), 16), parseInt(full.slice(2, 4), 16), parseInt(full.slice(4, 6), 16)]
}

function linearChannel(v: number): number {
  const c = v / 255
  return c <= 0.04045 ? c / 12.92 : Math.pow((c + 0.055) / 1.055, 2.4)
}

/** WCAG relative luminance of an sRGB colour, 0 (black) to 1 (white). */
export function relativeLuminance(rgb: Rgb): number {
  return 0.2126 * linearChannel(rgb[0]) + 0.7152 * linearChannel(rgb[1]) + 0.0722 * linearChannel(rgb[2])
}

/**
 * WCAG contrast ratio between two CSS hex colours, 1 (identical) to 21
 * (black on white). Unparseable inputs fall back to the maximum, so callers
 * comparing against a threshold only ever see "no information", never a
 * spurious failure.
 */
export function contrastRatio(a: string, b: string): number {
  const ra = parseHexColor(a)
  const rb = parseHexColor(b)
  if (!ra || !rb) return 21
  const la = relativeLuminance(ra)
  const lb = relativeLuminance(rb)
  const [hi, lo] = la >= lb ? [la, lb] : [lb, la]
  return (hi + 0.05) / (lo + 0.05)
}

/**
 * The default mark colour for a page background: whichever of black and
 * white contrasts more. Dark decks get a light mark, light pages keep black.
 * An absent or unparseable background keeps the black default, so call sites
 * that pass no background (docx/xlsx are always white) behave exactly as
 * before this derivation existed.
 */
export function defaultMarkColor(background?: string): string {
  if (!background || !parseHexColor(background)) return WATERMARK_DEFAULTS.color
  return contrastRatio('#FFFFFF', background) >= contrastRatio(WATERMARK_DEFAULTS.color, background)
    ? '#FFFFFF'
    : WATERMARK_DEFAULTS.color
}

/**
 * Draw the watermark onto a page. Safe to call with an unnormalized options
 * object; anything unusable is simply not drawn.
 *
 * Rotating about each stamp's own centre keeps tiled marks from sweeping away
 * from the page as the angle changes.
 */
export function paintWatermark(
  ctx: CanvasRenderingContext2D,
  page: WatermarkPage,
  options: WatermarkOptions | ResolvedWatermark | undefined,
): void {
  const mark = normalizeWatermark(options as WatermarkOptions | undefined)
  if (!mark || muted) return
  const fontPx = fontSizePx(page, mark)
  // An explicit colour always wins (even black on a dark page); otherwise the
  // mark is derived from the background so the default stays visible.
  const color = validColor(ctx, mark.color, defaultMarkColor(page.background))
  ctx.save()
  ctx.globalAlpha = mark.opacity
  ctx.fillStyle = color
  ctx.font = `${mark.fontSizePt ?? fontPxToPt(fontPx)}pt ${quoteFamily(mark.fontFamily)}`
  ctx.textAlign = 'center'
  ctx.textBaseline = 'middle'
  const rad = (mark.rotate * Math.PI) / 180
  for (const stamp of watermarkStamps(page, mark)) {
    ctx.save()
    ctx.translate(stamp.xPx, stamp.yPx)
    if (rad !== 0) ctx.rotate(rad)
    // textAlign is 'center', so x=0 centres the mark ON the stamp. Offsetting by
    // -width/2 as well would shift it half a text-width to the left and rotate
    // about the text's right edge instead of its centre.
    ctx.fillText(mark.text, 0, 0)
    ctx.restore()
  }
  ctx.restore()
}

/**
 * A canvas silently IGNORES an invalid fillStyle and keeps the previous value,
 * which in the page flow is white — so a bad colour would paint an invisible
 * mark with no error at all. Assign a sentinel first: if the colour is rejected,
 * fillStyle is still the sentinel and we know to fall back.
 */
const PROBE_COLOR = 'rgb(1, 2, 3)'

function validColor(ctx: CanvasRenderingContext2D, color: string | undefined, fallback: string): string {
  // No colour specified: the caller already chose the fallback.
  if (!color) return fallback
  const prev = ctx.fillStyle
  // Assign a KNOWN-GOOD colour first. An invalid assignment is silently ignored
  // and leaves the previous value in place, so that previous value is the tell.
  // (Assigning an invalid sentinel instead would be ignored too, leaving us
  // unable to distinguish "rejected" from "still the old colour".)
  ctx.fillStyle = PROBE_COLOR
  // read the probe back NORMALISED: node-canvas rewrites 'rgb(1, 2, 3)' as
  // '#010203', so comparing raw strings would never match
  const probe = String(ctx.fillStyle)
  ctx.fillStyle = color
  const applied = String(ctx.fillStyle)
  ctx.fillStyle = prev
  return applied === probe && probe !== color ? fallback : color
}

function fontPxToPt(px: number): number {
  return Math.round((px * 72) / 96)
}

/** Quote a font family list for a CSS font shorthand. */
function quoteFamily(family: string): string {
  return family
    .split(',')
    .map((f) => f.trim())
    .filter((f) => f.length > 0)
    .map((f) => (/^[A-Za-z][A-Za-z0-9 ]*$/.test(f) ? f : `"${f}"`))
    .join(', ')
}