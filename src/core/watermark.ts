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
  /** default '#000000' */
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

/** A resolved watermark: every field present and sane. */
export interface ResolvedWatermark {
  text: string
  placement: WatermarkPlacement
  rotate: number
  opacity: number
  color: string
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
    color: options.color && options.color.trim().length > 0 ? options.color.trim() : WATERMARK_DEFAULTS.color,
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
 * Pairs are separated by `;` (or `,`, for symmetry with the protection
 * header) and values are percent-decoded so non-ASCII survives a header.
 *
 * Returns undefined when the header is absent or carries no text — the common
 * case, and the one that must never affect an existing render.
 */
export function watermarkFromHeaders(headers: Headers): WatermarkOptions | undefined {
  const raw = headers.get(WATERMARK_HEADER)
  if (!raw) return undefined
  const fields: Record<string, string> = {}
  for (const part of raw.split(/[;,]/)) {
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
      // Spacing is derived from the mark's own height so tiles neither overlap
      // nor leave a gap that looks accidental; rotated marks need extra room,
      // hence the tileGap multiplier on both axes.
      const stepY = fontPx * (1 + WATERMARK_DEFAULTS.tileGap)
      const stepX = fontPx * (3 + WATERMARK_DEFAULTS.tileGap)
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

/** Font size in CSS px: points convert at 96/72, or scales with page width. */
export function fontSizePx(page: WatermarkPage, mark: ResolvedWatermark): number {
  if (mark.fontSizePt !== undefined) return (mark.fontSizePt * 96) / 72
  const scaled = page.widthPx * WATERMARK_DEFAULTS.sizeRatioOfPageWidth
  return Math.min(120, Math.max(12, scaled))
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
  if (!mark) return
  const fontPx = fontSizePx(page, mark)
  ctx.save()
  ctx.globalAlpha = mark.opacity
  ctx.fillStyle = mark.color
  ctx.font = `${mark.fontSizePt ?? fontPxToPt(fontPx)}pt ${quoteFamily(mark.fontFamily)}`
  ctx.textAlign = 'center'
  ctx.textBaseline = 'middle'
  const rad = (mark.rotate * Math.PI) / 180
  for (const stamp of watermarkStamps(page, mark)) {
    ctx.save()
    ctx.translate(stamp.xPx, stamp.yPx)
    if (rad !== 0) ctx.rotate(rad)
    const w = ctx.measureText(mark.text).width
    ctx.fillText(mark.text, -w / 2, 0)
    ctx.restore()
  }
  ctx.restore()
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