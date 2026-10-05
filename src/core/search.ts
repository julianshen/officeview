/**
 * Document text search.
 *
 * The index is built by running each page's paint function against a *recording*
 * context: every `fillText` call is captured with its document-space position.
 * That gives us one geometry-accurate index for all three formats without each
 * renderer having to expose its own text model. The recording canvas is tiny
 * (4x4) — we capture text, not pixels — so indexing is cheap.
 */

import { setWatermarkMuted } from './watermark'
import { createLocalTextMeasurer } from './text-metrics'
import type { HighlightRect } from './overlay'
import { RECORD_TEXT, type LogicalTextRange, type LogicalTextSource } from './text-recording'

export interface TextTransform {
  a: number; b: number; c: number; d: number; e: number; f: number
}

/** Local text band plus the actual Canvas transform, retained for range projection/hit testing. */
export interface TextPlacement {
  x: number
  y: number
  width: number
  top: number
  bottom: number
  transform: TextTransform
  clip?: ReadonlyArray<{ x: number; y: number }>
}

export interface TextSpan {
  text: string
  /** Reading origin of the baseline, in page (document) coordinates. */
  x: number
  /** Baseline y, in page coordinates. */
  y: number
  /** Length along the transformed baseline (positive even when flipped). */
  width: number
  fontSize: number
  /** Optional so existing handwritten indexes retain their geometry contract. */
  placement?: TextPlacement
  /** Optional paragraph-scoped source range; visual wraps do not change these offsets. */
  logical?: LogicalTextRange
}

export interface IndexLine {
  spans: TextSpan[]
  /** Concatenated line text, ordered along its reading direction. */
  text: string
  y: number
  top: number
  bottom: number
}

export interface TextIndexPage {
  index: number
  lines: IndexLine[]
}

export interface TextIndex {
  pages: TextIndexPage[]
}

export interface SearchMatch {
  pageIndex: number
  lineIndex: number
  /** Character offsets within the line's text. */
  start: number
  end: number
  /** Highlight rectangles in page coordinates (one per spanned segment). */
  rects: HighlightRect[]
  /** Present for a source-scoped match, including matches across visual lines. */
  logical?: LogicalTextRange
}

export interface SearchOptions {
  caseSensitive?: boolean
}

/** A 2D context that records fillText calls and forwards everything else. */
export function createRecordingContext(
  base: CanvasRenderingContext2D,
  sink: TextSpan[],
): CanvasRenderingContext2D {
  const measure = createLocalTextMeasurer(base)
  const seenLogical = new WeakMap<LogicalTextSource, Set<string>>()
  let pendingSourceLess: { operation: 'fill' | 'stroke'; key: string } | undefined
  const capture = (text: string, x: number, y: number, maxWidth?: number, knownWidth?: number,
    logical?: LogicalTextRange, operation?: 'fill' | 'stroke'): void => {
    const ctx = base
    const str = String(text)
    // Whitespace spans carry spaces between runs, so they remain indexable.
    if (str.length > 0 || logical) {
      const sizeMatch = /(\d+(?:\.\d+)?)(px|pt)\b/.exec(String(ctx.font ?? '10pt sans-serif'))
      const localFontSize = sizeMatch ? parseFloat(sizeMatch[1]) * (sizeMatch[2] === 'pt' ? 96 / 72 : 1) : 10
      const measured = measure(str).width
      const canonical = logical?.canonical
      const width = canonical?.width ?? knownWidth ?? (maxWidth === undefined ? measured : Math.min(measured, maxWidth))
      const matrix = ctx.getTransform()
      const transform: TextTransform = canonical?.transform ?? { a: matrix.a, b: matrix.b, c: matrix.c, d: matrix.d, e: matrix.e, f: matrix.f }
      const baselineScale = Math.hypot(transform.a, transform.b), normalScale = Math.hypot(transform.c, transform.d)
      const determinant = transform.a * transform.d - transform.b * transform.c
      const recordX = canonical?.x ?? x, recordY = canonical?.y ?? y
      if (![recordX, recordY, width, localFontSize, baselineScale, normalScale, determinant, ...Object.values(transform)].every(Number.isFinite) || width < 0 || determinant === 0) return undefined
      const align = String(ctx.textAlign ?? 'left')
      let left = recordX
      if (align === 'center' || align === 'middle') left -= width / 2
      else if (align === 'right' || align === 'end') left -= width
      // Normalize the local band to an alphabetic baseline before transforming it.
      let baseline = recordY
      if (ctx.textBaseline === 'top' || ctx.textBaseline === 'hanging') baseline += localFontSize * .85
      else if (ctx.textBaseline === 'middle') baseline += localFontSize * .3
      else if (ctx.textBaseline === 'bottom' || ctx.textBaseline === 'ideographic') baseline -= localFontSize * .25
      const ink = logical?.ink
      const measuredInk = ink && Number.isFinite(ink.ascent) && Number.isFinite(ink.descent) && ink.ascent >= 0 && ink.descent >= 0
      const placement: TextPlacement = { x: left, y: baseline, width,
        top: baseline - (measuredInk ? ink.ascent : localFontSize * .85),
        bottom: baseline + (measuredInk ? ink.descent : localFontSize * .25), transform }
      if (logical?.clip) {
        const clip = logical.clip, cm = clip.transform ?? matrix
        if ([clip.x, clip.y, clip.width, clip.height, cm.a, cm.b, cm.c, cm.d, cm.e, cm.f].every(Number.isFinite) && clip.width >= 0 && clip.height >= 0)
          placement.clip = [[clip.x, clip.y], [clip.x + clip.width, clip.y], [clip.x + clip.width, clip.y + clip.height], [clip.x, clip.y + clip.height]].map(([x, y]) => projectTextPoint(cm, x, y))
      }
      const point = projectTextPoint(transform, left, baseline)
      const span: TextSpan = { text: str, x: point.x, y: point.y, width: width * baselineScale, fontSize: localFontSize * normalScale, placement, ...(logical ? { logical } : {}) }
      const bounds = projectedSpanRect(span, 0, 1)
      if ([span.x, span.y, span.width, span.fontSize, bounds.x, bounds.y, bounds.width, bounds.height].every(Number.isFinite)) {
        // Fill, outline and shadow passes may paint one logical source range
        // more than once. Geometry is indexed once for search and copying.
        const placementKey = JSON.stringify([str, placement.x, placement.y, placement.width, transform.a, transform.b, transform.c, transform.d, transform.e, transform.f])
        if (logical) {
          pendingSourceLess = undefined
          const seen = seenLogical.get(logical.source) ?? new Set<string>()
          // Source identity and offsets name the logical glyphs; shadow/fill/
          // outline geometry is merely another appearance of the same glyphs.
          const key = JSON.stringify([str, logical.start, logical.end, logical.line, logical.run])
          if (!seen.has(key)) { sink.push(span); seen.add(key); seenLogical.set(logical.source, seen) }
        } else {
          // An adjacent fill/stroke pair can be one appearance. Equal fills or
          // equal strokes remain independent source-less drawing operations.
          if (operation && pendingSourceLess?.key === placementKey && pendingSourceLess.operation !== operation)
            pendingSourceLess = undefined
          else {
            sink.push(span)
            pendingSourceLess = operation ? { operation, key: placementKey } : undefined
          }
        }
      }
    }
  }
  const proxy = new Proxy(base as unknown as object, {
    get(target, prop) {
      if (prop === RECORD_TEXT) return (text: string, x: number, y: number, width: number, logical: LogicalTextRange) => capture(text, x, y, undefined, width, logical)
      if (prop === 'fillText') return (text: string, x: number, y: number, maxWidth?: number) => capture(text, x, y, maxWidth, undefined, undefined, 'fill')
      if (prop === 'strokeText') return (text: string, x: number, y: number, maxWidth?: number) => capture(text, x, y, maxWidth, undefined, undefined, 'stroke')
      const value = (target as Record<string | symbol, unknown>)[prop]
      return typeof value === 'function' ? value.bind(target) : value
    },
    // assign directly to the target: the default trap would pass the proxy as
    // the receiver, which node-canvas rejects with "Invalid argument"
    set(target, prop, value) {
      ;(target as Record<string | symbol, unknown>)[prop] = value
      return true
    },
  }) as unknown as CanvasRenderingContext2D
  return proxy
}

async function createTinyContext(): Promise<CanvasRenderingContext2D> {
  if (typeof document !== 'undefined') {
    const c = document.createElement('canvas')
    c.width = 4
    c.height = 4
    const ctx = c.getContext('2d')
    if (ctx) return ctx
  }
  // node/bun fallback (golden CLI, scripts) — @vite-ignore keeps this out of
  // browser bundles since it is only reached when document is undefined
  const { createCanvas } = (await import(/* @vite-ignore */ 'canvas')) as unknown as {
    createCanvas: (w: number, h: number) => { getContext: (t: string) => CanvasRenderingContext2D | null }
  }
  const ctx = createCanvas(4, 4).getContext('2d')
  if (!ctx) throw new Error('search: no 2d context available')
  return ctx
}

export function projectTextPoint(transform: TextTransform, x: number, y: number): { x: number; y: number } {
  return { x: transform.a * x + transform.c * y + transform.e, y: transform.b * x + transform.d * y + transform.f }
}

/** Project all four text-band corners; rotated/sheared ranges use bounded page rectangles. */
function projectedSpanRect(span: TextSpan, from: number, to: number): HighlightRect {
  const p = span.placement!
  const points = visibleTextPolygon(p, from, to)
  if (!points.length) return { x: p.clip?.[0]?.x ?? span.x, y: p.clip?.[0]?.y ?? span.y, width: 0, height: 0 }
  const x = Math.min(...points.map(point => point.x)), y = Math.min(...points.map(point => point.y))
  return { x, y, width: Math.max(...points.map(point => point.x)) - x, height: Math.max(...points.map(point => point.y)) - y }
}

export function pointInTextClip(clip: NonNullable<TextPlacement['clip']>, x: number, y: number): boolean {
  const area = clip.reduce((sum, p, i) => { const q = clip[(i + 1) % clip.length]; return sum + p.x * q.y - q.x * p.y }, 0)
  if (Math.abs(area) <= 1e-10) return false
  return clip.every((p, i) => { const q = clip[(i + 1) % clip.length]; return Math.sign(area) * ((q.x - p.x) * (y - p.y) - (q.y - p.y) * (x - p.x)) >= -1e-8 })
}

/** Intersect the actual transformed text band with a convex caller viewport. */
export function visibleTextPolygon(p: TextPlacement, from = 0, to = 1): Array<{ x: number; y: number }> {
  const left = p.x + p.width * from, right = p.x + p.width * to
  let points = [[left, p.top], [right, p.top], [right, p.bottom], [left, p.bottom]].map(([x, y]) => projectTextPoint(p.transform, x, y))
  const clip = p.clip
  if (!clip) return points
  const area = clip.reduce((sum, p, i) => { const q = clip[(i + 1) % clip.length]; return sum + p.x * q.y - q.x * p.y }, 0)
  if (Math.abs(area) <= 1e-10) return []
  for (let i = 0; i < clip.length && points.length; i++) {
    const a = clip[i], b = clip[(i + 1) % clip.length]
    const distance = (p: { x: number; y: number }) => Math.sign(area) * ((b.x - a.x) * (p.y - a.y) - (b.y - a.y) * (p.x - a.x))
    const next: typeof points = []
    for (let j = 0; j < points.length; j++) {
      const p = points[j], q = points[(j + 1) % points.length], dp = distance(p), dq = distance(q)
      const pi = dp >= -1e-10, qi = dq >= -1e-10
      if (pi) next.push(p)
      if (pi !== qi) { const t = dp / (dp - dq); next.push({ x: p.x + t * (q.x - p.x), y: p.y + t * (q.y - p.y) }) }
    }
    points = next
  }
  return points
}

/** Merge collinear baselines and order along their directed axis, including rotation/reflection. */
function spansToLines(spans: TextSpan[]): IndexLine[] {
  const buckets = new Map<string, { spans: TextSpan[]; ux: number; uy: number }>()
  const sourceIds = new Map<LogicalTextSource, number>()
  for (const span of spans) {
    if (span.logical && !sourceIds.has(span.logical.source)) sourceIds.set(span.logical.source, sourceIds.size)
    const transform = span.placement?.transform
    const scale = transform ? Math.hypot(transform.a, transform.b) : 1
    const ux = transform ? transform.a / scale : 1, uy = transform ? transform.b / scale : 0
    const normalOffset = -uy * span.x + ux * span.y
    const key = span.logical?.line !== undefined
      ? `${sourceIds.get(span.logical.source)}:source-line:${span.logical.line}`
      : `${span.logical ? sourceIds.get(span.logical.source) : 'legacy'}:${Math.round(ux * 1e6)},${Math.round(uy * 1e6)}:${Math.round(normalOffset)}`
    const bucket = buckets.get(key) ?? { spans: [], ux, uy }
    bucket.spans.push(span)
    buckets.set(key, bucket)
  }
  const lines: IndexLine[] = []
  for (const { spans: bucket, ux, uy } of buckets.values()) {
    bucket.sort((a, b) => a.logical && b.logical ? a.logical.start - b.logical.start : (a.x - b.x) * ux + (a.y - b.y) * uy)
    const bounds = bucket.map(span => span.placement ? projectedSpanRect(span, 0, 1) : { x: span.x, y: span.y - span.fontSize * .85, width: span.width, height: span.fontSize * 1.1 })
    lines.push({ spans: bucket, text: bucket.map(span => span.text).join(''), y: bucket[0].y,
      top: Math.min(...bounds.map(rect => rect.y)), bottom: Math.max(...bounds.map(rect => rect.y + rect.height)),
    })
  }
  lines.sort((a, b) => a.top - b.top || a.spans[0].x - b.spans[0].x)
  // CaretPos line order remains spatial between unrelated scopes and for
  // legacy painters. Within a recorded paragraph it follows source order,
  // including baselines reversed by rotation/reflection. Replace only that
  // source's occupied slots rather than globally sorting text boxes/cells.
  const sourceSlots = new Map<object, number[]>()
  lines.forEach((line, i) => {
    const source = line.spans[0]?.logical?.source
    if (!source || !line.spans.every(span => span.logical?.source === source)) return
    const group = source.scope ?? source
    const slots = sourceSlots.get(group) ?? []
    slots.push(i); sourceSlots.set(group, slots)
  })
  for (const slots of sourceSlots.values()) {
    const ordered = slots.map(i => lines[i]).sort((a, b) =>
      (a.spans[0].logical!.source.order ?? 0) - (b.spans[0].logical!.source.order ?? 0) ||
      a.spans[0].logical!.start - b.spans[0].logical!.start)
    slots.forEach((slot, i) => { lines[slot] = ordered[i] })
  }
  return lines
}

export interface PaintableLike {
  spec: { widthPx: number; heightPx: number }
  paint: (ctx: CanvasRenderingContext2D) => void
}

/** Build the text index for a document by replaying its paint functions. */
export async function buildTextIndex(paintables: PaintableLike[]): Promise<TextIndex> {
  const ctx = await createTinyContext()
  // The watermark is painted into the same bitmaps the reader sees, so replaying
  // them would capture its text too: searching the mark would match every page,
  // and mark baselines that round-collide with text lines would merge into the
  // indexed line text, corrupting hits, highlight rects and copy. The replay is
  // synchronous, so set/unset around it in a finally is safe.
  setWatermarkMuted(true)
  try {
    const pages: TextIndexPage[] = paintables.map((p, index) => {
      const spans: TextSpan[] = []
      // Each page starts in the caller's original state, even if a painter translates without restoring.
      ctx.save()
      try { p.paint(createRecordingContext(ctx, spans)) } finally { ctx.restore() }
      return { index, lines: spansToLines(spans) }
    })
    return { pages }
  } finally {
    setWatermarkMuted(false)
  }
}

/**
 * Rectangles covering [start, end) of a line. Character offsets are mapped
 * proportionally within each span — an approximation that is exact for
 * single-span matches and close for text spanning several runs.
 */
export function rectsForRange(line: IndexLine, start: number, end: number): HighlightRect[] {
  const height = Math.max(2, line.bottom - line.top)
  const out: HighlightRect[] = []
  let offset = 0
  for (const span of line.spans) {
    const spanStart = offset
    const spanEnd = offset + span.text.length
    offset = spanEnd
    if (end <= spanStart || start >= spanEnd) continue
    const from = Math.max(0, start - spanStart)
    const to = Math.min(span.text.length, end - spanStart)
    if (to <= from) continue
    const len = span.text.length || 1
    const x = span.x + (from / len) * span.width
    const width = ((to - from) / len) * span.width
    const rect = span.placement ? projectedSpanRect(span, from / len, to / len) : { x, y: line.top, width, height }
    if (rect.width > 0 && rect.height > 0) out.push(rect)
  }
  return out
}

/** Find every occurrence of `query` across the indexed pages. */
export function findMatches(index: TextIndex, query: string, options: SearchOptions = {}): SearchMatch[] {
  const needle = options.caseSensitive ? query : query.toLowerCase()
  if (needle.length === 0) return []
  const matches: SearchMatch[] = []
  for (const page of index.pages) {
    const sources = new Map<LogicalTextSource, Array<{ span: TextSpan; lineIndex: number; offset: number }>>()
    page.lines.forEach((line, lineIndex) => {
      let offset = 0
      for (const span of line.spans) {
        if (span.logical) {
          const ranges = sources.get(span.logical.source) ?? []
          ranges.push({ span, lineIndex, offset }); sources.set(span.logical.source, ranges)
        }
        offset += span.text.length
      }
    })
    for (const [source, ranges] of sources) {
      ranges.sort((a, b) => a.span.logical!.start - b.span.logical!.start)
      const haystack = options.caseSensitive ? source.text : source.text.toLowerCase()
      let from = 0
      for (;;) {
        const at = haystack.indexOf(needle, from)
        if (at < 0) break
        const end = at + needle.length
        const covered = ranges.filter(r => r.span.logical!.end > at && r.span.logical!.start < end)
        if (covered.length) {
          const first = covered[0], rects: HighlightRect[] = []
          for (const r of covered) {
            const logical = r.span.logical!
            rects.push(...rectsForRange(page.lines[r.lineIndex], r.offset + Math.max(0, at - logical.start), r.offset + Math.min(r.span.text.length, end - logical.start)))
          }
          const start = first.offset + Math.max(0, at - first.span.logical!.start)
          matches.push({ pageIndex: page.index, lineIndex: first.lineIndex, start,
            end: Math.min(page.lines[first.lineIndex].text.length, start + needle.length), rects, logical: { source, start: at, end } })
        }
        from = end
      }
    }
    for (let li = 0; li < page.lines.length; li++) {
      const line = page.lines[li]
      if (line.spans.some(span => span.logical)) continue
      const haystack = options.caseSensitive ? line.text : line.text.toLowerCase()
      let from = 0
      for (;;) {
        const at = haystack.indexOf(needle, from)
        if (at < 0) break
        const end = at + needle.length
        matches.push({
          pageIndex: page.index,
          lineIndex: li,
          start: at,
          end,
          rects: rectsForRange(line, at, end),
        })
        from = end
      }
    }
  }
  return matches
}

/** Wrap-around navigation helper: step through matches by ±1. */
export function stepMatch(matches: SearchMatch[], current: number, direction: 1 | -1): number {
  if (matches.length === 0) return -1
  return (current + direction + matches.length) % matches.length
}
