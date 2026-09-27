/**
 * Document text search.
 *
 * The index is built by running each page's paint function against a *recording*
 * context: every `fillText` call is captured with its document-space position.
 * That gives us one geometry-accurate index for all three formats without each
 * renderer having to expose its own text model. The recording canvas is tiny
 * (4x4) — we capture text, not pixels — so indexing is cheap.
 */

export interface TextSpan {
  text: string
  /** Left edge of the span, in page (document) coordinates. */
  x: number
  /** Baseline y, in page coordinates. */
  y: number
  width: number
  fontSize: number
}

export interface IndexLine {
  spans: TextSpan[]
  /** Concatenated line text, in visual order. */
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
  rects: Array<{ x: number; y: number; width: number; height: number }>
}

export interface SearchOptions {
  caseSensitive?: boolean
}

type AnyCtx = CanvasRenderingContext2D & Record<string | symbol, unknown>

/** A 2D context that records fillText calls and forwards everything else. */
export function createRecordingContext(
  base: CanvasRenderingContext2D,
  sink: TextSpan[],
): CanvasRenderingContext2D {
  const proxy = new Proxy(base as unknown as object, {
    get(target, prop) {
      if (prop === 'fillText') {
        return (text: string, x: number, y: number) => {
          const ctx = target as unknown as AnyCtx
          const str = String(text)
          // keep whitespace-only spans: they carry the spaces between words,
          // which the line text (and therefore search) depends on
          if (str.length > 0) {
            const font = String(ctx.font ?? '10pt sans-serif')
            const sizeMatch = /(\d+(?:\.\d+)?)px/.exec(font)
            const fontSize = sizeMatch ? parseFloat(sizeMatch[1]) : 10
            const width = ctx.measureText(str).width as number
            const align = String(ctx.textAlign ?? 'left')
            let left = x
            if (align === 'center' || align === 'middle') left = x - width / 2
            else if (align === 'right' || align === 'end') left = x - width
            sink.push({ text: str, x: left, y, width, fontSize })
          }
          return undefined
        }
      }
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

/** Group captured spans into visual lines (same baseline, ordered by x). */
function spansToLines(spans: TextSpan[]): IndexLine[] {
  const byY = new Map<number, TextSpan[]>()
  for (const span of spans) {
    // baselines within 1px belong to the same line
    const key = Math.round(span.y)
    let bucket = byY.get(key)
    if (!bucket) {
      bucket = []
      byY.set(key, bucket)
    }
    bucket.push(span)
  }
  const lines: IndexLine[] = []
  for (const [y, bucket] of [...byY.entries()].sort((a, b) => a[0] - b[0])) {
    bucket.sort((a, b) => a.x - b.x)
    const fontSize = Math.max(...bucket.map((s) => s.fontSize), 10)
    lines.push({
      spans: bucket,
      text: bucket.map((s) => s.text).join(''),
      y,
      top: y - fontSize * 0.85,
      bottom: y + fontSize * 0.25,
    })
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
  const pages: TextIndexPage[] = paintables.map((p, index) => {
    const spans: TextSpan[] = []
    p.paint(createRecordingContext(ctx, spans))
    return { index, lines: spansToLines(spans) }
  })
  return { pages }
}

/**
 * Rectangles covering [start, end) of a line. Character offsets are mapped
 * proportionally within each span — an approximation that is exact for
 * single-span matches and close for text spanning several runs.
 */
function rectsForRange(line: IndexLine, start: number, end: number): Array<{ x: number; y: number; width: number; height: number }> {
  const height = Math.max(2, line.bottom - line.top)
  const out: Array<{ x: number; y: number; width: number; height: number }> = []
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
    out.push({ x, y: line.top, width, height })
  }
  return out
}

/** Find every occurrence of `query` across the indexed pages. */
export function findMatches(index: TextIndex, query: string, options: SearchOptions = {}): SearchMatch[] {
  const needle = options.caseSensitive ? query : query.toLowerCase()
  if (needle.length === 0) return []
  const matches: SearchMatch[] = []
  for (const page of index.pages) {
    for (let li = 0; li < page.lines.length; li++) {
      const line = page.lines[li]
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
