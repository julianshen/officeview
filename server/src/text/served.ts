/**
 * Stage 3a — the served text index of one unit, and the worker-side text
 * operations answered from it (search, select, copy, a11y).
 *
 * A unit is recorded (spans in paint order), joined against the occlusion
 * run details of a visual pass (`joinSpans`, R1), then SANITISED: a served
 * span's `logical.source` holds the whole paragraph, so it is replaced by a
 * per-unit segment source built only from served ranges (merged across
 * whitespace-only gaps). Invariant: no character outside a served span — bar
 * whitespace inside a segment — is reachable from a served unit. Copy may
 * therefore break a paragraph at a withheld gap or a page break.
 */
import { createRecordingContext, rectsForRange, spansToLines, type IndexLine, type TextIndex, type TextSpan } from '../../../src/core/search'
import { hitTest, normalizeRange, rectsForSelectionOnPage, selectionSlices, textForRange, type CaretPos, type SelectionRange } from '../../../src/core/selection'
import type { HighlightRect } from '../../../src/core/overlay'
import { RECORD_GROUP, RECORD_TEXT, type LogicalTextRange, type LogicalTextSource } from '../../../src/core/text-recording'
import { setWatermarkMuted } from '../../../src/core/watermark'
import type { RunDetail } from '../occlusion'
import { emptyCauses, joinSpans, type WithheldCauses } from './join'

export interface UnitPoint { unit: number; x: number; y: number }                 // page px, scale 1
export type TextRange = { all: true } | { anchor: UnitPoint; focus: UnitPoint }
export type TextJob =
  | { op: 'prepare'; unit: number }
  | { op: 'search'; query: string; caseSensitive: boolean; fromUnit: number; toUnit: number; maxMatches: number }
  | { op: 'select'; range: TextRange }
  | { op: 'copy'; range: TextRange; maxParagraphs: number | null; maxBytes: number }
  | { op: 'a11y'; unit: number; paragraph?: number }
export type TextReply =
  | { op: 'prepare'; served: number; withheld: number; bytes: number; /** Diagnostic counts only (yield script); never sent on the wire. */ causes: WithheldCauses }
  | { op: 'search'; matches: Array<{ unit: number; rects: number[] /* x,y,w,h … */ }>; /** The job stopped at SEARCH_MAX_RECTS: `matches` is a prefix of the answer. */ truncated?: true }
  | { op: 'select'; units: Array<{ unit: number; rects: number[] }>; paragraphs: number }
  | { op: 'copy'; text: string; paragraphs: number } | { op: 'copy'; refused: 'too-large'; paragraphs: number; bytes: number }
  | { op: 'a11y'; paragraphs: Array<{ rect: number[]; text?: string }> }
  | { missing: number[] }                                                         // units not prepared in this worker

export interface ServedUnit { lines: IndexLine[]; served: number; withheld: number; bytes: number; causes: WithheldCauses }

/** R6 work bound: one search job builds at most this many highlight rects
 * (and returns at most this many matches). A job that would pass it stops
 * with the matches it has completed and answers `truncated: true`. */
export const SEARCH_MAX_RECTS = 100_000
/** R6: served-index cache bound per document per worker. */
export const SERVED_CACHE_BYTES = 32 * 1024 * 1024
/** Fixed allowance for one prepare's text structures, split into the two
 * running caps below: 3 × TEXT_RECORD_BUDGET + 2 × OBSERVER_BUDGET = 15 MiB. */
export const TEXT_PREPARE_ALLOWANCE = 16 * 1024 * 1024
/** Running estimate cap on recorded spans (≤ 3 copies live: recorded, sanitised, lines). */
export const TEXT_RECORD_BUDGET = 3 * 1024 * 1024
/** Running estimate cap on the prepare visual pass's observer records (≤ 2 copies: records + run details). */
export const OBSERVER_BUDGET = 3 * 1024 * 1024
/** Workspace one prepare needs: the scale-1 RGBA visual canvas plus the allowance. */
export const prepareWorkspace = (w: number, h: number) => 4 * w * h + TEXT_PREPARE_ALLOWANCE

/** Thrown once recorded spans would pass the recording budget. */
export class TextBudgetExceeded extends Error { constructor() { super('recorded text exceeds the prepare allowance'); this.name = 'TextBudgetExceeded' } }

/** Paragraph/cell group of a recorded span (RECORD_GROUP), kept off the
 * library's TextSpan. The value is an opaque per-recording token, never the
 * painter's key object, so a served unit retains no layout or cell model.
 * `sanitizeSpans` carries the tag onto the spans it copies. */
const groups = new WeakMap<TextSpan, object>()
/** A copy of `from` that keeps its group tag. */
function carry(from: TextSpan, to: TextSpan): TextSpan {
  const g = groups.get(from)
  if (g) groups.set(to, g)
  return to
}

/** Records one unit's text spans in paint order (the `buildTextIndex` setup:
 * a tiny context, the watermark muted, the caller's state restored).
 * The recording context defines RECORD_GROUP: each kept span is tagged with
 * the painter's current group (a token per distinct painter key).
 * `maxBytes`: a running cap on the spans' estimated size (`spansBytes`). The
 * span that would pass it is not kept and the paint is aborted by throwing
 * TextBudgetExceeded; if a painter swallows that, every later span is refused
 * too and the throw is repeated after the paint returns. */
export function recordSpans(paint: (ctx: CanvasRenderingContext2D) => void, tiny: CanvasRenderingContext2D, maxBytes = Number.POSITIVE_INFINITY): TextSpan[] {
  const spans: TextSpan[] = []
  const sources = new Set<LogicalTextSource>()
  let spent = 0, over = false
  // Painter key → token, minted only when a span is kept under it.
  const tokens = new Map<object, object>()
  let current: object | undefined
  const setGroup = (key: unknown) => { current = key !== null && (typeof key === 'object' || typeof key === 'function') ? key : undefined }
  const keep = spans.push.bind(spans)
  spans.push = (...items: TextSpan[]) => {
    for (const span of items) {
      let n = spanBytes(span)
      if (span.logical && !sources.has(span.logical.source)) n += 64 + span.logical.source.text.length * 2
      if (over || spent + n > maxBytes) { over = true; throw new TextBudgetExceeded() }
      spent += n
      if (span.logical) sources.add(span.logical.source)
      keep(span)
      if (current) {
        let token = tokens.get(current)
        if (!token) { token = {}; tokens.set(current, token) }
        groups.set(span, token)
      }
    }
    return spans.length
  }
  /** Preflight, before core capture measures, maps or serialises anything:
   * the record's cost from its inputs alone (array lengths are read, never
   * copied), against the same running total the sink charges. */
  const preflight = (text: unknown, logical?: LogicalTextRange) => {
    let n = 256 + String(text).length * 2
    if (logical) {
      n += 8 * (logical.graphemeBoundaries?.length ?? 0)
      if (logical.source && !sources.has(logical.source)) n += 64 + String(logical.source.text).length * 2
      if (logical.clip) n += 16 * 4
      const v = logical.visual
      if (v) n += 16 * ((v.polygon?.length ?? 0) + (v.clusters?.length ?? 0) + (v.bands?.length ?? 0) * 2 + (v.cells?.length ?? 0) * 4 + (v.tris?.length ?? 0) * 3)
    }
    if (over || spent + n > maxBytes) { over = true; throw new TextBudgetExceeded() }
  }
  setWatermarkMuted(true)
  try {
    tiny.save()
    try { paint(guardRecording(createRecordingContext(tiny, spans), preflight, setGroup)) } finally { tiny.restore() }
  } finally {
    setWatermarkMuted(false)
  }
  if (over) throw new TextBudgetExceeded()
  // A plain array for every later consumer.
  return [...spans]
}

/** Intercepts the recording context's text entry points (RECORD_TEXT,
 * fillText, strokeText) so `preflight` runs before core capture, and defines
 * RECORD_GROUP (which the library's recording context does not). */
function guardRecording(rec: CanvasRenderingContext2D, preflight: (text: unknown, logical?: LogicalTextRange) => void, setGroup: (key: unknown) => void): CanvasRenderingContext2D {
  const target = rec as unknown as Record<string | symbol, unknown>
  return new Proxy(target, {
    get(t, prop) {
      if (prop === RECORD_GROUP) return setGroup
      if (prop === RECORD_TEXT) {
        const record = t[RECORD_TEXT] as ((text: string, x: number, y: number, width: number, logical: LogicalTextRange) => void) | undefined
        return record && ((text: string, x: number, y: number, width: number, logical: LogicalTextRange) => { preflight(text, logical); record(text, x, y, width, logical) })
      }
      if (prop === 'fillText' || prop === 'strokeText') {
        const draw = t[prop] as (text: string, x: number, y: number, maxWidth?: number) => void
        return (text: string, x: number, y: number, maxWidth?: number) => { preflight(text); maxWidth === undefined ? draw(text, x, y) : draw(text, x, y, maxWidth) }
      }
      return t[prop]
    },
    set(t, prop, value) { t[prop] = value; return true },
  }) as unknown as CanvasRenderingContext2D
}

const isBlank = (s: string) => s.trim() === ''

/** Replaces each served span's paragraph source by a per-unit segment source
 * holding only served text (see the module comment). Inputs are not mutated. */
export function sanitizeSpans(served: TextSpan[]): TextSpan[] {
  const out: TextSpan[] = served.map((span) => {
    const { logical, ...rest } = span
    if (!logical) return carry(span, { ...rest })
    // A span whose source slice is not its own text carries no logical range.
    if (logical.source.text.slice(logical.start, logical.end) !== span.text) return carry(span, { ...rest })
    return carry(span, { ...rest, logical })
  })
  const bySource = new Map<LogicalTextSource, number[]>()
  out.forEach((span, i) => {
    if (!span.logical) return
    const list = bySource.get(span.logical.source) ?? []
    list.push(i); bySource.set(span.logical.source, list)
  })
  for (const [source, members] of bySource) {
    members.sort((a, b) => out[a]!.logical!.start - out[b]!.logical!.start || out[a]!.logical!.end - out[b]!.logical!.end)
    const segments: Array<{ start: number; end: number; members: number[] }> = []
    for (const i of members) {
      const { start, end } = out[i]!.logical!
      const last = segments[segments.length - 1]
      if (last && (start <= last.end || isBlank(source.text.slice(last.end, start)))) {
        last.end = Math.max(last.end, end); last.members.push(i)
      } else segments.push({ start, end, members: [i] })
    }
    segments.forEach((seg, k) => {
      // Fractional order keeps a paragraph's segments in source order among
      // its scope (spansToLines sorts a scope's lines by order, then start).
      const fresh: LogicalTextSource = {
        text: source.text.slice(seg.start, seg.end),
        ...(source.scope ? { scope: source.scope } : {}),
        order: (source.order ?? 0) + k / (segments.length + 1),
      }
      for (const i of seg.members) {
        const old = out[i]!.logical!
        const logical: LogicalTextRange = { ...old, source: fresh, start: old.start - seg.start, end: old.end - seg.start }
        delete logical.graphemeBoundaries
        const gb = old.graphemeBoundaries
        // Boundaries are absolute source offsets within the span; any outside
        // its own range is dropped with the whole list (fail closed).
        if (gb && gb.every((b) => b >= old.start && b <= old.end)) logical.graphemeBoundaries = gb.map((b) => b - seg.start)
        out[i] = carry(out[i]!, { ...out[i]!, logical })
      }
    })
  }
  return out
}

/** Rough retained size of one span (its source text not included). */
function spanBytes(s: TextSpan): number {
  let n = 256 + s.text.length * 2
  const v = s.placement?.visual
  if (v) n += 16 * (v.polygon.length + (v.clusters?.length ?? 0) + (v.bands?.length ?? 0) * 2 + (v.cells?.length ?? 0) * 4 + (v.tris?.length ?? 0) * 3)
  if (s.placement?.clip) n += 16 * s.placement.clip.length
  if (s.logical) n += 8 * (s.logical.graphemeBoundaries?.length ?? 0)
  return n
}

/** Rough size of recorded spans, sources counted once. */
export function spansBytes(spans: TextSpan[]): number {
  const sources = new Set<LogicalTextSource>()
  let n = 0
  for (const s of spans) { n += spanBytes(s); if (s.logical) sources.add(s.logical.source) }
  for (const s of sources) n += 64 + s.text.length * 2
  return n
}

/** Rough retained size of a served unit, for the R6 cache bound. */
export function servedBytes(lines: IndexLine[]): number {
  return lines.reduce((n, line) => n + 64 + line.text.length * 2, 0) + spansBytes(lines.flatMap((l) => l.spans))
}

/** Join + sanitise + line building for one unit. */
export function buildServedUnit(spans: TextSpan[], runs: RunDetail[], scale = 1): ServedUnit {
  const causes = emptyCauses()
  const { served, withheld } = joinSpans(spans, runs, scale, causes)
  const lines = spansToLines(sanitizeSpans(served))
  return { lines, served: served.length, withheld, bytes: servedBytes(lines), causes }
}

/** Per-document LRU of served units (R6 bound per document). */
export class ServedCache {
  private docs = new Map<string, { units: Map<number, ServedUnit>; bytes: number }>()
  constructor(public limit = SERVED_CACHE_BYTES) {}
  get(docKey: string, unit: number): ServedUnit | undefined {
    const d = this.docs.get(docKey), u = d?.units.get(unit)
    if (d && u) { d.units.delete(unit); d.units.set(unit, u) }
    return u
  }
  /** Peeks without touching the LRU order. */
  peek(docKey: string, unit: number): ServedUnit | undefined { return this.docs.get(docKey)?.units.get(unit) }
  put(docKey: string, unit: number, served: ServedUnit): void {
    let d = this.docs.get(docKey)
    if (!d) { d = { units: new Map(), bytes: 0 }; this.docs.set(docKey, d) }
    const old = d.units.get(unit)
    if (old) { d.units.delete(unit); d.bytes -= old.bytes }
    if (served.bytes > this.limit) return
    for (const [k, v] of d.units) {
      if (d.bytes + served.bytes <= this.limit) break
      d.units.delete(k); d.bytes -= v.bytes
    }
    d.units.set(unit, served); d.bytes += served.bytes
  }
  drop(docKey: string): void { this.docs.delete(docKey) }
  bytesOf(docKey: string): number { return this.docs.get(docKey)?.bytes ?? 0 }
}

const flat = (rects: HighlightRect[]) => rects.flatMap((r) => [r.x, r.y, r.width, r.height])
const range = (a: number, b: number) => Array.from({ length: b - a + 1 }, (_, i) => a + i)

/** Paragraph identity (R2, Task 3b): the span's group (paragraph or cell),
 * else a non-blank logical source, else the span itself. Blank text has none. */
function paragraphKey(span: TextSpan): object | undefined {
  const group = groups.get(span)
  if (group) return isBlank(span.text) ? undefined : group
  if (span.logical) return isBlank(span.logical.source.text) ? undefined : span.logical.source
  return isBlank(span.text) ? undefined : span
}

function paragraphsIn(index: TextIndex, sel: SelectionRange): number {
  const keys = new Set<object>()
  for (const s of selectionSlices(index, sel)) {
    const line = index.pages.find((p) => p.index === s.pageIndex)?.lines[s.lineIndex]
    if (!line) continue
    let off = 0
    for (const span of line.spans) {
      const a = off, b = off + span.text.length; off = b
      const lo = Math.max(a, s.from), hi = Math.min(b, s.to)
      if (hi <= lo || isBlank(span.text.slice(lo - a, hi - a))) continue
      const key = paragraphKey(span)
      if (key) keys.add(key)
    }
  }
  return keys.size
}

const endOf = (lines: IndexLine[], unit: number): CaretPos => ({ pageIndex: unit, lineIndex: Math.max(0, lines.length - 1), charIndex: lines[lines.length - 1]?.text.length ?? 0 })
const startOf = (unit: number): CaretPos => ({ pageIndex: unit, lineIndex: 0, charIndex: 0 })

// ── Bounded search (final fix, item 2) ─────────────────────────────────────
// The library's findMatches builds every match of the index, rects included,
// before anything can cut the answer. The worker matches the same text the
// same way but keeps only the first `limit` matches of a unit, in line →
// position order, and builds rects for those alone.

/** One span of a paragraph source on a line, as findMatches collects them. */
interface SourceRange { span: TextSpan; lineIndex: number; offset: number }
/** A match not yet given rects. `line`/`start`: its first line and offset
 * there (the order key); `item`/`at`: findMatches' emission order (ties).
 * `ranges`/`first`: a source match's spans and the first one it covers. */
interface Candidate { line: number; start: number; item: number; at: number; ranges?: SourceRange[]; first?: number }
const candidateOrder = (a: Candidate, b: Candidate) => a.line - b.line || a.start - b.start || a.item - b.item || a.at - b.at

/**
 * The first `limit` matches of `needle` on one unit's lines, ordered by
 * line, then position, then findMatches' own order — the order a stable
 * sort of findMatches' result by (line, start) gives. At most 2 × `limit`
 * candidates are held at once; nothing unbounded is built or sorted. Matching
 * stops as soon as no later match can come before the `limit`-th kept: at
 * the end of a source-less line, in a paragraph source whose spans lie in
 * source order (every painter's), and before every item that starts on a
 * later line. A source whose spans are out of order is scanned to its end
 * (positions only: a candidate past the cut is not built).
 */
function unitMatches(lines: IndexLine[], needle: string, caseSensitive: boolean, limit: number): Candidate[] {
  if (!(limit > 0)) return []
  const sources = new Map<LogicalTextSource, SourceRange[]>()
  lines.forEach((line, lineIndex) => {
    let offset = 0
    for (const span of line.spans) {
      if (span.logical) {
        const ranges = sources.get(span.logical.source) ?? []
        ranges.push({ span, lineIndex, offset }); sources.set(span.logical.source, ranges)
      }
      offset += span.text.length
    }
  })
  // Items in findMatches' emission order: paragraph sources, then source-less lines.
  const items: Array<{ minLine: number; text: string; ranges?: SourceRange[]; ordered: boolean }> = []
  for (const [source, ranges] of sources) {
    ranges.sort((a, b) => a.span.logical!.start - b.span.logical!.start)
    let minLine = Infinity, ordered = true
    ranges.forEach((r, i) => {
      minLine = Math.min(minLine, r.lineIndex)
      const p = ranges[i - 1]
      if (p && !(r.lineIndex > p.lineIndex || (r.lineIndex === p.lineIndex && r.offset >= p.offset + (p.span.logical!.end - p.span.logical!.start)))) ordered = false
    })
    items.push({ minLine, text: source.text, ranges, ordered })
  }
  lines.forEach((line, li) => { if (!line.spans.some((span) => span.logical)) items.push({ minLine: li, text: line.text, ordered: true }) })
  const byLine = items.map((_, i) => i).sort((a, b) => items[a]!.minLine - items[b]!.minLine || a - b)

  const kept: Candidate[] = []
  /** The `limit`-th kept candidate once there are that many: nothing at or after it is wanted. */
  let cut: Candidate | undefined
  const settle = () => { kept.sort(candidateOrder); if (kept.length > limit) kept.length = limit; cut = kept.length === limit ? kept[limit - 1] : undefined }
  /** Keeps the candidate unless it is at or after the cut (false). */
  const offer = (line: number, start: number, item: number, at: number, ranges?: SourceRange[], first?: number): boolean => {
    if (cut && (line - cut.line || start - cut.start || item - cut.item || at - cut.at) >= 0) return false
    kept.push(ranges ? { line, start, item, at, ranges, first } : { line, start, item, at })
    if ((!cut && kept.length === limit) || kept.length >= 2 * limit) settle()
    return true
  }
  for (const i of byLine) {
    const it = items[i]!
    if (cut && cut.line < it.minLine) break
    const haystack = caseSensitive ? it.text : it.text.toLowerCase()
    const ranges = it.ranges
    let from = 0, lo = 0
    for (;;) {
      const at = haystack.indexOf(needle, from)
      if (at < 0) break
      const end = at + needle.length
      from = end
      if (!ranges) { if (!offer(it.minLine, at, i, at)) break; continue }
      // The first span the match covers (logical end > at and start < end), as findMatches' filter finds it.
      while (lo < ranges.length && ranges[lo]!.span.logical!.end <= at) lo++
      let first = -1
      for (let j = lo; j < ranges.length && ranges[j]!.span.logical!.start < end; j++) if (ranges[j]!.span.logical!.end > at) { first = j; break }
      if (first < 0) continue
      const f = ranges[first]!
      if (!offer(f.lineIndex, f.offset + Math.max(0, at - f.span.logical!.start), i, at, ranges, first) && it.ordered) break
    }
  }
  settle()
  return kept
}

/** A candidate's rects, as findMatches builds them; undefined once they would take the job past `budget` rects. */
function candidateRects(c: Candidate, lines: IndexLine[], length: number, budget: number): HighlightRect[] | undefined {
  if (!c.ranges) { const r = rectsForRange(lines[c.line]!, c.at, c.at + length); return r.length <= budget ? r : undefined }
  const end = c.at + length
  let rects: HighlightRect[] | undefined
  for (let j = c.first!; j < c.ranges.length && c.ranges[j]!.span.logical!.start < end; j++) {
    const r = c.ranges[j]!, logical = r.span.logical!
    if (!(logical.end > c.at)) continue
    const part = rectsForRange(lines[r.lineIndex]!, r.offset + Math.max(0, c.at - logical.start), r.offset + Math.min(r.span.text.length, end - logical.start))
    if (!rects) rects = part
    else for (const rect of part) rects.push(rect)
    if (rects.length > budget) return undefined
  }
  return rects ?? []
}

export class TextRefusal extends Error { constructor(readonly code: 'unit-range' | 'bad-job') { super(code); this.name = 'TextRefusal' } }

/**
 * Answers one text job from served units. `get(u)` returns the cached unit or
 * undefined; any needed unit that is missing yields `{ missing }`.
 * (`prepare` is the host's job, not answered here.)
 */
export function answerText(job: Exclude<TextJob, { op: 'prepare' }>, unitCount: number, get: (unit: number) => ServedUnit | undefined): TextReply {
  const valid = (u: unknown): u is number => Number.isInteger(u) && (u as number) >= 0 && (u as number) < unitCount
  let needed: number[]
  if (job.op === 'search') {
    if (!valid(job.fromUnit) || !valid(job.toUnit) || job.fromUnit > job.toUnit) throw new TextRefusal('unit-range')
    needed = range(job.fromUnit, job.toUnit)
  } else if (job.op === 'a11y') {
    if (!valid(job.unit)) throw new TextRefusal('unit-range')
    needed = [job.unit]
  } else {
    const r = job.range
    if ('all' in r) { if (r.all !== true) throw new TextRefusal('bad-job'); needed = unitCount ? range(0, unitCount - 1) : [] }
    else {
      if (!valid(r.anchor?.unit) || !valid(r.focus?.unit)) throw new TextRefusal('unit-range')
      needed = range(Math.min(r.anchor.unit, r.focus.unit), Math.max(r.anchor.unit, r.focus.unit))
    }
  }
  const units = new Map<number, ServedUnit>()
  const missing: number[] = []
  for (const u of needed) { const s = get(u); if (s) units.set(u, s); else missing.push(u) }
  if (missing.length) return { missing }
  const index: TextIndex = { pages: needed.map((u) => ({ index: u, lines: units.get(u)!.lines })) }

  if (job.op === 'search') {
    // Unit → line → position order; the budget is spent unit by unit and no
    // unit is matched once it is met. The caller asks for one match more than
    // it will send, so it sees whether the answer went on.
    const needle = job.caseSensitive ? job.query : job.query.toLowerCase()
    const limit = Math.min(Math.floor(Math.max(0, job.maxMatches)) || 0, SEARCH_MAX_RECTS)
    const matches: Array<{ unit: number; rects: number[] }> = []
    let built = 0
    if (needle.length > 0) for (const u of needed) {
      if (matches.length >= limit) break
      const lines = units.get(u)!.lines
      for (const c of unitMatches(lines, needle, job.caseSensitive, limit - matches.length)) {
        const rects = candidateRects(c, lines, needle.length, SEARCH_MAX_RECTS - built)
        if (!rects) return { op: 'search', matches, truncated: true }
        built += rects.length
        matches.push({ unit: u, rects: flat(rects) })
      }
    }
    // A job that asked for more matches than the bound allows cannot tell whether the answer went on.
    return matches.length >= limit && job.maxMatches > SEARCH_MAX_RECTS ? { op: 'search', matches, truncated: true } : { op: 'search', matches }
  }

  if (job.op === 'a11y') {
    // A group (paragraph or cell) is one paragraph: the text textForRange
    // gives for its served spans, line by line (a source-backed group slices
    // its source; a source-less one joins lines with newlines). Its blank
    // spans join it, but a group with no ink is no paragraph.
    const order: object[] = []
    const paras = new Map<object, { rects: HighlightRect[]; text: string; ink: boolean; pieces?: Array<{ line: IndexLine; spans: TextSpan[] }> }>()
    for (const line of units.get(job.unit)!.lines) {
      let off = 0
      for (const span of line.spans) {
        const a = off; off += span.text.length
        const group = groups.get(span)
        const key = group ?? paragraphKey(span)
        if (!key) continue
        let p = paras.get(key)
        if (!p) {
          p = group ? { rects: [], text: '', ink: false, pieces: [] } : { rects: [], text: span.logical ? span.logical.source.text : span.text, ink: true }
          paras.set(key, p); order.push(key)
        }
        if (p.pieces) {
          const last = p.pieces[p.pieces.length - 1]
          if (last?.line === line) last.spans.push(span)
          else p.pieces.push({ line, spans: [span] })
          if (!isBlank(span.text)) p.ink = true
        }
        p.rects.push(...rectsForRange(line, a, off))
      }
    }
    for (const p of paras.values()) {
      if (!p.pieces?.length) continue
      const lines = p.pieces.map(({ line, spans }) => ({ ...line, spans, text: spans.map((s) => s.text).join('') }))
      const last = lines[lines.length - 1]!
      p.text = textForRange({ pages: [{ index: 0, lines }] }, { start: startOf(0), end: { pageIndex: 0, lineIndex: lines.length - 1, charIndex: last.text.length } })
    }
    return {
      op: 'a11y',
      paragraphs: order.filter((key) => paras.get(key)!.ink).map((key, k) => {
        const p = paras.get(key)!
        const x0 = Math.min(...p.rects.map((r) => r.x)), y0 = Math.min(...p.rects.map((r) => r.y))
        const x1 = Math.max(...p.rects.map((r) => r.x + r.width)), y1 = Math.max(...p.rects.map((r) => r.y + r.height))
        const rect = p.rects.length ? [x0, y0, x1 - x0, y1 - y0] : []
        return job.paragraph === undefined || job.paragraph === k ? { rect, text: p.text } : { rect }
      }),
    }
  }

  // select / copy
  const r = job.range
  let sel: SelectionRange
  if ('all' in r) {
    const last = needed[needed.length - 1]
    sel = last === undefined ? { start: startOf(0), end: startOf(0) } : { start: startOf(needed[0]!), end: endOf(units.get(last)!.lines, last) }
  } else {
    const before = r.anchor.unit !== r.focus.unit ? r.anchor.unit < r.focus.unit : r.anchor.y < r.focus.y || (r.anchor.y === r.focus.y && r.anchor.x <= r.focus.x)
    const resolve = (p: UnitPoint, isStart: boolean): CaretPos =>
      hitTest(index, p.unit, p.x, p.y, { strict: false }) ?? (isStart ? startOf(p.unit) : endOf(units.get(p.unit)!.lines, p.unit))
    sel = normalizeRange(resolve(r.anchor, before), resolve(r.focus, !before))
  }
  const paragraphs = paragraphsIn(index, sel)
  if (job.op === 'select') {
    return { op: 'select', paragraphs, units: needed.flatMap((u) => { const rects = rectsForSelectionOnPage(index, u, sel); return rects.length ? [{ unit: u, rects: flat(rects) }] : [] }) }
  }
  if (paragraphs === 0) return { op: 'copy', text: '', paragraphs: 0 }
  // Caps fail closed: a malformed cap (NaN, undefined) refuses.
  if (job.maxParagraphs !== null && !(paragraphs <= job.maxParagraphs)) return { op: 'copy', refused: 'too-large', paragraphs, bytes: 0 }
  const text = textForRange(index, sel)
  const bytes = Buffer.byteLength(text, 'utf8')
  if (!(bytes <= job.maxBytes)) return { op: 'copy', refused: 'too-large', paragraphs, bytes }
  return { op: 'copy', text, paragraphs }
}
