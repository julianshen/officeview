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
import { createRecordingContext, findMatches, rectsForRange, spansToLines, type IndexLine, type TextIndex, type TextSpan } from '../../../src/core/search'
import { hitTest, normalizeRange, rectsForSelectionOnPage, selectionSlices, textForRange, type CaretPos, type SelectionRange } from '../../../src/core/selection'
import type { HighlightRect } from '../../../src/core/overlay'
import { RECORD_TEXT, type LogicalTextRange, type LogicalTextSource } from '../../../src/core/text-recording'
import { setWatermarkMuted } from '../../../src/core/watermark'
import type { RunDetail } from '../occlusion'
import { joinSpans } from './join'

export interface UnitPoint { unit: number; x: number; y: number }                 // page px, scale 1
export type TextRange = { all: true } | { anchor: UnitPoint; focus: UnitPoint }
export type TextJob =
  | { op: 'prepare'; unit: number }
  | { op: 'search'; query: string; caseSensitive: boolean; fromUnit: number; toUnit: number; maxMatches: number }
  | { op: 'select'; range: TextRange }
  | { op: 'copy'; range: TextRange; maxParagraphs: number | null; maxBytes: number }
  | { op: 'a11y'; unit: number; paragraph?: number }
export type TextReply =
  | { op: 'prepare'; served: number; withheld: number; bytes: number }
  | { op: 'search'; matches: Array<{ unit: number; rects: number[] /* x,y,w,h … */ }> }
  | { op: 'select'; units: Array<{ unit: number; rects: number[] }>; paragraphs: number }
  | { op: 'copy'; text: string; paragraphs: number } | { op: 'copy'; refused: 'too-large'; paragraphs: number; bytes: number }
  | { op: 'a11y'; paragraphs: Array<{ rect: number[]; text?: string }> }
  | { missing: number[] }                                                         // units not prepared in this worker

export interface ServedUnit { lines: IndexLine[]; served: number; withheld: number; bytes: number }

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

/** Records one unit's text spans in paint order (the `buildTextIndex` setup:
 * a tiny context, the watermark muted, the caller's state restored).
 * `maxBytes`: a running cap on the spans' estimated size (`spansBytes`). The
 * span that would pass it is not kept and the paint is aborted by throwing
 * TextBudgetExceeded; if a painter swallows that, every later span is refused
 * too and the throw is repeated after the paint returns. */
export function recordSpans(paint: (ctx: CanvasRenderingContext2D) => void, tiny: CanvasRenderingContext2D, maxBytes = Number.POSITIVE_INFINITY): TextSpan[] {
  const spans: TextSpan[] = []
  const sources = new Set<LogicalTextSource>()
  let spent = 0, over = false
  const keep = spans.push.bind(spans)
  spans.push = (...items: TextSpan[]) => {
    for (const span of items) {
      let n = spanBytes(span)
      if (span.logical && !sources.has(span.logical.source)) n += 64 + span.logical.source.text.length * 2
      if (over || spent + n > maxBytes) { over = true; throw new TextBudgetExceeded() }
      spent += n
      if (span.logical) sources.add(span.logical.source)
      keep(span)
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
    try { paint(guardRecording(createRecordingContext(tiny, spans), preflight)) } finally { tiny.restore() }
  } finally {
    setWatermarkMuted(false)
  }
  if (over) throw new TextBudgetExceeded()
  // A plain array for every later consumer.
  return [...spans]
}

/** Intercepts the recording context's text entry points (RECORD_TEXT,
 * fillText, strokeText) so `preflight` runs before core capture. */
function guardRecording(rec: CanvasRenderingContext2D, preflight: (text: unknown, logical?: LogicalTextRange) => void): CanvasRenderingContext2D {
  const target = rec as unknown as Record<string | symbol, unknown>
  return new Proxy(target, {
    get(t, prop) {
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
    if (!logical) return { ...rest }
    // A span whose source slice is not its own text carries no logical range.
    if (logical.source.text.slice(logical.start, logical.end) !== span.text) return { ...rest }
    return { ...rest, logical }
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
        out[i] = { ...out[i]!, logical }
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
  const { served, withheld } = joinSpans(spans, runs, scale)
  const lines = spansToLines(sanitizeSpans(served))
  return { lines, served: served.length, withheld, bytes: servedBytes(lines) }
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

/** Paragraph identity (R2): a non-blank logical source, else the span itself. */
function paragraphKey(span: TextSpan): object | undefined {
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
    const found = findMatches(index, job.query, { caseSensitive: job.caseSensitive })
      .sort((a, b) => a.pageIndex - b.pageIndex || a.lineIndex - b.lineIndex || a.start - b.start)
    return { op: 'search', matches: found.slice(0, Math.max(0, job.maxMatches)).map((m) => ({ unit: m.pageIndex, rects: flat(m.rects) })) }
  }

  if (job.op === 'a11y') {
    const order: object[] = []
    const paras = new Map<object, { rects: HighlightRect[]; text: string }>()
    for (const line of units.get(job.unit)!.lines) {
      let off = 0
      for (const span of line.spans) {
        const a = off; off += span.text.length
        const key = paragraphKey(span)
        if (!key) continue
        let p = paras.get(key)
        if (!p) { p = { rects: [], text: span.logical ? span.logical.source.text : span.text }; paras.set(key, p); order.push(key) }
        p.rects.push(...rectsForRange(line, a, off))
      }
    }
    return {
      op: 'a11y',
      paragraphs: order.map((key, k) => {
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
