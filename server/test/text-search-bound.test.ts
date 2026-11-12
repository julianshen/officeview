/** Stage 3a final fix, item 2 — a search job stops at its match budget and bounds the rects it builds. */
import { describe, expect, test } from 'bun:test'
import { findMatches, spansToLines, type IndexLine, type TextSpan } from '../../src/core/search'
import type { LogicalTextSource } from '../../src/core/text-recording'
import { buildDocx, buildXlsx } from '../../src/testdata/ooxml-builders'
import { DocumentHost } from '../src/render'
import { emptyCauses } from '../src/text/join'
import { answerText, SEARCH_MAX_RECTS, type ServedUnit, type TextReply } from '../src/text/served'
import { corpus } from './harness'

type Search = Extract<TextReply, { op: 'search' }>
const unitOf = (lines: IndexLine[]): ServedUnit => ({ lines, served: lines.length, withheld: 0, bytes: 0, causes: emptyCauses() })
const search = (units: ServedUnit[], query: string, maxMatches: number, caseSensitive = true, fromUnit = 0, toUnit = units.length - 1) =>
  answerText({ op: 'search', query, caseSensitive, fromUnit, toUnit, maxMatches }, units.length, (u) => units[u]) as Search
/** The answer as it was before this fix: every match built by the library, stably sorted, then cut. */
const reference = (units: ServedUnit[], query: string, maxMatches: number, caseSensitive = true, fromUnit = 0, toUnit = units.length - 1) =>
  findMatches({ pages: units.slice(fromUnit, toUnit + 1).map((u, i) => ({ index: fromUnit + i, lines: u.lines })) }, query, { caseSensitive })
    .sort((a, b) => a.pageIndex - b.pageIndex || a.lineIndex - b.lineIndex || a.start - b.start)
    .slice(0, maxMatches).map((m) => ({ unit: m.pageIndex, rects: m.rects.flatMap((r) => [r.x, r.y, r.width, r.height]) }))

/** Counts what a search constructs while `f` runs: match-shaped objects (they carry a `start` offset — the
 * library's SearchMatch and the worker's candidates alike) and highlight rects, as they are pushed onto arrays. */
function counted<T>(f: () => T): { result: T; matches: number; rects: number } {
  const push = Array.prototype.push
  let matches = 0, rects = 0
  Array.prototype.push = function (this: unknown[], ...xs: unknown[]) {
    for (const x of xs) if (x && typeof x === 'object') {
      if ('start' in x && ('line' in x || 'lineIndex' in x)) matches++
      else if ('x' in x && 'y' in x && 'width' in x && 'height' in x && !('text' in x)) rects++
    }
    return push.apply(this, xs)
  } as typeof push
  try { return { result: f(), matches, rects } } finally { Array.prototype.push = push }
}

const plain = (text: string, y: number, x = 0): TextSpan => ({ text, x, y, width: text.length * 6, fontSize: 12 })

describe('a search job stops at its match budget (final fix, item 2)', () => {
  test('the final review probe: 100,000 matches on one line build no more than the budget and answer its first 5,001 in order', () => {
    const units = [unitOf(spansToLines([plain('a'.repeat(100_000), 20)]))]
    const want = reference(units, 'a', 5001)
    expect(want.length).toBe(5001)
    const got = counted(() => search(units, 'a', 5001))
    expect(got.result.matches).toEqual(want)
    expect(got.result.truncated).toBeUndefined()
    expect(got.matches).toBeLessThanOrEqual(5001 + 8)
    expect(got.rects).toBeLessThanOrEqual(5001 + 8)
    // In order: each match starts where the one before it ended.
    const xs = got.result.matches.map((m) => m.rects[0]!)
    expect(xs.every((x, i) => i === 0 || x > xs[i - 1]!)).toBe(true)
  })

  test('the budget is spent unit by unit: later units are not matched once it is met', () => {
    const units = [0, 1, 2].map(() => unitOf(spansToLines([plain('a'.repeat(20_000), 20), plain('a'.repeat(20_000), 40)])))
    const got = counted(() => search(units, 'a', 5001))
    expect(got.result.matches).toEqual(reference(units, 'a', 5001))
    expect(got.result.matches.every((m) => m.unit === 0)).toBe(true)
    expect(got.matches).toBeLessThanOrEqual(5001 + 8)
    // A budget that runs into the second unit takes all of the first and the head of the second.
    const more = counted(() => search(units, 'a', 45_000))
    expect(more.result.matches).toEqual(reference(units, 'a', 45_000))
    expect(more.result.matches.filter((m) => m.unit === 1).length).toBe(5000)
    expect(more.matches).toBeLessThanOrEqual(45_000 + 8)
  })

  test('one paragraph source with 100,000 matches over many lines stops at the budget too', () => {
    const text = 'a'.repeat(100_000), source: LogicalTextSource = { text, scope: {}, order: 0 }
    const spans: TextSpan[] = Array.from({ length: 100 }, (_, i) => ({ ...plain('a'.repeat(1000), 20 + 20 * i), logical: { source, start: i * 1000, end: (i + 1) * 1000 } }))
    const units = [unitOf(spansToLines(spans))]
    const got = counted(() => search(units, 'a', 5001))
    expect(got.result.matches).toEqual(reference(units, 'a', 5001))
    expect(got.matches).toBeLessThanOrEqual(5001 + 8)
    expect(got.rects).toBeLessThanOrEqual(5001 + 8)
  })
})

describe('an ordinary search is unchanged (final fix, item 2)', () => {
  const QUERIES = ['a', 'e', ' ', 'the', 'an', 'ing', 'north', 'e s', 'alpha beta', 'beta gamma delta', 'zzz', 'Ab']
  const LIMITS = [0, 1, 2, 3, 7, 50, 5001]
  const same = (units: ServedUnit[], label: string) => {
    let hits = 0
    for (const q of QUERIES) for (const cs of [true, false]) for (const max of LIMITS) {
      const got = search(units, q, max, cs), want = reference(units, q, max, cs)
      expect({ label, q, cs, max, matches: got.matches }).toEqual({ label, q, cs, max, matches: want })
      expect(got.truncated).toBeUndefined()
      hits += want.length
    }
    for (let u = 0; u < units.length; u++) expect(search(units, 'e', 5001, false, u, u).matches).toEqual(reference(units, 'e', 5001, false, u, u))
    return hits
  }

  test('a mixed page: two paragraph sources side by side, a source wrapped over lines, and source-less lines between them', () => {
    const scope = {}
    const left: LogicalTextSource = { text: 'alpha beta gamma delta and the north east area', scope, order: 0 }
    const right: LogicalTextSource = { text: 'an eastern beta gamma delta thing, the northern sea', scope, order: 1 }
    const wrap = (source: LogicalTextSource, x: number, y0: number, width: number) => {
      const out: TextSpan[] = []
      for (let at = 0, i = 0; at < source.text.length; at += width, i++) {
        const t = source.text.slice(at, at + width)
        out.push({ ...plain(t, y0 + 20 * i, x), logical: { source, start: at, end: at + t.length } })
      }
      return out
    }
    const spans = [...wrap(left, 0, 20, 13), ...wrap(right, 300, 20, 9), plain('the legacy line between: alpha beta', 30), plain('Ab ab AB aB the end', 70), plain('another the', 70, 300)]
    const units = [unitOf(spansToLines(spans)), unitOf(spansToLines([plain('second unit: the alpha beta gamma delta', 20), ...wrap(left, 0, 60, 7)]))]
    expect(same(units, 'mixed')).toBeGreaterThan(200)
  })

  test('lines of one source out of source order (hand-built index): still the old answer', () => {
    // The first match's line decides the order; here the source's later text sits on the earlier line.
    const source: LogicalTextSource = { text: 'the east and the west and the north' }
    const part = (start: number, end: number, y: number): TextSpan => ({ ...plain(source.text.slice(start, end), y), logical: { source, start, end } })
    const line = (spans: TextSpan[], y: number): IndexLine => ({ spans, text: spans.map((s) => s.text).join(''), y, top: y - 10, bottom: y + 3 })
    const lines = [line([part(24, 35, 20)], 20), line([plain('the the legacy the', 40)], 40), line([part(12, 24, 60), part(0, 12, 60)], 60)]
    expect(same([unitOf(lines)], 'unordered')).toBeGreaterThan(30)
  })

  test('real documents through the worker: DOCX paragraphs and a table, XLSX cells, a PPTX deck', async () => {
    const cell = (text: string) => ({ paragraphs: [{ runs: [{ text }] }] })
    const docs: Array<[string, Uint8Array]> = [
      ['docx', await buildDocx([{ runs: [{ text: 'The north wind and the sun were disputing which was the stronger, ' }, { text: 'when a traveller came along wrapped in a warm cloak. They agreed that the one who first succeeded in making the traveller take his cloak off should be considered stronger than the other.' }] }, { runs: [{ text: 'Then the north wind blew as hard as he could.' }] }],
        [{ gridCols: ['2500', '2500'], rows: [{ cells: [cell('north west and the east'), cell('the north east area')] }, { cells: [cell('south west then'), cell('the south east')] }] }])],
      ['xlsx', await buildXlsx([{ name: 'S', cols: '<col min="1" max="3" width="24" customWidth="1"/>', rows: [1, 2, 3].map((r) => ({ r, cells: ['A', 'B', 'C'].map((c, i) => ({ ref: `${c}${r}`, t: 's', v: (r + i) % 4 })) })) }], ['the north area', 'an eastern thing', 'alpha beta', 'the sea'])],
      ['pptx', corpus('pypptx-test.pptx')],
    ]
    for (const [label, bytes] of docs) {
      const host = new DocumentHost()
      const o = await host.open(`bound-${label}`, bytes, 'visible-only')
      const units = o.units.map((_, u) => { host.text(o.docKey, { op: 'prepare', unit: u }); return host.servedUnit(o.docKey, u)! })
      expect({ label, hits: same(units, label) > 20 }).toEqual({ label, hits: true })
      // And through the host's own entry point.
      expect((host.text(o.docKey, { op: 'search', query: 'e', caseSensitive: false, fromUnit: 0, toUnit: units.length - 1, maxMatches: 5001 }) as Search).matches).toEqual(reference(units, 'e', 5001, false))
      host.drop(o.docKey)
    }
  }, 120_000)
})

describe('a search job bounds the rects it builds (final fix, item 2)', () => {
  // 100 lines of 2,000 one-character spans; a 200-character query matches ten times a line, 200 rects a match.
  const lines = Array.from({ length: 100 }, (_, i) => spansToLines(Array.from({ length: 2000 }, (_, k) => plain('a', 20 + 20 * i, k * 6)))[0]!)
  const units = [unitOf(lines)]
  const query = 'a'.repeat(200)

  test(`a job stops at ${SEARCH_MAX_RECTS} rects with the matches it has completed, flagged truncated`, () => {
    const want = reference(units, query, 5001)
    expect(want.length).toBe(1000)
    expect(want[0]!.rects.length).toBe(800)
    const got = counted(() => search(units, query, 5001))
    expect(got.result.truncated).toBe(true)
    expect(got.result.matches.length).toBe(SEARCH_MAX_RECTS / 200)
    expect(got.result.matches).toEqual(want.slice(0, SEARCH_MAX_RECTS / 200))
    expect(got.rects).toBeLessThanOrEqual(SEARCH_MAX_RECTS + 2000)
  })

  test('a job asking for more matches than the bound gets the bound, flagged truncated', () => {
    const many = [unitOf(spansToLines([plain('a'.repeat(SEARCH_MAX_RECTS + 50), 20)]))]
    const got = counted(() => search(many, 'a', 1e9))
    expect(got.result.matches.length).toBe(SEARCH_MAX_RECTS)
    expect(got.result.truncated).toBe(true)
    expect(got.matches).toBeLessThanOrEqual(SEARCH_MAX_RECTS + 8)
    for (const bad of [NaN, -1, 0]) expect(search(many, 'a', bad)).toEqual({ op: 'search', matches: [] })
  })

  test('a job under the bound is whole and not flagged', () => {
    const got = search(units, query, 400)
    expect(got.matches).toEqual(reference(units, query, 400))
    expect('truncated' in got).toBe(false)
  })
})
