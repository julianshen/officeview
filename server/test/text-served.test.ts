/** Stage 3a Task 3 — served text index in the render worker; `text` jobs. */
import { afterAll, describe, expect, test } from 'bun:test'
import JSZip from 'jszip'
import { buildTextIndex, type TextSpan } from '../../src/core/search'
import { textForRange } from '../../src/core/selection'
import type { LogicalTextSource } from '../../src/core/text-recording'
import { occlusionObserver, type RunDetail } from '../src/occlusion'
import { joinSpans } from '../src/text/join'
import { Pool, type PoolTextJob } from '../src/pool'
import { DocumentHost, FIXED_MARK_TEXT } from '../src/render'
import { answerText, buildServedUnit, recordSpans, type ServedUnit, type TextReply } from '../src/text/served'
import { augmentedFixtures } from './fixtures'
import { corpus, disposePool, pool } from './harness'

afterAll(() => disposePool())
const pools: Pool[] = []
afterAll(() => { for (const p of pools) p.dispose() })

const deck = corpus('pypptx-test.pptx')
let refs = 0
async function opened(bytes: Uint8Array, policy: 'visible-only' | 'full' = 'visible-only', p = pool()) {
  const docRef = `doc-${++refs}`
  const open = await p.open(docRef, bytes, policy)
  const text = (job: Omit<PoolTextJob, 'docRef' | 'contentPolicy'> & Record<string, unknown>) => p.text('s1', { docRef, contentPolicy: policy, ...job } as PoolTextJob)
  return { docRef, open, text }
}
const copyAll = { op: 'copy', range: { all: true }, maxParagraphs: null, maxBytes: 4 << 20 } as const
const units = (n: number) => Array.from({ length: n }, (_, i) => i)
const rectsOf = (flat: number[]) => Array.from({ length: flat.length / 4 }, (_, i) => flat.slice(i * 4, i * 4 + 4))

describe('served text via the pool', () => {
  test('search finds visible text and answers page-px rects inside the unit', async () => {
    const d = await opened(deck)
    const r = await d.text({ op: 'search', query: 'subtitle', caseSensitive: false, fromUnit: 0, toUnit: 0, maxMatches: 100 }) as Extract<TextReply, { op: 'search' }>
    expect(r.op).toBe('search')
    expect(r.matches.length).toBe(1)
    const { w, h } = d.open.units[0]!
    for (const m of r.matches) {
      expect(m.unit).toBe(0)
      expect(m.rects.length % 4).toBe(0)
      for (const [x, y, rw, rh] of rectsOf(m.rects)) {
        expect(rw!).toBeGreaterThan(0); expect(rh!).toBeGreaterThan(0)
        expect(x!).toBeGreaterThanOrEqual(0); expect(y!).toBeGreaterThanOrEqual(0)
        expect(x! + rw!).toBeLessThanOrEqual(w); expect(y! + rh!).toBeLessThanOrEqual(h)
      }
    }
    const none = await d.text({ op: 'search', query: 'subtitle', caseSensitive: true, fromUnit: 0, toUnit: 0, maxMatches: 100 }) as Extract<TextReply, { op: 'search' }>
    expect(none.matches).toEqual([])
  }, 60_000)

  test('select between two points answers rects and a paragraph count ≥ 1', async () => {
    const d = await opened(deck)
    const { w, h } = d.open.units[0]!
    const r = await d.text({ op: 'select', range: { anchor: { unit: 0, x: 0, y: 0 }, focus: { unit: 0, x: w, y: h } } }) as Extract<TextReply, { op: 'select' }>
    expect(r.op).toBe('select')
    expect(r.paragraphs).toBeGreaterThanOrEqual(1)
    expect(r.units.length).toBe(1)
    expect(r.units[0]!.unit).toBe(0)
    expect(r.units[0]!.rects.length).toBeGreaterThanOrEqual(4)
  }, 60_000)

  test('copy answers the same string as local textForRange for the same range', async () => {
    // pypptx-test.pptx: both spans of its one slide are served (2/0), so the
    // whole-unit range is a range where every span is served.
    const d = await opened(deck)
    const r = await d.text(copyAll) as Extract<TextReply, { op: 'copy' }>
    const host = new DocumentHost()
    const o = await host.open('local', deck, 'visible-only')
    const paintables = (host as unknown as { docs: Map<string, { paintables: Parameters<typeof buildTextIndex>[0] }> }).docs.get(o.docKey)!.paintables
    const local = await buildTextIndex([...paintables])
    const lines = local.pages[0]!.lines
    const expected = textForRange(local, { start: { pageIndex: 0, lineIndex: 0, charIndex: 0 }, end: { pageIndex: 0, lineIndex: lines.length - 1, charIndex: lines[lines.length - 1]!.text.length } })
    expect(expected).toContain('Subtitle Text')
    expect(r).toEqual({ op: 'copy', text: expected, paragraphs: 2 })
    host.drop(o.docKey)
  }, 60_000)

  test('copy over maxParagraphs is refused with the count, no text', async () => {
    const d = await opened(deck)
    const r = await d.text({ ...copyAll, maxParagraphs: 1 })
    expect(r).toEqual({ op: 'copy', refused: 'too-large', paragraphs: 2, bytes: 0 })
    expect('text' in r).toBe(false)
  }, 60_000)

  test('copy over maxBytes is refused', async () => {
    const d = await opened(deck)
    const r = await d.text({ ...copyAll, maxBytes: 8 })
    expect(r).toMatchObject({ op: 'copy', refused: 'too-large', paragraphs: 2 })
    expect((r as { bytes: number }).bytes).toBeGreaterThan(8)
    expect('text' in r).toBe(false)
  }, 60_000)

  test('a11y paragraph:k carries text only for k', async () => {
    const d = await opened(deck)
    const all = await d.text({ op: 'a11y', unit: 0 }) as Extract<TextReply, { op: 'a11y' }>
    expect(all.paragraphs.map((p) => p.text)).toEqual(['Presentation Title Text', 'Subtitle Text'])
    for (const p of all.paragraphs) expect(p.rect.length).toBe(4)
    const one = await d.text({ op: 'a11y', unit: 0, paragraph: 1 }) as Extract<TextReply, { op: 'a11y' }>
    expect(one.paragraphs).toEqual([{ rect: all.paragraphs[0]!.rect }, { rect: all.paragraphs[1]!.rect, text: 'Subtitle Text' }])
  }, 60_000)

  test('after a worker recycle, select then copy still answers', async () => {
    const p = new Pool({ workers: 1, recyclePages: 1 }); pools.push(p)
    const d = await opened(deck, 'visible-only', p)
    const { w, h } = d.open.units[0]!
    const sel = await d.text({ op: 'select', range: { anchor: { unit: 0, x: 0, y: 0 }, focus: { unit: 0, x: w, y: h } } })
    expect(sel).toMatchObject({ op: 'select', paragraphs: 2 })
    const recycledAfterSelect = p.stats().recycled
    expect(recycledAfterSelect).toBeGreaterThanOrEqual(1)
    const copy = await d.text(copyAll)
    expect(copy).toMatchObject({ op: 'copy', paragraphs: 2 })
    expect((copy as { text: string }).text).toContain('Subtitle Text')
    expect(p.stats().recycled).toBeGreaterThan(recycledAfterSelect)
  }, 60_000)

  test('a multi-unit op prepares every missing unit, then answers; a cache too small answers missing', async () => {
    // Pool-level: the first answer is { missing: [0, 1] }; both are prepared and the retry answers.
    const sheet = corpus('lo-calc-calcfields.xlsx')
    const p = new Pool({ workers: 1 }); pools.push(p)
    const d = await opened(sheet, 'visible-only', p)
    expect(d.open.units.length).toBeGreaterThan(1)
    const r = await d.text(copyAll)
    expect(r).toMatchObject({ op: 'copy' })
    expect((r as { text: string }).text.length).toBeGreaterThan(0)
    // Host-level bound: a 1-byte cache keeps nothing.
    const host = new DocumentHost(undefined, 1)
    const o = await host.open('tiny', sheet, 'visible-only')
    expect(host.text(o.docKey, { op: 'prepare', unit: 0 })).toMatchObject({ op: 'prepare' })
    expect(host.text(o.docKey, copyAll)).toEqual({ missing: units(o.units.length) })
    host.drop(o.docKey)
  }, 60_000)

  test('select across units wholly selects the units between; anchor/focus order does not matter', async () => {
    const d = await opened(corpus('lo-tdf115394.pptx'))
    const n = d.open.units.length
    expect(n).toBe(5)
    const last = d.open.units[n - 1]!
    const fwd = await d.text({ op: 'select', range: { anchor: { unit: 0, x: 0, y: 0 }, focus: { unit: n - 1, x: last.w, y: last.h } } }) as Extract<TextReply, { op: 'select' }>
    const back = await d.text({ op: 'select', range: { anchor: { unit: n - 1, x: last.w, y: last.h }, focus: { unit: 0, x: 0, y: 0 } } })
    const all = await d.text({ op: 'select', range: { all: true } }) as Extract<TextReply, { op: 'select' }>
    expect(back).toEqual(fwd)
    const middle = (r: typeof fwd) => r.units.filter((x) => x.unit > 0 && x.unit < n - 1)
    expect(middle(fwd).length).toBeGreaterThan(0)
    expect(middle(fwd)).toEqual(middle(all))
    expect(fwd.paragraphs).toBe(all.paragraphs)
  }, 60_000)

  test('malformed copy caps fail closed', async () => {
    const d = await opened(deck)
    for (const bad of [{ maxBytes: Number.NaN }, { maxBytes: undefined }, { maxParagraphs: Number.NaN }, { maxParagraphs: undefined }])
      expect(await d.text({ ...copyAll, ...bad } as never)).toMatchObject({ op: 'copy', refused: 'too-large' })
  }, 60_000)

  test('xlsx copy answers painted display text, not the formula', async () => {
    const base = (await augmentedFixtures()).find((f) => f.id === 'xlsx-hidden-rows')!.withoutHidden
    const zip = await JSZip.loadAsync(base)
    const sheet = await zip.file('xl/worksheets/sheet1.xml')!.async('string')
    zip.file('xl/worksheets/sheet1.xml', sheet.replace(/<sheetData>[\s\S]*<\/sheetData>/, '<sheetData><row r="1"><c r="A1"><f>SUM(40,2)</f><v>42</v></c><c r="B1" t="inlineStr"><is><t>label</t></is></c></row></sheetData>'))
    const d = await opened(await zip.generateAsync({ type: 'uint8array' }))
    const r = await d.text(copyAll) as { text: string }
    expect(r.text).toContain('42')
    expect(r.text).toContain('label')
    expect(r.text).not.toContain('SUM')
    const s = await d.text({ op: 'search', query: 'SUM', caseSensitive: false, fromUnit: 0, toUnit: 0, maxMatches: 10 })
    expect(s).toEqual({ op: 'search', matches: [] })
  }, 60_000)

  test('FIXED_MARK_TEXT is not searchable', async () => {
    const d = await opened(deck)
    for (const q of [FIXED_MARK_TEXT, 'officeview', 'protected']) {
      expect(await d.text({ op: 'search', query: q, caseSensitive: false, fromUnit: 0, toUnit: 0, maxMatches: 10 })).toEqual({ op: 'search', matches: [] })
    }
    expect((await d.text(copyAll) as { text: string }).text).not.toContain('protected')
  }, 60_000)
})

const fixtures = await augmentedFixtures()

/** The generated xlsx fixtures crop the unit to column A (64 px) and their
 * control string is ~100–115 px wide, so its run is `off-surface` and the
 * join withholds it (correctly). The control check therefore also runs on a
 * copy whose column A is widened to fit; the marker check runs on both. */
async function widenColumnA(bytes: Uint8Array): Promise<Uint8Array> {
  const zip = await JSZip.loadAsync(bytes)
  for (const path of Object.keys(zip.files).filter((p) => /^xl\/worksheets\/sheet\d+\.xml$/.test(p))) {
    const xml = await zip.file(path)!.async('string')
    const col = '<col min="1" max="1" width="40" customWidth="1"/>'
    zip.file(path, xml.includes('<cols>') ? xml.replace('<cols>', `<cols>${col}`) : xml.replace('<sheetData', `<cols>${col}</cols><sheetData`))
  }
  return zip.generateAsync({ type: 'uint8array' })
}

async function allReplies(d: Awaited<ReturnType<typeof opened>>, f: (typeof fixtures)[number]) {
  const n = d.open.units.length
  const replies: TextReply[] = []
  for (const u of units(n)) replies.push(await d.text({ op: 'prepare', unit: u }), await d.text({ op: 'a11y', unit: u }))
  const search = async (query: string) => await d.text({ op: 'search', query, caseSensitive: false, fromUnit: 0, toUnit: n - 1, maxMatches: 100 }) as Extract<TextReply, { op: 'search' }>
  const marker = [await search(f.marker), await search('MARKER')]
  const control = await search(f.control)
  const copy = await d.text(copyAll) as Extract<TextReply, { op: 'copy'; text: string }>
  replies.push(...marker, control, copy, await d.text({ op: 'select', range: { all: true } }))
  return { json: JSON.stringify(replies), marker, control, copy }
}

describe('hidden constructs (visible-only)', () => {
  test.each(fixtures.map((f) => f.id))('hidden construct %s: marker never in any text reply; control is', async (id) => {
    const f = fixtures.find((x) => x.id === id)!
    const variants = f.kind === 'xlsx' ? [f.withHidden, await widenColumnA(f.withHidden)] : [f.withHidden]
    for (const [i, bytes] of variants.entries()) {
      const r = await allReplies(await opened(bytes), f)
      expect(r.json).not.toContain(f.marker)
      expect(r.json).not.toContain('MARKER')
      // Search for the marker answers nothing at all (not even rects).
      expect(r.marker).toEqual([{ op: 'search', matches: [] }, { op: 'search', matches: [] }])
      if (f.kind === 'xlsx' && i === 0) {
        // Cropped unit: the control overflows it and is withheld (fail closed).
        expect(r.control.matches).toEqual([])
        continue
      }
      expect(r.control.matches.length).toBeGreaterThanOrEqual(1)
      expect(r.copy.text).toContain(f.control)
    }
  }, 60_000)

  test('contentPolicy full: the docx-vanish marker is answered', async () => {
    const f = fixtures.find((x) => x.id === 'docx-vanish')!
    const d = await opened(f.withHidden, 'full')
    const copy = await d.text(copyAll) as Extract<TextReply, { op: 'copy'; text: string }>
    expect(copy.text).toContain(f.marker)
    const s = await d.text({ op: 'search', query: f.marker, caseSensitive: true, fromUnit: 0, toUnit: d.open.units.length - 1, maxMatches: 10 }) as Extract<TextReply, { op: 'search' }>
    expect(s.matches.length).toBe(1)
  }, 60_000)
})

/** Every string reachable from a served unit (cycles followed once). */
function strings(v: unknown, out: string[] = [], seen = new WeakSet<object>()): string[] {
  if (typeof v === 'string') out.push(v)
  else if (v && typeof v === 'object' && !seen.has(v)) { seen.add(v); for (const x of Object.values(v)) strings(x, out, seen) }
  return out
}
const stripped = (s: string) => s.replace(/\s+/gu, '')

/** The sanitiser invariant: span text is served text; a line is its spans;
 * a source holds only its own spans' text (bar whitespace). */
function assertInvariant(unit: ServedUnit, servedTexts: string[]) {
  const allowed = new Set(servedTexts)
  const sources = new Map<LogicalTextSource, TextSpan[]>()
  for (const line of unit.lines) {
    expect(line.text).toBe(line.spans.map((s) => s.text).join(''))
    for (const s of line.spans) {
      expect(allowed.has(s.text)).toBe(true)
      if (s.logical) { const l = sources.get(s.logical.source) ?? []; l.push(s); sources.set(s.logical.source, l) }
    }
  }
  for (const [source, spans] of sources) {
    spans.sort((a, b) => a.logical!.start - b.logical!.start)
    let at = 0, covered = ''
    for (const s of spans) { if (s.logical!.end > at) { covered += source.text.slice(Math.max(at, s.logical!.start), s.logical!.end); at = s.logical!.end } }
    expect(stripped(source.text)).toBe(stripped(covered))
  }
  const known = new Set([...unit.lines.map((l) => l.text), ...servedTexts, ...[...sources.keys()].map((s) => s.text), 'vertical'])
  for (const s of strings(unit)) expect(known.has(s), `stray string ${JSON.stringify(s)}`).toBe(true)
}

describe('sanitised served units', () => {
  // One paragraph over three lines; the middle line's run is covered.
  const text = 'Alpha line one Secret covered words Omega line three'
  const source: LogicalTextSource = { text, scope: {}, order: 0 }
  const parts = [['Alpha line one ', 0], ['Secret covered words ', 15], ['Omega line three', 36]] as const
  const spans: TextSpan[] = parts.map(([t, start], i) => ({ text: t, x: 10, y: 20 + 20 * i, width: t.length * 6, fontSize: 12, logical: { source, start, end: start + t.length } }))
  const runs: RunDetail[] = parts.map(([t], i) => ({ seq: i, text: t, ink: { x: 10, y: 10 + 20 * i, width: t.length * 6, height: 12 }, servable: i !== 1, ...(i === 1 ? { reason: 'covered' as const } : {}) }))
  const unit = buildServedUnit(spans, runs, 1)
  const ask = (job: Parameters<typeof answerText>[0]) => answerText(job, 1, () => unit)
  const covered = ['Secret', 'covered', 'words']

  test('the covered middle line is withheld and nowhere in the served unit', () => {
    expect(unit).toMatchObject({ served: 2, withheld: 1 })
    for (const s of strings(unit)) for (const w of covered) expect(s).not.toContain(w)
    assertInvariant(unit, [parts[0][0], parts[2][0]])
  })

  test('copy across the covered line does not reveal it', () => {
    const r = ask({ op: 'copy', range: { anchor: { unit: 0, x: 0, y: 0 }, focus: { unit: 0, x: 400, y: 80 } }, maxParagraphs: null, maxBytes: 1 << 20 }) as { text: string; paragraphs: number }
    expect(r.text).toBe('Alpha line one \nOmega line three')
    expect(r.paragraphs).toBe(2)
  })

  test('search straddling served and withheld text answers nothing', () => {
    for (const query of ['one Secret', 'one S', 'words Omega', 'Secret', 'covered'])
      expect(ask({ op: 'search', query, caseSensitive: false, fromUnit: 0, toUnit: 0, maxMatches: 10 })).toEqual({ op: 'search', matches: [] })
    expect((ask({ op: 'search', query: 'line', caseSensitive: false, fromUnit: 0, toUnit: 0, maxMatches: 10 }) as { matches: unknown[] }).matches.length).toBe(2)
  })

  test('a11y text does not reveal the covered line', () => {
    const r = ask({ op: 'a11y', unit: 0 }) as Extract<TextReply, { op: 'a11y' }>
    expect(r.paragraphs.map((p) => p.text)).toEqual(['Alpha line one ', 'Omega line three'])
  })

  test('a span whose source slice differs from its text loses logical', () => {
    const odd: TextSpan[] = [{ text: 'HELLO', x: 0, y: 20, width: 30, fontSize: 12, logical: { source: { text: 'hello' }, start: 0, end: 5 } }]
    const u = buildServedUnit(odd, [{ seq: 0, text: 'HELLO', ink: { x: 0, y: 10, width: 30, height: 12 }, servable: true }], 1)
    expect(u.served).toBe(1)
    expect(u.lines[0]!.spans[0]!.logical).toBeUndefined()
    expect(strings(u)).not.toContain('hello')
  })

  test('cached served units on real documents hold no unserved text', async () => {
    const host = new DocumentHost()
    const counts: string[] = []
    for (const name of ['pypptx-test.pptx', 'lo-tdf146223.pptx', 'lo-tdf104788.pptx', 'poi-table-indent.docx', 'poi-rtl.docx', 'poi-56295.xlsx', 'lo-calc-calcfields.xlsx']) {
      const o = await host.open(name, corpus(name), 'visible-only')
      let served = 0, withheld = 0
      for (const u of units(o.units.length)) {
        const r = host.text(o.docKey, { op: 'prepare', unit: u }) as Extract<TextReply, { op: 'prepare' }>
        served += r.served; withheld += r.withheld
        const unit = host.servedUnit(o.docKey, u)!
        // An independent join (recorded order, scale 1, unit-size canvas, no mark).
        const paint = (host as unknown as { docs: Map<string, { paintables: Array<{ paint(c: CanvasRenderingContext2D): void }> }> }).docs.get(o.docKey)!.paintables[u]!
        const spans = recordSpans((c) => paint.paint(c), host.backend.createCanvas(4, 4).getContext('2d') as unknown as CanvasRenderingContext2D)
        const { w, h } = o.units[u]!
        const occ = occlusionObserver(w, h)
        paint.paint(occ.wrap(host.backend.createCanvas(w, h).getContext('2d') as unknown as CanvasRenderingContext2D))
        const join = joinSpans(spans, occ.runDetails(), 1)
        expect({ served: join.served.length, withheld: join.withheld }).toEqual({ served: r.served, withheld: r.withheld })
        const cached = unit.lines.flatMap((l) => l.spans.map((s) => s.text)).sort()
        expect(cached).toEqual(join.served.map((s) => s.text).sort())
        assertInvariant(unit, join.served.map((s) => s.text))
      }
      counts.push(`${name} served=${served} withheld=${withheld}`)
      host.drop(o.docKey)
    }
    console.log(`[served] ${counts.join('; ')}`)
  }, 120_000)
})
