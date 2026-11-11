/** Stage 3a Task 3 — served text index in the render worker; `text` jobs. */
import { afterAll, describe, expect, test } from 'bun:test'
import JSZip from 'jszip'
import { buildTextIndex, type TextSpan } from '../../src/core/search'
import { textForRange } from '../../src/core/selection'
import { RECORD_TEXT, type LogicalTextRange, type LogicalTextSource } from '../../src/core/text-recording'
import { occlusionObserver, type RunDetail } from '../src/occlusion'
import { joinSpans } from '../src/text/join'
import { DEFAULT_CAPS } from '../src/admission'
import { serverBackend } from '../src/backend'
import { AdmissionRefused, Pool, TEXT_JOB_BYTES, TextTooLarge, type PoolTextJob } from '../src/pool'
import { DocumentHost, FIXED_MARK_TEXT } from '../src/render'
import { answerText, buildServedUnit, OBSERVER_BUDGET, prepareWorkspace, recordSpans, TEXT_PREPARE_ALLOWANCE, TEXT_RECORD_BUDGET, type ServedUnit, type TextReply } from '../src/text/served'
import { augmentedFixtures } from './fixtures'
import { buildDocx, buildXlsx } from '../../src/testdata/ooxml-builders'
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

describe('fix round 1: prepare admission and recycle on refusal', () => {
  test('a prepare under an insufficient grant is refused before any canvas or index is built', async () => {
    const base = serverBackend()
    const sizes: Array<[number, number]> = []
    const counting = { ...base, createCanvas: (w: number, h: number) => { sizes.push([w, h]); return base.createCanvas(w, h) } }
    const host = new DocumentHost(counting)
    const o = await host.open('g', deck, 'visible-only')
    const { w, h } = o.units[0]!
    const need = prepareWorkspace(w, h)
    expect(need).toBe(4 * w * h + TEXT_PREPARE_ALLOWANCE)
    sizes.length = 0
    expect(() => host.text(o.docKey, { op: 'prepare', unit: 0 }, need - 1)).toThrow(expect.objectContaining({ name: 'RenderRefusal', code: 'grant' }))
    expect(sizes).toEqual([])
    expect(host.pages).toBe(0)
    expect(host.servedUnit(o.docKey, 0)).toBeUndefined()
    expect(host.text(o.docKey, { op: 'prepare', unit: 0 }, need)).toMatchObject({ op: 'prepare', served: 2 })
    expect(sizes).toContainEqual([w, h])
    host.drop(o.docKey)
  }, 60_000)

  test('the pool charges each prepare 4·w·h + allowance; a grant the session cannot hold is refused before the worker paints', async () => {
    const probe = new Pool({ workers: 1 }); pools.push(probe)
    const { w, h } = (await probe.open('p', deck, 'visible-only')).units[0]!
    const p = new Pool({ workers: 1, caps: { ...DEFAULT_CAPS, sessionTransientBytes: TEXT_JOB_BYTES + prepareWorkspace(w, h) - 1 } }); pools.push(p)
    const d = await opened(deck, 'visible-only', p)
    await expect(d.text(copyAll)).rejects.toBeInstanceOf(AdmissionRefused)
    await expect(d.text({ op: 'prepare', unit: 0 })).rejects.toBeInstanceOf(AdmissionRefused)
    expect(p.stats().workers[0]!.pages).toBe(0)
    expect(p.stats()).toMatchObject({ reserved: 0, running: 0, queued: 0 })
    const ok = new Pool({ workers: 1, caps: { ...DEFAULT_CAPS, sessionTransientBytes: TEXT_JOB_BYTES + prepareWorkspace(w, h) } }); pools.push(ok)
    const e = await opened(deck, 'visible-only', ok)
    expect(await e.text(copyAll)).toMatchObject({ op: 'copy', paragraphs: 2 })
    expect(ok.stats()).toMatchObject({ reserved: 0, running: 0 })
  }, 60_000)

  test('a range beyond the served cache is refused with TextTooLarge, the worker still recycles, charges end once', async () => {
    const name = 'lo-tdf115394.pptx', bytes = corpus(name)
    const probe = new Pool({ workers: 1 }); pools.push(probe)
    const pd = await opened(bytes, 'visible-only', probe)
    const sizes: number[] = []
    for (const u of units(pd.open.units.length)) sizes.push((await pd.text({ op: 'prepare', unit: u }) as { bytes: number }).bytes)
    expect(Math.min(...sizes)).toBeGreaterThan(0)
    // Room for the largest single unit, never for two.
    const p = new Pool({ workers: 1, recyclePages: 3, servedCacheBytes: Math.max(...sizes) }); pools.push(p)
    const d = await opened(bytes, 'visible-only', p)
    const before = p.stats()
    const err = await d.text(copyAll).catch((e) => e)
    expect(err).toBeInstanceOf(TextTooLarge)
    expect((err as TextTooLarge).missing.length).toBeGreaterThan(0)
    const after = p.stats()
    expect(after.recycled).toBe(before.recycled + 1)
    expect(after.workers[0]!.generation).toBe(before.workers[0]!.generation + 1)
    expect(after).toMatchObject({ reserved: 0, running: 0, queued: 0 })
    // The pool still answers afterwards.
    expect(await d.text({ op: 'a11y', unit: 0 })).toMatchObject({ op: 'a11y' })
    expect(p.stats()).toMatchObject({ reserved: 0, running: 0 })
  }, 60_000)
})

describe('fix round 2: running text-work limits and size metadata lifecycle', () => {
  type Doc = { paintables: Array<{ spec: { widthPx: number; heightPx: number }; paint(c: CanvasRenderingContext2D): void }> }
  async function hostWith(paint: (c: CanvasRenderingContext2D) => void) {
    const host = new DocumentHost()
    const o = await host.open('lim', deck, 'visible-only')
    const doc = (host as unknown as { docs: Map<string, Doc> }).docs.get(o.docKey)!
    const p0 = doc.paintables[0]!
    doc.paintables[0] = { spec: p0.spec, paint }
    const { w, h } = o.units[0]!
    return { host, docKey: o.docKey, grant: prepareWorkspace(w, h) }
  }

  test('150,000 fillText calls are refused while recording, long before the paint completes', async () => {
    let calls = 0
    const { host, docKey, grant } = await hostWith((c) => { for (let i = 0; i < 150_000; i++) { calls++; c.fillText('x', 10 + (i % 50), 20) } })
    expect(() => host.text(docKey, { op: 'prepare', unit: 0 }, grant)).toThrow(expect.objectContaining({ name: 'RenderRefusal', code: 'grant' }))
    expect(calls).toBeGreaterThan(0)
    expect(calls).toBeLessThan(150_000 / 4)
    expect(calls * 258).toBeLessThanOrEqual(TEXT_RECORD_BUDGET + 1024)
    expect(host.pages).toBe(0)
    expect(host.servedUnit(docKey, 0)).toBeUndefined()
  }, 60_000)

  test('the visual pass observer is capped: a paint of 200,000 rects is refused before it completes', async () => {
    let calls = 0
    const { host, docKey, grant } = await hostWith((c) => { c.fillText('ok', 10, 20); for (let i = 0; i < 200_000; i++) { calls++; c.fillRect(i % 90, 30, 1, 1) } })
    expect(() => host.text(docKey, { op: 'prepare', unit: 0 }, grant)).toThrow(expect.objectContaining({ name: 'RenderRefusal', code: 'grant' }))
    // The recording pass runs the paint once in full (rects are not text); the
    // observed pass must stop early.
    expect(calls - 200_000).toBeLessThan(200_000 / 2)
    expect(host.servedUnit(docKey, 0)).toBeUndefined()
    expect(TEXT_RECORD_BUDGET * 3 + OBSERVER_BUDGET * 2).toBeLessThanOrEqual(TEXT_PREPARE_ALLOWANCE)
  }, 60_000)

  test('open/forget cycles return the pool size metadata to baseline', async () => {
    const p = new Pool({ workers: 1 }); pools.push(p)
    const base = p.stats().unitSizes
    for (let i = 0; i < 4; i++) {
      await p.open(`cycle-${i}`, deck, 'visible-only')
      await p.open(`cycle-${i}`, deck, 'full')
      await p.text('s1', { docRef: `cycle-${i}`, contentPolicy: 'visible-only', op: 'prepare', unit: 0 })
      expect(p.stats().unitSizes).toBe(base + 2)
      p.forget(`cycle-${i}`)
      expect(p.stats().unitSizes).toBe(base)
    }
    p.dispose()
    expect(p.stats().unitSizes).toBe(0)
  }, 60_000)
})

describe('fix round 3: a single record is preflighted before capture maps it', () => {
  type Doc = { paintables: Array<{ spec: { widthPx: number; heightPx: number }; paint(c: CanvasRenderingContext2D): void }> }
  async function hostWith(paint: (c: CanvasRenderingContext2D) => void) {
    const host = new DocumentHost()
    const o = await host.open('pre', deck, 'visible-only')
    const doc = (host as unknown as { docs: Map<string, Doc> }).docs.get(o.docKey)!
    doc.paintables[0] = { spec: doc.paintables[0]!.spec, paint }
    const { w, h } = o.units[0]!
    return { host, docKey: o.docKey, grant: prepareWorkspace(w, h) }
  }
  /** Arrays whose every/map report how many elements core capture walked. */
  function counted<T>(items: T[], c: { points: number }): T[] {
    items.every = function (this: T[], ...a: Parameters<T[]['every']>) { c.points += this.length; return Array.prototype.every.apply(this, a as never) } as T[]['every']
    items.map = function (this: T[], ...a: Parameters<T[]['map']>) { c.points += this.length; return Array.prototype.map.apply(this, a as never) } as T[]['map']
    return items
  }
  const recordingPaint = (clusters: number, c: { points: number }) => (ctx: CanvasRenderingContext2D) => {
    const rec = (ctx as unknown as Record<symbol, unknown>)[RECORD_TEXT] as undefined | ((t: string, x: number, y: number, w: number, l: LogicalTextRange) => void)
    if (!rec) { ctx.font = '16px sans-serif'; ctx.fillText('hello', 10, 20); return }
    ctx.font = '16px sans-serif'
    const polygon = counted<[number, number]>([[0, 0], [100, 0], [100, 40], [0, 40]], c)
    const cl = counted<[number, number]>(Array.from({ length: clusters }, (_, i) => [10 + (i % 5) * 8, 15] as [number, number]), c)
    rec('hello', 10, 20, 40, { source: { text: 'hello' }, start: 0, end: 5, visual: { polygon, clusters: cl } })
  }

  test('a RECORD_TEXT record with 1,100,000 cluster points is refused with zero points mapped', async () => {
    const c = { points: 0 }
    const { host, docKey, grant } = await hostWith(recordingPaint(1_100_000, c))
    expect(() => host.text(docKey, { op: 'prepare', unit: 0 }, grant)).toThrow(expect.objectContaining({ name: 'RenderRefusal', code: 'grant' }))
    expect(c.points).toBe(0)
    expect(host.servedUnit(docKey, 0)).toBeUndefined()
  }, 60_000)

  test('a normal RECORD_TEXT record still records (and is mapped)', async () => {
    const c = { points: 0 }
    const { host, docKey, grant } = await hostWith(recordingPaint(5, c))
    expect(host.text(docKey, { op: 'prepare', unit: 0 }, grant)).toMatchObject({ op: 'prepare', served: 1, withheld: 0 })
    expect(c.points).toBeGreaterThan(0)
    expect(host.servedUnit(docKey, 0)!.lines[0]!.spans[0]!.placement!.visual!.clusters!.length).toBe(5)
  }, 60_000)

  test('an over-long RECORD_TEXT string (and its source) is refused before its geometry is mapped', async () => {
    const c = { points: 0 }
    const long = 'x'.repeat(1024 * 1024)
    const { host, docKey, grant } = await hostWith((ctx) => {
      const rec = (ctx as unknown as Record<symbol, unknown>)[RECORD_TEXT] as undefined | ((t: string, x: number, y: number, w: number, l: LogicalTextRange) => void)
      if (!rec) return
      rec(long, 0, 20, 40, { source: { text: long }, start: 0, end: long.length, visual: { polygon: counted<[number, number]>([[0, 0], [10, 0], [10, 10]], c) } })
    })
    expect(() => host.text(docKey, { op: 'prepare', unit: 0 }, grant)).toThrow(expect.objectContaining({ name: 'RenderRefusal', code: 'grant' }))
    expect(c.points).toBe(0)
  }, 60_000)
})

describe('task 3b: paragraph and cell identity (RECORD_GROUP)', () => {
  const EXAMPLE = 'The quick brown fox jumps over the lazy dog and keeps running across the wide green field until the sun goes down behind the distant hills.'
  /** A narrow page (5000 twips, 720 margins) so EXAMPLE wraps over several lines. */
  async function docx(paras: string[], tables: Parameters<typeof buildDocx>[1] = []): Promise<Uint8Array> {
    const zip = await JSZip.loadAsync(await buildDocx(paras.map((text) => ({ runs: [{ text }] })), tables))
    const doc = await zip.file('word/document.xml')!.async('string')
    zip.file('word/document.xml', doc.replace('<w:pgSz w:w="12240" w:h="15840"/>', '<w:pgSz w:w="5000" w:h="15840"/>').replace('w:right="1440" w:bottom="1440" w:left="1440"', 'w:right="720" w:bottom="1440" w:left="720"'))
    return zip.generateAsync({ type: 'uint8array' })
  }
  /** A1 wraps (wrapText, narrow column, tall row); the other cells are plain. */
  async function xlsx(others: string[]): Promise<Uint8Array> {
    const refs = others.map((_, i) => String.fromCharCode(66 + i) + '1')
    const base = await buildXlsx([{ name: 'S', cols: '<col min="1" max="1" width="12" customWidth="1"/>', rows: [
      { r: 1, cells: [{ ref: 'A1', t: 's', v: 0, style: 3 }, ...refs.map((ref, i) => ({ ref, t: 's' as const, v: i + 1 }))] },
    ] }], ['alpha beta gamma delta epsilon zeta', ...others])
    const zip = await JSZip.loadAsync(base)
    const styles = await zip.file('xl/styles.xml')!.async('string')
    zip.file('xl/styles.xml', styles.replace('<cellXfs count="3">', '<cellXfs count="4">').replace('</cellXfs>', '<xf numFmtId="0" fontId="0" fillId="0" borderId="0" xfId="0" applyAlignment="1"><alignment wrapText="1"/></xf></cellXfs>'))
    const sheet = await zip.file('xl/worksheets/sheet1.xml')!.async('string')
    zip.file('xl/worksheets/sheet1.xml', sheet.replace('<row r="1">', '<row r="1" ht="80" customHeight="1">'))
    return zip.generateAsync({ type: 'uint8array' })
  }
  /** Opens on a host and prepares unit 0; every span must be served, so the counts are not vacuous. */
  async function prepared(bytes: Uint8Array) {
    const host = new DocumentHost()
    const o = await host.open(`g-${++refs}`, bytes, 'visible-only')
    const prep = host.text(o.docKey, { op: 'prepare', unit: 0 }) as Extract<TextReply, { op: 'prepare' }>
    expect(prep.served).toBeGreaterThan(0)
    expect(prep.withheld).toBe(0)
    const ask = (job: Exclude<PoolTextJob, { op: 'prepare' }> | Record<string, unknown>) => host.text(o.docKey, job as never)
    const lineCount = host.servedUnit(o.docKey, 0)!.lines.length
    return { host, o, ask, lineCount }
  }
  const unit0 = { op: 'copy', range: { anchor: { unit: 0, x: 0, y: 0 }, focus: { unit: 0, x: 1e6, y: 1e6 } }, maxParagraphs: null, maxBytes: 4 << 20 } as const

  test('a wrapped DOCX body paragraph counts as one paragraph', async () => {
    const d = await prepared(await docx([EXAMPLE]))
    expect(d.lineCount).toBeGreaterThanOrEqual(3)
    expect(d.ask({ op: 'select', range: { all: true } })).toMatchObject({ op: 'select', paragraphs: 1 })
    const c = d.ask(unit0) as { text: string; paragraphs: number }
    expect(c.paragraphs).toBe(1)
    expect(c.text.replace(/\s+/g, ' ')).toBe(EXAMPLE)
  }, 60_000)

  test('two DOCX body paragraphs count as two', async () => {
    const d = await prepared(await docx([EXAMPLE, 'Second paragraph.']))
    expect(d.ask({ op: 'select', range: { all: true } })).toMatchObject({ op: 'select', paragraphs: 2 })
  }, 60_000)

  test('a DOCX table with 2×2 single-paragraph cells counts four', async () => {
    const cell = (text: string) => ({ paragraphs: [{ runs: [{ text }] }] })
    const d = await prepared(await docx([], [{ gridCols: ['1500', '1500'], rows: [{ cells: [cell('north west'), cell('north east')] }, { cells: [cell('south west'), cell('south east')] }] }]))
    expect(d.ask(unit0)).toMatchObject({ op: 'copy', paragraphs: 4 })
  }, 60_000)

  test('a wrapped XLSX cell counts as one; N non-empty cells count N', async () => {
    const one = await prepared(await xlsx([]))
    expect(one.lineCount).toBeGreaterThan(1)
    expect(one.ask(copyAll)).toMatchObject({ op: 'copy', paragraphs: 1 })
    const cellA11y = one.ask({ op: 'a11y', unit: 0 }) as Extract<TextReply, { op: 'a11y' }>
    expect(cellA11y.paragraphs.map((p) => p.text)).toEqual([(one.ask(copyAll) as { text: string }).text])
    const three = await prepared(await xlsx(['bee', 'sea']))
    expect(three.ask(copyAll)).toMatchObject({ op: 'copy', paragraphs: 3 })
  }, 60_000)

  test('copy maxParagraphs: 1 allows one whole wrapped paragraph and refuses two', async () => {
    const one = await prepared(await docx([EXAMPLE]))
    const ok = one.ask({ ...unit0, maxParagraphs: 1 }) as { text: string; paragraphs: number }
    expect(ok.paragraphs).toBe(1)
    expect(ok.text.replace(/\s+/g, ' ')).toBe(EXAMPLE)
    const two = await prepared(await docx([EXAMPLE, 'Second paragraph.']))
    expect(two.ask({ ...unit0, maxParagraphs: 1 })).toEqual({ op: 'copy', refused: 'too-large', paragraphs: 2, bytes: 0 })
  }, 60_000)

  test('a11y returns one entry per DOCX paragraph, not per word', async () => {
    const d = await prepared(await docx([EXAMPLE, 'Second paragraph.']))
    const r = d.ask({ op: 'a11y', unit: 0 }) as Extract<TextReply, { op: 'a11y' }>
    expect(r.paragraphs.length).toBe(2)
    expect(r.paragraphs[0]!.text!.replace(/\s+/g, ' ')).toBe(EXAMPLE)
    expect(r.paragraphs[1]!.text).toBe('Second paragraph.')
    for (const p of r.paragraphs) expect(p.rect.length).toBe(4)
    // Lines are joined as textForRange joins them: the whole-unit copy is the paragraphs joined.
    expect(r.paragraphs.map((p) => p.text).join('\n')).toBe((d.ask(copyAll) as { text: string }).text)
    const k = d.ask({ op: 'a11y', unit: 0, paragraph: 1 }) as Extract<TextReply, { op: 'a11y' }>
    expect(k.paragraphs).toEqual([{ rect: r.paragraphs[0]!.rect }, { rect: r.paragraphs[1]!.rect, text: 'Second paragraph.' }])
  }, 60_000)

  test('served units hold no painter group objects (tokens only)', async () => {
    const d = await prepared(await docx([EXAMPLE, 'Second paragraph.']))
    const unit = d.host.servedUnit(d.o.docKey, 0)!
    assertInvariant(unit, unit.lines.flatMap((l) => l.spans.map((s) => s.text)))
  }, 60_000)
})

describe('text decorations do not withhold the decorated word (Task 3c)', () => {
  /** One paragraph: an underlined hyperlink (words with descenders, so the
   * underline overlaps their ink) and a struck word, between plain words. */
  async function decorated(): Promise<Uint8Array> {
    const zip = await JSZip.loadAsync(await buildDocx([{ runs: [{ text: 'Visit the ' }, { text: 'typing guide' }, { text: ' today, not ' }, { text: 'gone' }, { text: ' tomorrow.' }] }]))
    let xml = await zip.file('word/document.xml')!.async('string')
    const run = (t: string) => new RegExp(`<w:r><w:rPr>([^<]*?)</w:rPr><w:t xml:space="preserve">${t}</w:t></w:r>`)
    xml = xml.replace(run('typing guide'), '<w:hyperlink w:anchor="target"><w:r><w:rPr>$1<w:color w:val="0563C1"/><w:u w:val="single"/></w:rPr><w:t xml:space="preserve">typing guide</w:t></w:r></w:hyperlink>')
    xml = xml.replace(run('gone'), '<w:r><w:rPr>$1<w:strike/></w:rPr><w:t xml:space="preserve">gone</w:t></w:r>')
    expect(xml).toContain('<w:u w:val="single"/>')
    expect(xml).toContain('<w:strike/>')
    zip.file('word/document.xml', xml)
    return zip.generateAsync({ type: 'uint8array' })
  }

  test('real DOCX with an underlined hyperlink: the link text is searchable and copy returns it with its neighbouring spaces', async () => {
    const d = await opened(await decorated())
    expect(await d.text({ op: 'prepare', unit: 0 })).toMatchObject({ op: 'prepare', withheld: 0 })
    const s = await d.text({ op: 'search', query: 'typing guide', caseSensitive: true, fromUnit: 0, toUnit: 0, maxMatches: 10 }) as Extract<TextReply, { op: 'search' }>
    expect(s.matches.length).toBe(1)
    const g = await d.text({ op: 'search', query: 'not gone tomorrow', caseSensitive: true, fromUnit: 0, toUnit: 0, maxMatches: 10 }) as Extract<TextReply, { op: 'search' }>
    expect(g.matches.length).toBe(1)
    const copy = await d.text(copyAll) as Extract<TextReply, { op: 'copy'; text: string }>
    expect(copy.text).toBe('Visit the typing guide today, not gone tomorrow.')
  }, 60_000)
})
