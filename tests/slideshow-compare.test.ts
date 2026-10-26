import { existsSync, mkdirSync, mkdtempSync, readFileSync, rmSync, writeFileSync } from 'node:fs'
import { tmpdir } from 'node:os'
import { join } from 'node:path'
import { fileURLToPath } from 'node:url'
import { afterEach, describe, expect, it } from 'vitest'
import { CATALOGUE } from '../src/slideshow/catalogue'
import {
  CAPTURE_PCTS, TOLERANCE_APPROX, TOLERANCE_EXACT, compareFrames, frameKey, parseProvenance, parseRefName,
  readCompareRecords, readRefs, refFileName, rowStatus, runCompareCli, writeCompareRecord, type CompareRecord,
} from '../scripts/slideshow-compare'
import { BEHAVIOURAL_QUESTIONS } from '../scripts/slideshow-capture-questions'

function makeImage(w: number, h: number, fill: (x: number, y: number) => [number, number, number, number?]): ImageData {
  const data = new Uint8ClampedArray(w * h * 4)
  for (let y = 0; y < h; y++) for (let x = 0; x < w; x++) {
    const [r, g, b, a = 255] = fill(x, y)
    const i = (y * w + x) * 4
    data[i] = r; data[i + 1] = g; data[i + 2] = b; data[i + 3] = a
  }
  return { data, width: w, height: h, colorSpace: 'srgb' } as ImageData
}
const base = () => makeImage(32, 32, () => [10, 20, 30])
/** One row (`row`) is `delta` levels brighter on every colour channel. */
const striped = (delta: number, row = 7) => makeImage(32, 32, (_x, y) => y === row ? [10 + delta, 20 + delta, 30 + delta] : [10, 20, 30])

const temps: string[] = []
const tempDir = () => { const d = mkdtempSync(join(tmpdir(), 'ov-refs-')); temps.push(d); return d }
afterEach(() => { while (temps.length) rmSync(temps.pop()!, { recursive: true, force: true }) })

const PROVENANCE = { powerpoint: 'Microsoft PowerPoint for Mac 16.89 (24091630)', os: 'macOS 15.1', resolution: [1920, 1080], capturedBy: 'Julian Shen', date: '2026-11-02' }
/** `provenance: null` writes no provenance.json. */
function deckDir(root: string, deck: string, files: string[], provenance: unknown = PROVENANCE) {
  const dir = join(root, deck)
  mkdirSync(dir, { recursive: true })
  if (provenance !== null) writeFileSync(join(dir, 'provenance.json'), JSON.stringify(provenance))
  for (const f of files) writeFileSync(join(dir, f), 'png')
}
const record = (over: Partial<CompareRecord> = {}): CompareRecord =>
  ({ row: 'transition:p:fade', deck: 'tdf115394', slide: 2, step: 0, pct: 50, pass: true, meanAbs: 0.001, rowMax: 0.01, ...over })

describe('compareFrames', () => {
  it('uses the catalogue tolerances, as fractions of 255', () => {
    const exactRow = CATALOGUE.find(r => r.status === 'exact')!
    const approxRow = CATALOGUE.find(r => r.status === 'approximate')!
    expect(TOLERANCE_EXACT).toEqual(exactRow.tolerance)
    expect(TOLERANCE_APPROX).toEqual(approxRow.tolerance)
    expect(TOLERANCE_EXACT).toEqual({ meanAbs: 0.02, rowMax: 0.10 })
    expect(TOLERANCE_APPROX).toEqual({ meanAbs: 0.08, rowMax: 0.30 })
  })

  it('identical frames pass; a one-row stripe fails rowMax but passes meanAbs', () => {
    const same = compareFrames(base(), base(), TOLERANCE_EXACT)
    expect(same).toMatchObject({ pass: true, meanAbs: 0, rowMax: 0 })
    // +40 levels on row 7 of 32: row mean 40/255 = 0.157 > 0.10; frame mean 0.157/32 = 0.0049 <= 0.02.
    const r = compareFrames(base(), striped(40), TOLERANCE_EXACT)
    expect(r.rowMax).toBeCloseTo(40 / 255, 6)
    expect(r.meanAbs).toBeCloseTo(40 / 255 / 32, 6)
    expect(r.meanAbs).toBeLessThanOrEqual(TOLERANCE_EXACT.meanAbs)
    expect(r.rowMax).toBeGreaterThan(TOLERANCE_EXACT.rowMax)
    expect(r.worstRow).toBe(7)
    expect(r.pass).toBe(false)
    // +20 levels (0.078) is inside the exact row limit.
    expect(compareFrames(base(), striped(20, 30), TOLERANCE_EXACT)).toMatchObject({ pass: true, worstRow: 30 })
  })

  it('approximate rows use the looser tolerance', () => {
    // 0.157 sits between the exact (0.10) and approximate (0.30) row limits.
    expect(compareFrames(base(), striped(40), TOLERANCE_EXACT).pass).toBe(false)
    expect(compareFrames(base(), striped(40), TOLERANCE_APPROX).pass).toBe(true)
    // +80 levels (0.314) fails even the approximate row limit.
    expect(compareFrames(base(), striped(80), TOLERANCE_APPROX).pass).toBe(false)
    // A uniform +10 shift: mean 0.039 fails exact meanAbs (0.02), passes approximate (0.08).
    const shifted = makeImage(32, 32, () => [20, 30, 40])
    expect(compareFrames(base(), shifted, TOLERANCE_EXACT)).toMatchObject({ pass: false })
    expect(compareFrames(base(), shifted, TOLERANCE_APPROX)).toMatchObject({ pass: true })
  })

  it('a single-channel or alpha difference is not diluted', () => {
    const red = makeImage(32, 32, (_x, y) => y === 3 ? [50, 20, 30] : [10, 20, 30])
    expect(compareFrames(base(), red, TOLERANCE_EXACT)).toMatchObject({ pass: false, worstRow: 3 })
    expect(compareFrames(base(), red, TOLERANCE_EXACT).rowMax).toBeCloseTo(40 / 255, 6)
    const clear = makeImage(32, 32, () => [10, 20, 30, 0])
    expect(compareFrames(base(), clear, TOLERANCE_APPROX)).toMatchObject({ pass: false, meanAbs: 1 })
  })

  it('a size mismatch or an empty or malformed frame never passes', () => {
    const solid = (w: number, h: number) => makeImage(w, h, () => [0, 0, 0])
    for (const [a, b] of [[solid(64, 64), solid(32, 32)], [solid(32, 32), solid(32, 16)], [solid(32, 16), solid(16, 32)], [solid(16, 16), solid(32, 32)]]) {
      const r = compareFrames(a, b, TOLERANCE_APPROX)
      expect(r).toMatchObject({ pass: false, reason: 'size-mismatch' })
    }
    expect(compareFrames(solid(0, 0), solid(0, 0), TOLERANCE_APPROX)).toMatchObject({ pass: false, reason: 'empty' })
    const short = { data: new Uint8ClampedArray(16), width: 32, height: 32 } as ImageData
    expect(compareFrames(short, short, TOLERANCE_APPROX)).toMatchObject({ pass: false, reason: 'malformed' })
  })
})

describe('reference layout', () => {
  it('names frames <slide>-<step>-<pct>.png and parses only that form', () => {
    expect(CAPTURE_PCTS).toEqual([0, 25, 50, 75, 100])
    expect(refFileName(6, 2, 25)).toBe('6-2-25.png')
    expect(parseRefName('6-2-25.png')).toEqual({ slide: 6, step: 2, pct: 25 })
    expect(parseRefName('1-0-100.png')).toEqual({ slide: 1, step: 0, pct: 100 })
    for (const bad of ['garbage.png', '0-1-50.png', '1-1-33.png', '01-1-50.png', '1-1-050.png', '1-1-50.PNG', '1-1-50.jpg', '1-1.png'])
      expect(parseRefName(bad), bad).toBeUndefined()
  })

  it('validates provenance and rejects placeholders', () => {
    expect(parseProvenance(PROVENANCE)).toEqual({ ok: true, value: PROVENANCE })
    const placeholder = { powerpoint: 'unknown', os: 'unknown', resolution: [1920, 1080], capturedBy: 'unknown', date: 'unknown' }
    const bad = parseProvenance(placeholder)
    expect(bad.ok).toBe(false)
    if (!bad.ok) expect(bad.errors.join(' ')).toMatch(/powerpoint.*os.*capturedBy.*date/s)
    expect(parseProvenance({ ...PROVENANCE, resolution: [1920] }).ok).toBe(false)
    expect(parseProvenance({ ...PROVENANCE, resolution: [0, 1080] }).ok).toBe(false)
    expect(parseProvenance({ ...PROVENANCE, date: '2026-13-01' }).ok).toBe(false)
    expect(parseProvenance({ ...PROVENANCE, capturedBy: undefined }).ok).toBe(false)
    expect(parseProvenance(null).ok).toBe(false)
  })

  it('round-trips exact frame timing while continuing to accept old provenance', () => {
    const withFrameTimes = {
      ...PROVENANCE,
      frames: { '2-0-25.png': { atMs: 264.026, movieFrame: 8 } },
    }
    expect(parseProvenance(JSON.parse(JSON.stringify(withFrameTimes)))).toEqual({ ok: true, value: withFrameTimes })
    expect(parseProvenance(PROVENANCE)).toEqual({ ok: true, value: PROVENANCE })
  })

  it('readRefs tells a missing directory from an empty one and skips malformed names', async () => {
    const root = tempDir()
    const missing = await readRefs(join(root, 'nope'))
    expect(missing).toMatchObject({ missing: true, frames: [] })
    const empty = await readRefs(root)
    expect(empty).toMatchObject({ missing: false, frames: [] })
    deckDir(root, 'tdf115394', ['2-0-50.png', 'garbage.png', '2-0-33.png'])
    mkdirSync(join(root, 'tdf115394', 'video'))
    deckDir(root, 'tdf44223', ['1-1-50.png'], { ...PROVENANCE, powerpoint: 'unknown' })
    deckDir(root, 'narration', ['1-0-0.png'], null)
    const set = await readRefs(root)
    expect(set.frames).toEqual([{ deck: 'tdf115394', slide: 2, step: 0, pct: 50, path: join(root, 'tdf115394', '2-0-50.png') }])
    expect(set.provenance.get('tdf115394')).toEqual(PROVENANCE)
    expect(set.problems.join('\n')).toMatch(/tdf115394\/garbage\.png/)
    expect(set.problems.join('\n')).toMatch(/tdf115394\/2-0-33\.png/)
    expect(set.problems.join('\n')).toMatch(/tdf44223\/provenance\.json.*powerpoint/)
    expect(set.problems.join('\n')).toMatch(/narration\/provenance\.json.*missing/)
  })
})

describe('verification status', () => {
  it('a missing ref leaves the row unverified, never failed', async () => {
    const root = tempDir()
    // No refs at all, even with a failing record: unverified.
    const none = await readRefs(join(root, 'absent'))
    expect(rowStatus('transition:p:fade', none, [record({ pass: false })])).toBe('unverified')
    expect(rowStatus('transition:p:fade', none, [record({ pass: false })])).not.toBe('failed')
    // A passing record whose ref is not on disk does not verify either.
    deckDir(root, 'tdf115394', ['2-0-0.png'])
    const some = await readRefs(root)
    expect(rowStatus('transition:p:fade', some, [record({ pct: 50 })])).toBe('unverified')
    // A ref without any record stays unverified.
    expect(rowStatus('transition:p:fade', some, [])).toBe('unverified')
    // A deck whose provenance is invalid contributes no refs.
    deckDir(root, 'tdf142915', ['1-0-50.png'], { ...PROVENANCE, date: 'soon' })
    const bad = await readRefs(root)
    expect(rowStatus('transition:p:none', bad, [record({ row: 'transition:p:none', deck: 'tdf142915', slide: 1, pct: 50 })])).toBe('unverified')
  })

  it('a row is verified only by passing records for refs on disk, and failed by any failing one', async () => {
    const root = tempDir()
    deckDir(root, 'tdf115394', ['2-0-0.png', '2-0-50.png', '2-0-100.png'])
    const refs = await readRefs(root)
    expect(frameKey('tdf115394', 2, 0, 50)).toBe('tdf115394/2-0-50')
    const pass = [record({ pct: 0 }), record({ pct: 50 }), record({ pct: 100 })]
    expect(rowStatus('transition:p:fade', refs, pass)).toBe('verified')
    expect(rowStatus('transition:p:fade', refs, [...pass, record({ pct: 50, pass: false })])).toBe('failed')
    expect(rowStatus('transition:p:push', refs, pass)).toBe('unverified')
  })

  it('writeCompareRecord round-trips through readCompareRecords and creates the directory', () => {
    const root = tempDir()
    const path = join(root, 'nested', 'dir', 'slideshow-compare.json')
    const rows = [record(), record({ pct: 100, pass: false, meanAbs: 0.5, rowMax: 0.9 })]
    writeCompareRecord(path, rows)
    expect(readCompareRecords(path)).toEqual(rows)
    expect(readFileSync(path, 'utf8').endsWith('\n')).toBe(true)
    expect(readCompareRecords(join(root, 'none.json'))).toEqual([])
    writeFileSync(path, JSON.stringify([{ row: 'x' }]))
    expect(() => readCompareRecords(path)).toThrow(/record 0/)
  })
})

describe('capture checklist', () => {
  const doc = (rel: string) => readFileSync(fileURLToPath(new URL(rel, import.meta.url)), 'utf8')
  const checklist = doc('../docs/validation/slideshow-captures.md')
  const corpusDecks = [...doc('../docs/validation/slideshow-corpus.md').matchAll(/^\| \[([\w-]+)\.pptx\]\(/gm)].map(m => m[1])

  it('the checklist lists every corpus deck and every behavioural question', () => {
    expect(corpusDecks).toHaveLength(16)
    for (const deck of corpusDecks) {
      expect(checklist, deck).toContain(`<!-- capture-plan:${deck}:begin -->`)
      expect(checklist, deck).toContain(`<!-- capture-plan:${deck}:end -->`)
      expect(checklist, deck).toContain(`tests/fixtures/slideshow-refs/${deck}/`)
    }
    const ids = BEHAVIOURAL_QUESTIONS.map(q => q.id)
    // The seven questions the brief names, each by its id.
    for (const id of ['Q-spd', 'Q-advtm-pending', 'Q-trigger-link', 'Q-kiosk-restart', 'Q-kiosk-input', 'Q-previous-into', 'Q-med-700'])
      expect(ids).toContain(id)
    for (const q of BEHAVIOURAL_QUESTIONS) {
      const at = checklist.indexOf(`### ${q.id}:`)
      expect(at, q.id).toBeGreaterThan(-1)
      const section = checklist.slice(at, checklist.indexOf('\n### ', at + 1) >>> 0)
      for (const field of ['**Open**', '**Do**', '**Record**', '**Engine assumption**', '**Verifies**', '**Answer**'])
        expect(section, `${q.id} ${field}`).toContain(field)
    }
    expect(checklist).not.toMatch(/^EOF$|^echo done/m)
    expect(checklist).not.toContain('tests/fixtures/slideshow/')
    expect(checklist.endsWith('\n')).toBe(true)
  })

  it('every frame file in the checklist parses as a reference name', () => {
    const files = [...checklist.matchAll(/`(\d[^`]*\.png)`/g)].map(m => m[1])
    expect(files.length).toBeGreaterThan(100)
    for (const f of files) expect(parseRefName(f), f).toBeDefined()
  })

  it.skipIf(!existsSync('corpus/lo-tdf104788.pptx'))('the deck tables match the plan generated from the corpus', async () => {
    const { CORPUS_DECKS, PLAN_BEGIN, PLAN_END, planDeck, renderPlan } = await import('../scripts/slideshow-capture-plan')
    expect([...CORPUS_DECKS]).toEqual(corpusDecks)
    for (const deck of CORPUS_DECKS) {
      const plan = await planDeck(deck, new Uint8Array(readFileSync(`corpus/lo-${deck}.pptx`)))
      const start = checklist.indexOf(PLAN_BEGIN(deck)) + PLAN_BEGIN(deck).length
      expect(checklist.slice(start, checklist.indexOf(PLAN_END(deck))).trim(), deck).toBe(renderPlan(plan).trim())
    }
  })
})

describe.skipIf(!existsSync('corpus/lo-tdf146223.pptx'))('compare CLI (S1-22)', () => {
  it('renders the checklist frames through the engine, compares the ones with references and writes records', async () => {
    // tdf146223: two 4:3 slides with picture backgrounds, each fading in (spd slow, 1000 ms); slide 1 fades from black.
    const dir = tempDir()
    const out = join(dir, 'ours'), refs = join(dir, 'refs'), records = join(dir, 'records.json')
    const args = ['tdf146223', '--refs', refs, '--records', records]
    // No references yet: every checklist frame is rendered (and saved with --out) and listed as unverified; exit 0.
    expect(await runCompareCli([...args, '--out', out])).toBe(0)
    const ours = (f: string) => readFileSync(join(out, 'tdf146223', f))
    for (const slide of [1, 2]) for (const pct of [0, 25, 50, 75, 100]) expect(ours(`${slide}-0-${pct}.png`).subarray(1, 4).toString()).toBe('PNG')
    expect(ours('1-0-50.png').equals(ours('1-0-100.png'))).toBe(false)
    // Our own frames as references (provenance at the rendered 1280 × 960) pass with the row's tolerance.
    deckDir(refs, 'tdf146223', [], {
      ...PROVENANCE,
      resolution: [1280, 960],
      frames: {
        '1-0-50.png': { atMs: 250, movieFrame: 8 },
        '2-0-100.png': { atMs: 1000, movieFrame: 30 },
      },
    })
    // Label a 25% rendered image as 50%: this passes only if compare uses the
    // reference's exact atMs (250), rather than the nominal 500 ms time.
    writeFileSync(join(refs, 'tdf146223', '1-0-50.png'), ours('1-0-25.png'))
    writeFileSync(join(refs, 'tdf146223', '2-0-100.png'), ours('2-0-100.png'))
    expect(await runCompareCli([...args, '--write'])).toBe(0)
    expect(readCompareRecords(records).map(r => [r.row, r.slide, r.step, r.pct, r.pass])).toEqual([
      ['transition:p:fade', 1, 0, 50, true], ['transition:p:fade', 2, 0, 100, true],
    ])
    // A reference that differs (slide 1's settled frame given for its 50 % frame) fails: exit 1, and the record says so.
    writeFileSync(join(refs, 'tdf146223', '1-0-50.png'), ours('1-0-100.png'))
    expect(await runCompareCli([...args, '--write', '--rows', 'transition:p:fade'])).toBe(1)
    expect(readCompareRecords(records).find(r => r.slide === 1)?.pass).toBe(false)
    expect(readCompareRecords(records)).toHaveLength(2)
    // Usage errors exit 2.
    expect(await runCompareCli([])).toBe(2)
    expect(await runCompareCli(['tdf146223', '--bogus'])).toBe(2)
  }, 60_000)
})
