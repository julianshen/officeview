// @vitest-environment node
/**
 * The bar all-equal follow-up (brief cf-bar-allequal, Phase 1 tasks 3 and 4; owner rulings 2026-11-23): the pre-registered
 * case stage3-bar-all-equal and the rehearsal of every reading's registration.
 *
 * The case has no capture yet, so each reading is planted on a synthetic capture (Excel's measured geometry convention:
 * 84 px column A, 195 px columns from 32.51 characters, 30 px rows, 1 px per pt) by an independent painter from the
 * predicted px of the recipe (docs/validation/xlsx-cf-recipes-stage3.md §2f), run through the real measurer and
 * hypothesis family, and must come back as the one match. A no-match plant must match nothing. Then the registration
 * each reading implies is injected (reading + evidence) and production must paint the generated workbook exactly as
 * planted (+-1 px), while every earlier Stage 3 capture keeps its production outcome except the all-equal shapes the
 * reading names. This checks the family and the implementation, not Excel.
 */
import { readFileSync } from 'node:fs'
import { join, resolve } from 'node:path'
import { afterEach, describe, expect, it } from 'vitest'
import { CF_STAGE3_CAPABILITIES, CF_STAGE3_CASES, CF_STAGE3_NONDISCRIMINATING, type CfManifest } from '../scripts/cf-refs/cases'
import { CF_MANIFEST_PATH } from '../scripts/cf-refs/generate'
import { computeWorksheetRangeGeometry, evaluateBarCapabilities, measureCaseBars, readZipPartsSync, type DecodedPng } from '../scripts/cf-refs/verify'
import { OfficePackage } from '../src/core/zip'
import { parseXlsx } from '../src/xlsx/parse'
import { prepareCfPaint } from '../src/xlsx/cf/render'
import { cfRuntimeEvidence } from '../src/xlsx/cf/predicates'
import { VERIFIED_CAPABILITIES, withExactCfEvidence } from '../src/xlsx/cf/evidence'
import { injectCfBarAllEqualReadings } from '../src/xlsx/cf/dataBar'
import { layoutDataBar } from '../src/xlsx/cf/dataBarPaint'

const REPO_ROOT = resolve(__dirname, '..')
const TIMEOUT_MS = 300_000
const ID = 'stage3-bar-all-equal'
const manifest = JSON.parse(readFileSync(CF_MANIFEST_PATH, 'utf8')) as CfManifest
const c = manifest.cases.find(x => x.id === ID)!
const parts = readZipPartsSync(new Uint8Array(readFileSync(resolve(REPO_ROOT, c.saved))))!
const COLS = 'BCDEFGHIJKLMNO'.split('')
const colWidthsPt = [84, ...COLS.map(() => 195)]
const rows = Number(/\d+$/.exec(c.range)![0])
const geom = computeWorksheetRangeGeometry(parts, c.range, { colWidthsPt, rowHeightsPt: Array.from({ length: rows }, () => 30) })!
const S = 192
const halfUp = (x: number): number => Math.floor(x + 0.5 + 1e-9)
const cellLeft = (col: string): number => 84 + 195 * (col.charCodeAt(0) - 66)

const DISTINCT = ['evaluate', 'full', 'degenerate-half', 'none'] as const
const LTR = ['half', 'none', 'full'] as const
const RTL = ['half', 'half-ltr', 'none', 'full'] as const
const D1090 = ['half-span', 'half-min', 'none', 'full'] as const
type Plant = { distinct: string; ltr: string; rtl: string; d1090: string; stat: string }

/**
 * The recipe's predicted bars (offset from the cell's leading inset, inclusive length; 0 = no bar) per column and reading,
 * on the 192 px span: written from the closed formulas, independent of the hypothesis code.
 */
type Seg = { from: number; len: number; rgb?: readonly number[] }
const BLUE = [0, 0, 255], RED = [255, 0, 0]
const lin = (t: number): Seg => ({ from: 0, len: halfUp(t * S) })
const d1090 = (t: number): Seg => ({ from: 0, len: halfUp((0.05 + 0.8 * t) * S) })
const NONE: Seg = { from: 0, len: 0 }
function predicted(col: string, row: number, p: Plant): Seg {
  const r = p.distinct
  switch (col) {
    case 'B': return r === 'none' ? NONE : r === 'degenerate-half' ? lin(0.5) : lin(1)
    case 'C': return r === 'none' || r === 'evaluate' ? NONE : r === 'full' ? lin(1) : lin(0.5)
    case 'D': return r === 'none' ? NONE : lin(r === 'evaluate' ? 0.3 : r === 'full' ? 1 : 0.5)
    case 'E': return r === 'none' ? NONE : lin(r === 'evaluate' ? 0.7 : r === 'full' ? 1 : 0.5)
    case 'F': return r === 'none' ? NONE : d1090(r === 'evaluate' ? 0 : r === 'full' ? 1 : 0.5)
    case 'G': return r === 'none' ? NONE : d1090(r === 'evaluate' ? 0.3 : r === 'full' ? 1 : 0.5)
    case 'H': return r === 'none' ? NONE : d1090(r === 'degenerate-half' ? 0.5 : 1)
    case 'I': {
      // num -100/-20, all -40: grows from the right edge, t = (v - hi) / (lo - hi); red (the negative fill).
      if (r === 'none') return NONE
      const len = halfUp((r === 'evaluate' ? 0.25 : r === 'full' ? 1 : 0.5) * S)
      return { from: S - len, len, rgb: RED }
    }
    case 'J': return lin([0.3, 0.6, 0.9][row - 2]!)
    case 'K': return d1090([0, 0.3, 1][row - 2]!)
    case 'L': return p.stat === 'none' ? NONE : lin(p.stat === 'full' ? 1 : 0.5)
    case 'M': return p.ltr === 'none' ? NONE : lin(p.ltr === 'full' ? 1 : 0.5)
    case 'N': {
      if (p.rtl === 'none') return NONE
      const len = halfUp((p.rtl === 'full' ? 1 : 0.5) * S)
      return { from: p.rtl === 'half-ltr' ? 0 : S - len, len }
    }
    case 'O': return p.d1090 === 'none' ? NONE : p.d1090 === 'half-span' ? lin(0.5) : d1090(p.d1090 === 'full' ? 1 : 0.5)
  }
  throw new Error(col)
}

/** Paints a synthetic capture: each designated cell's predicted bar, at Excel's insets. */
function synthetic(p: Plant, seg: (col: string, row: number, p: Plant) => Seg = predicted): DecodedPng {
  const width = colWidthsPt.reduce((a, b) => a + b, 0) + 1, height = rows * 30 + 1
  const data = new Uint8ClampedArray(width * height * 4).fill(255)
  const put = (x: number, y: number, rgb: readonly number[]) => { const j = (y * width + x) * 4; data[j] = rgb[0]!; data[j + 1] = rgb[1]!; data[j + 2] = rgb[2]! }
  for (const cell of c.expected.cells.filter(x => x.matched)) {
    const col = cell.ref.replace(/\d+/, ''), row = Number(cell.ref.slice(col.length))
    const s = seg(col, row, p)
    const x0 = cellLeft(col) + 2, top = (row - 1) * 30
    for (let y = top + 2; y <= top + 27; y++) for (let i = 0; i < s.len; i++) put(x0 + s.from + i, y, s.rgb ?? BLUE)
  }
  return { width, height, data }
}

const judge = (png: DecodedPng) => {
  const measured = measureCaseBars(c as never, png, geom, parts)
  expect(measured instanceof Map, JSON.stringify(measured).slice(0, 400)).toBe(true)
  return evaluateBarCapabilities(c as never, parts, png, geom, measured as never)
}
const BASE: Plant = { distinct: 'evaluate', ltr: 'half', rtl: 'half', d1090: 'half-span', stat: 'half' }
const CAPS = { distinct: 'bar:all-equal-distinct-bounds', ltr: 'bar:degenerate-equal-ltr', rtl: 'bar:degenerate-equal-rtl', d1090: 'bar:degenerate-equal-10-90' } as const

afterEach(() => injectCfBarAllEqualReadings(null))

describe('stage3-bar-all-equal (pre-registered, synthetic capture)', () => {
  it('is appended last, recipe-pending, and owns its four questions; the controls stay owned elsewhere', () => {
    expect(c.status).toBe('recipe-pending')
    expect(CF_STAGE3_CASES.at(-1)!.id).toBe(ID)
    for (const cap of Object.values(CAPS)) {
      expect(CF_STAGE3_CAPABILITIES.find(e => e.capabilityId === cap)!.caseId, cap).toBe(ID)
      expect(CF_STAGE3_NONDISCRIMINATING[cap], cap).toMatch(/^Pre-registered \(stage3-bar-all-equal/)
    }
    expect(CF_STAGE3_CAPABILITIES.find(e => e.capabilityId === 'bar:degenerate-equal-values')!.caseId).toBe('stage3-bar-degenerate')
    expect(c.expected.cells.filter(x => x.matched).map(x => x.ref)).toEqual(COLS.flatMap(col => [2, 3, 4].map(r => `${col}${r}`)))
  })

  it('no predicted edge sits on a half pixel, and every pair of readings is >= 3 px apart (or bar vs none) at >= 3 cells', () => {
    const exact = (x: number) => Math.abs((x % 1) - 0.5) > 0.05
    for (const t of [0.3, 0.7, 0.5, 1, 0, 0.6, 0.9]) expect(exact(t * S), String(t)).toBe(true)
    for (const t of [0, 0.3, 0.5, 1]) expect(exact((0.05 + 0.8 * t) * S), `10/90 ${t}`).toBe(true)
    expect(exact(0.25 * S)).toBe(true)
    const apart = (a: Seg, b: Seg) => (a.len === 0) !== (b.len === 0) || (a.len > 0 && Math.max(Math.abs(a.from - b.from), Math.abs(a.from + a.len - b.from - b.len)) >= 3)
    const families: Array<[keyof Plant, readonly string[], string[]]> = [['distinct', DISTINCT, 'BCDEFGHI'.split('')], ['ltr', LTR, ['M']], ['rtl', RTL, ['N']], ['d1090', D1090, ['O']]]
    for (const [knob, readings, cols] of families) {
      for (const a of readings) for (const b of readings) {
        if (a >= b) continue
        const cells = cols.flatMap(col => [2, 3, 4].filter(r => apart(predicted(col, r, { ...BASE, [knob]: a }), predicted(col, r, { ...BASE, [knob]: b }))))
        expect(cells.length, `${knob}: ${a} vs ${b}`).toBeGreaterThanOrEqual(3)
      }
    }
  })

  for (const [knob, readings] of [['distinct', DISTINCT], ['ltr', LTR], ['rtl', RTL], ['d1090', D1090]] as const) {
    for (const reading of readings) {
      it(`planted ${knob} = ${reading}: ${CAPS[knob]} comes back as ${reading}, the one match; the controls hold`, () => {
        const ev = judge(synthetic({ ...BASE, [knob]: reading }))
        expect([ev[CAPS[knob]]!.status, ev[CAPS[knob]]!.matched], ev[CAPS[knob]]!.reason).toEqual(['discriminating', reading])
        // The statistical control M is the verified answer (half), whatever is planted elsewhere.
        expect(ev['bar:degenerate-equal-values']!.matched).toBe('half')
      }, TIMEOUT_MS)
    }
  }

  it('a no-match plant (every all-equal cell at three quarters) matches no reading: refused, never fitted', () => {
    const ev = judge(synthetic(BASE, (col, row, p) => 'BCDEFGHMNO'.includes(col) ? lin(0.75) : predicted(col, row, p)))
    for (const cap of ['bar:all-equal-distinct-bounds', 'bar:degenerate-equal-ltr', 'bar:degenerate-equal-rtl', 'bar:degenerate-equal-10-90']) expect(ev[cap]!.status, cap).toBe('none')
  }, TIMEOUT_MS)
})

/** Production's bars on one workbook, as px segments from each cell's leading inset (Lesson 1's capture-scale layout). */
async function productionBars(saved: string, refs?: readonly string[]): Promise<Map<string, string>> {
  const doc = await parseXlsx(await OfficePackage.load(new Uint8Array(readFileSync(join(REPO_ROOT, saved)))))
  const sheet = doc.sheets[0]!
  const batch = (await prepareCfPaint(sheet, { evidence: cfRuntimeEvidence }))!
  const out = new Map<string, string>()
  try {
    const widths = (sheet as unknown as { cols?: unknown }).cols
    void widths
    for (let row = 0; row < 40; row++) for (let col = 0; col < 26; col++) {
      const ref = `${String.fromCharCode(65 + col)}${row + 1}`
      if (refs && !refs.includes(ref)) continue
      const bar = batch.barAt(row, col)
      out.set(ref, bar ? `${bar.startFraction}..${bar.endFraction}${bar.direction === 'rightToLeft' ? ' rtl' : ''}${bar.isNegative ? ' neg' : ''}` : 'none')
    }
  } finally { batch.dispose() }
  return out
}

describe('rehearsal: each reading registered (reading + evidence) paints as planted, and no earlier capture moves', () => {
  const measured3 = manifest.cases.filter(x => x.stage === 3 && x.status === 'measured')
  let baseline: Map<string, Map<string, string>> | undefined
  const base = async () => {
    if (baseline) return baseline
    baseline = new Map()
    for (const m of measured3) baseline.set(m.id, await productionBars(m.saved))
    return baseline
  }

  for (const [knob, readings] of [['distinct', DISTINCT], ['ltr', LTR], ['rtl', RTL], ['d1090', D1090]] as const) {
    for (const reading of readings) {
      it(`${CAPS[knob]} = ${reading}`, async () => {
        const before = await base()
        injectCfBarAllEqualReadings({ [CAPS[knob]]: reading })
        await withExactCfEvidence([...VERIFIED_CAPABILITIES, CAPS[knob]], async () => {
          // The generated workbook: production's px at capture scale equal the planted ones (+-1 px) on the reading's cells.
          const doc = await parseXlsx(await OfficePackage.load(new Uint8Array(readFileSync(join(REPO_ROOT, c.saved)))))
          const batch = (await prepareCfPaint(doc.sheets[0]!, { evidence: cfRuntimeEvidence }))!
          const plant = { ...BASE, [knob]: reading }
          const cols = knob === 'distinct' ? 'BCDEFGHIJK' : knob === 'ltr' ? 'M' : knob === 'rtl' ? 'N' : 'O'
          try {
            for (const col of cols.split('')) for (const r of [2, 3, 4]) {
              const want = predicted(col, r, plant)
              const bar = batch.barAt(r - 1, col.charCodeAt(0) - 65)
              const ref = `${col}${r}`
              // No bar: production withholds or paints nothing (a zero-length bar lays out no fill).
              if (want.len === 0 && !bar) continue
              expect(bar, `${ref} paints`).toBeDefined()
              const host = { x: cellLeft(col) * 4 / 3, y: (r - 1) * 30 * 4 / 3, width: 195 * 4 / 3, height: 30 * 4 / 3 }
              const p = layoutDataBar(bar!, host, { zoom: 0.75, dpr: 1, output: 'device' })!
              expect(p, `${ref} lays out`).toBeDefined()
              const x0 = cellLeft(col) + 2
              if (want.len === 0) expect(p.fill, `${ref}: no fill`).toBeUndefined()
              else {
                expect(Math.abs(p.fill!.x - (x0 + want.from)), `${ref} left ${p.fill!.x} vs ${x0 + want.from}`).toBeLessThanOrEqual(1)
                expect(Math.abs(p.fill!.x + p.fill!.w - (x0 + want.from + want.len)), `${ref} right`).toBeLessThanOrEqual(1)
              }
              // The statistical twins announce no position; distinct bounds announce the real value's.
              if (knob === 'distinct') expect(bar!.valueFraction, ref).toBeDefined()
              else expect('valueFraction' in bar!, ref).toBe(false)
            }
          } finally { batch.dispose() }
          // Every earlier Stage 3 capture: production unchanged except the all-equal shapes this id names.
          for (const m of measured3) {
            const now = await productionBars(m.saved)
            const moved = [...now].filter(([ref, v]) => before.get(m.id)!.get(ref) !== v).map(([ref, v]) => `${m.id}:${ref} ${before.get(m.id)!.get(ref)} -> ${v}`)
            const expectMoved = knob === 'distinct' && reading !== 'none'
              ? ['H', 'I'].flatMap(col => [2, 3, 4, 5, 6].map(r => `stage3-bar-geometry:${col}${r}`))
              : []
            expect(moved.map(x => x.split(' ')[0]).sort(), m.id).toEqual(expectMoved.filter(k => k.startsWith(`${m.id}:`)).sort())
          }
        })
      }, TIMEOUT_MS)
    }
  }

  it('stage3-bar-geometry H and I (all 100 under num 0/100, Excel full) agree with evaluate and full, refute degenerate-half and none', async () => {
    const side = JSON.parse(readFileSync(join(REPO_ROOT, 'tests/fixtures/xlsx-cf/refs/stage3-bar-geometry.json'), 'utf8')) as { observedBars: Array<{ ref: string; cellLeftPx: number; cellWidthPx: number; bar: { leftPx: number; rightPx: number } | null }> }
    const verdict: Record<string, boolean> = {}
    for (const reading of DISTINCT) {
      injectCfBarAllEqualReadings({ 'bar:all-equal-distinct-bounds': reading })
      verdict[reading] = await withExactCfEvidence([...VERIFIED_CAPABILITIES, 'bar:all-equal-distinct-bounds'], async () => {
        const g = manifest.cases.find(x => x.id === 'stage3-bar-geometry')!
        const doc = await parseXlsx(await OfficePackage.load(new Uint8Array(readFileSync(join(REPO_ROOT, g.saved)))))
        const batch = (await prepareCfPaint(doc.sheets[0]!, { evidence: cfRuntimeEvidence }))!
        try {
          return ['H4', 'H5', 'I2', 'I3', 'I4'].every(ref => {
            const o = side.observedBars.find(x => x.ref === ref)!
            const row = Number(ref.slice(1)) - 1, col = ref.charCodeAt(0) - 65
            const bar = batch.barAt(row, col)
            if (!bar) return o.bar === null
            const p = layoutDataBar(bar, { x: o.cellLeftPx * 4 / 3, y: 0, width: o.cellWidthPx * 4 / 3, height: 40 }, { zoom: 0.75, dpr: 1, output: 'device' })
            return p?.fill !== undefined && o.bar !== null && Math.abs(p.fill.x - o.bar.leftPx) <= 1 && Math.abs(p.fill.x + p.fill.w - 1 - o.bar.rightPx) <= 1
          })
        } finally { batch.dispose() }
      })
    }
    expect(verdict).toEqual({ evaluate: true, full: true, 'degenerate-half': false, none: false })
  }, TIMEOUT_MS)
})
