// @vitest-environment node
/**
 * The percent-order follow-up (brief cf-percent-order), Phase 1 tasks 3-5: stage2-scale-percent-order (colour scales:
 * scale:percent-order, scale:percentile-order, scale:formula-bounds-v2) and stage5-icon-percent-order (icon sets:
 * icon:percent-order, icon:percentile-order), pre-registered and registrable, and the rehearsal of every world. Checked here:
 * 1. the design: ownership, registries, size, every designated cell's role, every pair of readings separated, the text scope;
 * 2. the independent model is production in every world: unregistered; each reading registered with its evidence; the
 *    pre-committed narrowing;
 * 3. the evaluator on synthetic observations: each planted reading selects exactly itself; the "neither" world refuses;
 * 4. every world against every committed capture: Stage 2 conformance has 0 violations and no painted cell changes; the
 *    Stage 5 records do not move; icon worlds leave every captured cell as it was.
 * This checks the cases and production, not Excel.
 */
import { readFileSync } from 'node:fs'
import { resolve } from 'node:path'
import { afterEach, describe, expect, it } from 'vitest'
import {
  CF_STAGE2_CAPABILITIES,
  CF_STAGE2_CASES,
  CF_STAGE2_NONDISCRIMINATING,
  CF_STAGE5_CASES,
  CF_STAGE5_NONDISCRIMINATING,
  S2_FORMULA_BOUNDS_V2_CAP,
  S2_FORMULA_BOUNDS_V2_READINGS,
  S2_PERCENT_ORDER_CAP,
  S2_PERCENT_ORDER_CELLS,
  S2_PERCENT_ORDER_DEFAULT_WORLD,
  S2_PERCENT_ORDER_ID,
  S2_PERCENT_ORDER_READINGS,
  S2_PERCENT_ORDER_TWIN,
  S2_PERCENTILE_ORDER_CAP,
  S5_ICON_PERCENT_ORDER_CAP,
  S5_ICON_PERCENTILE_ORDER_CAP,
  S5_ICON_PO_ID,
  S5_ICON_PO_READINGS,
  S5_ICON_PO_TWIN,
  s2PercentOrderCells,
  s2PercentOrderLayout,
  s2PercentOrderPrediction,
  s2PercentOrderReadingPrediction,
  s2PercentOrderReadingsOf,
  s2PercentOrderThresholds,
  s2PercentOrderWithheldToday,
  s5IconPoBlocks,
  s5IconPoBucket,
  type CfManifest,
  type CfOracleCase,
  type S2PercentOrderCell,
  type S2PercentOrderWorld,
} from '../scripts/cf-refs/cases'
import { CF_MANIFEST_PATH } from '../scripts/cf-refs/generate'
import { evaluateFormulaCapabilities, formulaControlsOf, stage2ReadingsObservationOf } from '../scripts/cf-refs/verify'
import { stage2Conformance } from './helpers/cf-stage2-conformance'
import { OfficePackage } from '../src/core/zip'
import { parseXlsx } from '../src/xlsx/parse'
import { cfDocumentEvaluationOf, prepareCfPaint } from '../src/xlsx/cf/render'
import { cfRuntimeEvidence } from '../src/xlsx/cf/predicates'
import {
  CF_FORMULA_BOUNDS_V2_CAP,
  CF_FORMULA_BOUNDS_V2_READINGS,
  CF_ICON_PERCENT_ORDER_CAP,
  CF_ICON_PERCENTILE_ORDER_CAP,
  CF_PERCENT_ORDER_READINGS,
  CF_SCALE_PERCENT_ORDER_CAP,
  CF_SCALE_PERCENTILE_ORDER_CAP,
  injectCfFormulaBoundsV2Readings,
  injectCfPercentOrderNarrowings,
  injectCfPercentOrderReadings,
  type CfFormulaBoundsV2Reading,
  type CfPercentOrderReading,
} from '../src/xlsx/cf/colorScale'
import { estimateCfPrepareBytes } from '../src/xlsx/cf/service'
import { hasEvidence, VERIFIED_CAPABILITIES, withExactCfEvidence } from '../src/xlsx/cf/evidence'
import { iconBucket, resolveIconThresholds } from '../src/xlsx/cf/iconSet'
import type { CfDefinition } from '../src/xlsx/cf/model'
import type { CfRangeStatistics } from '../src/xlsx/cf/statistics'

function statsOf(values: readonly number[]): CfRangeStatistics {
  const sorted = Float64Array.from([...values].sort((a, b) => a - b))
  return { scopeKey: 't', numericCount: values.length, sum: 0, mean: undefined, min: sorted[0], max: sorted[sorted.length - 1], stddevPop: undefined, sortedAsc: sorted, frequencies: new Map(), declaredCells: values.length, spanCells: values.length, storedCells: values.length }
}

const REPO_ROOT = resolve(__dirname, '..')
const HEAVY = 300_000
afterEach(() => { injectCfPercentOrderReadings(null); injectCfPercentOrderNarrowings(null); injectCfFormulaBoundsV2Readings(null) })

const manifest = JSON.parse(readFileSync(CF_MANIFEST_PATH, 'utf8')) as CfManifest
const caseOf = (id: string): CfOracleCase => {
  const c = manifest.cases.find(x => x.id === id)
  if (!c) throw new Error(`${id} is not in the manifest (generate it with generate.ts --case ${id})`)
  return c
}
const colOf = (ref: string): string => /^[A-Z]+/.exec(ref)![0]
const rowOf = (ref: string): number => Number(/\d+$/.exec(ref)![0])
const rc = (ref: string): { row: number; col: number } => ({ row: rowOf(ref) - 1, col: [...colOf(ref)].reduce((n, ch) => n * 26 + ch.charCodeAt(0) - 64, 0) - 1 })
const layout = s2PercentOrderLayout()
const twinFill = new Map(layout.twins.map(t => [t.ref, t.fill]))
const designated = s2PercentOrderCells()
const S2_CAPS = [S2_PERCENT_ORDER_CAP, S2_PERCENTILE_ORDER_CAP, S2_FORMULA_BOUNDS_V2_CAP] as const

/** A registration world: readings by id (with their evidence held), and the narrowed ids. */
type Registration = { readings?: Partial<Record<string, CfPercentOrderReading>>; v2?: CfFormulaBoundsV2Reading; narrowed?: readonly string[] }
async function withRegistration<T>(r: Registration, fn: () => Promise<T>): Promise<T> {
  injectCfPercentOrderReadings(r.readings ?? null)
  injectCfFormulaBoundsV2Readings(r.v2 ? { [CF_FORMULA_BOUNDS_V2_CAP]: r.v2 } : null)
  injectCfPercentOrderNarrowings(r.narrowed ?? null)
  const caps = [...VERIFIED_CAPABILITIES, ...Object.keys(r.readings ?? {}), ...(r.v2 ? [CF_FORMULA_BOUNDS_V2_CAP] : [])]
  try { return await withExactCfEvidence(caps, fn) } finally { injectCfPercentOrderReadings(null); injectCfFormulaBoundsV2Readings(null); injectCfPercentOrderNarrowings(null) }
}
/** The model world a registration selects (narrowing aside). */
const worldOf = (r: Registration): S2PercentOrderWorld => ({
  percent: r.readings?.[CF_SCALE_PERCENT_ORDER_CAP] ?? 'lerp-first',
  percentile: r.readings?.[CF_SCALE_PERCENTILE_ORDER_CAP] ?? 'lerp-first',
  formula: r.v2 ? (r.v2.split('+')[0] as S2PercentOrderWorld['formula']) : 'as-literal',
})
/** Whether a v2 registration lifts the formula-bound shapes in a registration (its order is the applied percent order). */
const v2Lifts = (r: Registration): boolean => r.v2 !== undefined && !(r.narrowed ?? []).includes(CF_SCALE_PERCENT_ORDER_CAP) && r.v2.split('+')[1] === worldOf(r).percent

async function sheetOf(c: CfOracleCase) {
  const doc = await parseXlsx(await OfficePackage.load(new Uint8Array(readFileSync(resolve(REPO_ROOT, c.saved)))))
  return { doc, sheet: doc.sheets.find(s => s.name === c.sheet) ?? doc.sheets[0]! }
}
async function productionFills(c: CfOracleCase, refs: readonly string[]): Promise<Map<string, string | undefined>> {
  const { doc, sheet } = await sheetOf(c)
  const paint = await prepareCfPaint(sheet, { evidence: cfRuntimeEvidence, evaluation: cfDocumentEvaluationOf(doc) })
  const out = new Map<string, string | undefined>()
  for (const ref of refs) { const { row, col } = rc(ref); out.set(ref, paint?.layer.styleAt(row, col, undefined)?.fillColor?.toUpperCase()) }
  paint?.dispose()
  return out
}
/** The model's expectation at a cell in a registration world: the world's prediction, or withheld (undefined). */
function expected(cell: S2PercentOrderCell, r: Registration): string | undefined {
  const narrowedHere = (cap: string): boolean => (r.narrowed ?? []).includes(cap)
  // The narrowing: a rule with a narrowed bound type whose two orders differ withholds every cell.
  const differs = (kind: 'percent' | 'percentile'): boolean => {
    const base = { ...S2_PERCENT_ORDER_DEFAULT_WORLD, formula: worldOf(r).formula }
    const [a, b] = (['lerp-first', 'multiply-first'] as const).map(o => s2PercentOrderThresholds(cell.block, { ...base, [kind]: o }))
    return a.some((x, i) => x !== b[i])
  }
  if (narrowedHere(CF_SCALE_PERCENT_ORDER_CAP) && cell.block.cfvo.some(x => x.type === 'percent') && differs('percent')) return undefined
  if (narrowedHere(CF_SCALE_PERCENTILE_ORDER_CAP) && cell.block.cfvo.some(x => x.type === 'percentile') && differs('percentile')) return undefined
  const w = worldOf(r)
  // A formula-bound cell in a settled shape at this world's thresholds, unlifted: withheld.
  if (s2PercentOrderWithheldToday(cell, w) && !v2Lifts(r)) return undefined
  return s2PercentOrderPrediction(cell, w)
}

/** A synthetic Stage 2 observation: designated cells by `fill`, twins and positive controls as fixed, controls white. */
function plantedFills(c: CfOracleCase, fill: (cell: S2PercentOrderCell) => string): Map<string, string> {
  const out = new Map<string, string>()
  for (const e of c.expected.cells) out.set(e.ref, e.matched ? e.fillRgb ?? '' : 'FFFFFFFF')
  for (const [ref, f] of twinFill) out.set(ref, f)
  for (const cell of designated) out.set(cell.ref, fill(cell))
  return out
}
const judge = (c: CfOracleCase, fills: Map<string, string>) => evaluateFormulaCapabilities(c.scaleReadings!, formulaControlsOf(c), stage2ReadingsObservationOf({
  observedCells: c.expected.cells.filter(e => fills.get(e.ref) !== undefined).map(e => ({ ref: e.ref, controlRef: e.controlRef, fillSource: 'excel-display-format', fillRgb: fills.get(e.ref) })),
}))

describe('stage2-scale-percent-order: the case', () => {
  it('is appended last; owns its four ids; every question unregistered and in the registry', async () => {
    expect(CF_STAGE2_CASES.at(-1)!.id).toBe(S2_PERCENT_ORDER_ID)
    for (const id of [...S2_CAPS, S2_PERCENT_ORDER_TWIN]) {
      expect(CF_STAGE2_CAPABILITIES.filter(e => e.capabilityId === id).map(e => e.caseId), id).toEqual([S2_PERCENT_ORDER_ID])
      expect(CF_STAGE2_NONDISCRIMINATING[id], id).toMatch(/./)
      expect(VERIFIED_CAPABILITIES.has(id), id).toBe(false)
    }
    expect([S2_PERCENT_ORDER_CAP, S2_PERCENTILE_ORDER_CAP, S2_FORMULA_BOUNDS_V2_CAP]).toEqual([CF_SCALE_PERCENT_ORDER_CAP, CF_SCALE_PERCENTILE_ORDER_CAP, CF_FORMULA_BOUNDS_V2_CAP])
    expect([...S2_PERCENT_ORDER_READINGS]).toEqual([...CF_PERCENT_ORDER_READINGS])
    expect([...S2_FORMULA_BOUNDS_V2_READINGS]).toEqual([...CF_FORMULA_BOUNDS_V2_READINGS])
    const { CF_STAGE2_PRODUCTION_READINGS } = await import('../scripts/cf-refs/audit')
    for (const id of S2_CAPS) expect(Object.prototype.hasOwnProperty.call(CF_STAGE2_PRODUCTION_READINGS, id), id).toBe(false)
    const c = caseOf(S2_PERCENT_ORDER_ID)
    expect(c.status).toBe('recipe-pending')
    expect(c.scaleReadings).toEqual(CF_STAGE2_CASES.find(t => t.id === S2_PERCENT_ORDER_ID)!.scaleReadings)
  })

  it('225 cells in A1:W14: 103 designated (63 decisive, 3 kept, 13 clear, 24 controls), 106 twins, 2 positive controls, 14 unformatted controls; 22 colour rules', () => {
    const c = caseOf(S2_PERCENT_ORDER_ID)
    expect(c.range).toBe('A1:W14')
    const count = (role: string): number => designated.filter(x => x.role === role).length
    expect([designated.length, count('decisive'), count('kept'), count('clear'), count('control')]).toEqual([103, 63, 3, 13, 24])
    expect(layout.twins.length).toBe(106)
    expect(c.expected.cells.length).toBe(225)
    expect(c.markers.length).toBe(22)
    expect(c.expected.cells.length).toBeLessThan(300)
  })

  it('every group has >= 3 decisive cells (5 or 8); every pair of a capability\'s readings is told apart; the formula cells hold their doubles', () => {
    for (const b of new Set(designated.map(x => x.block))) expect(designated.filter(x => x.block === b && x.role === 'decisive').length, b.id).toBeGreaterThanOrEqual(5)
    for (const cap of S2_CAPS) {
      const rs = s2PercentOrderReadingsOf(cap)
      for (let i = 0; i < rs.length; i++) for (let j = i + 1; j < rs.length; j++) {
        expect(designated.filter(x => x.block.cap === cap && s2PercentOrderReadingPrediction(x, rs[i]!) !== s2PercentOrderReadingPrediction(x, rs[j]!)).length, `${rs[i]} / ${rs[j]}`).toBeGreaterThanOrEqual(3)
      }
    }
    for (const [ref, text] of Object.entries(S2_PERCENT_ORDER_CELLS)) expect(String(Number(text)), ref).toBe(text)
    // Both signs of the threshold difference, for percent and for percentile.
    const sign = (id: string, kind: 'percent' | 'percentile'): number => {
      const b = designated.find(x => x.block.id === id)!.block
      const [l, m] = (['lerp-first', 'multiply-first'] as const).map(o => s2PercentOrderThresholds(b, { ...S2_PERCENT_ORDER_DEFAULT_WORLD, [kind]: o }))
      return Math.sign(l!.find((x, i) => x !== m![i])! - m!.find((x, i) => x !== l![i])!)
    }
    expect(['po-lo-pq', 'po-lo-qp', 'po-hi-qp', 'po-hi-pq'].map(id => sign(id, 'percent'))).toEqual([1, -1, -1, 1])
    expect(['pl-pq', 'pl-qp'].map(id => sign(id, 'percentile'))).toEqual([1, -1])
  })

  it('the text scope admits the whole range in one batch, well under the 8,388,608-byte cap', async () => {
    const { sheet } = await sheetOf(caseOf(S2_PERCENT_ORDER_ID))
    const bytes = estimateCfPrepareBytes(sheet.conditionalFormatting!, [{ firstRow: 0, lastRow: 13, firstCol: 0, lastCol: 22 }])
    expect(bytes!).toBeLessThan(8388608 * 0.6)
  })
})

/** Every registration world rehearsed (each reading of each id, the v2 readings at each order, the narrowings). */
const WORLDS: ReadonlyArray<[string, Registration]> = [
  ['unregistered', {}],
  ['percent multiply-first', { readings: { [CF_SCALE_PERCENT_ORDER_CAP]: 'multiply-first' } }],
  ['percent lerp-first', { readings: { [CF_SCALE_PERCENT_ORDER_CAP]: 'lerp-first' } }],
  ['percentile multiply-first', { readings: { [CF_SCALE_PERCENTILE_ORDER_CAP]: 'multiply-first' } }],
  ['percentile lerp-first', { readings: { [CF_SCALE_PERCENTILE_ORDER_CAP]: 'lerp-first' } }],
  ['v2 as-literal+lerp-first', { v2: 'as-literal+lerp-first' }],
  ['v2 as-literal+multiply-first with percent multiply-first', { readings: { [CF_SCALE_PERCENT_ORDER_CAP]: 'multiply-first' }, v2: 'as-literal+multiply-first' }],
  ['v2 as-literal+multiply-first alone (order unregistered: no lift)', { v2: 'as-literal+multiply-first' }],
  ['v2 fifteen-digit+lerp-first', { v2: 'fifteen-digit+lerp-first' }],
  ['v2 fifteen-digit+multiply-first with percent multiply-first', { readings: { [CF_SCALE_PERCENT_ORDER_CAP]: 'multiply-first' }, v2: 'fifteen-digit+multiply-first' }],
  ['narrowed percent', { narrowed: [CF_SCALE_PERCENT_ORDER_CAP] }],
  ['narrowed percentile', { narrowed: [CF_SCALE_PERCENTILE_ORDER_CAP] }],
]

describe('stage2-scale-percent-order: the model is production in every world', () => {
  for (const [name, r] of WORLDS) {
    it(name, async () => {
      const got = await withRegistration(r, () => productionFills(caseOf(S2_PERCENT_ORDER_ID), designated.map(x => x.ref)))
      const off = designated.filter(x => got.get(x.ref) !== expected(x, r)).map(x => `${x.ref}: ${got.get(x.ref)} vs ${expected(x, r)}`)
      expect(off).toEqual([])
    }, HEAVY)
  }
  it('unregistered: exactly the formula-bound cells of the settled shapes are withheld (fp-33, fp-28, fn-hi, f3-lower, f3-upper); every literal cell paints', () => {
    const withheld = designated.filter(x => s2PercentOrderWithheldToday(x))
    expect([...new Set(withheld.map(x => x.block.id))]).toEqual(['fp-33', 'fp-28', 'fn-hi', 'f3-lower', 'f3-upper'])
    expect(withheld.length).toBe(28)
    expect(withheld.every(x => x.block.cap === S2_FORMULA_BOUNDS_V2_CAP)).toBe(true)
    // fp-47's decisive cells are painted today (as-literal+lerp-first is not at a guard boundary there).
    expect(designated.filter(x => x.block.id === 'fp-47' && x.role === 'decisive' && !s2PercentOrderWithheldToday(x)).length).toBe(8)
  })
})

describe('stage2-scale-percent-order: the evaluator', () => {
  const c = caseOf(S2_PERCENT_ORDER_ID)
  it('each planted reading selects exactly itself, for every capability', () => {
    for (const cap of S2_CAPS) for (const r of s2PercentOrderReadingsOf(cap)) {
      const fills = plantedFills(c, cell => (cell.block.cap === cap ? s2PercentOrderReadingPrediction(cell, r) : s2PercentOrderPrediction(cell, S2_PERCENT_ORDER_DEFAULT_WORLD)))
      const out = judge(c, fills)
      expect(out[cap]?.status, `${cap} ${r}: ${out[cap]?.reason}`).toBe('discriminating')
      expect(out[cap]?.matched).toBe(r)
    }
  })
  it('the neither world (one decisive cell per group off both readings: its rival\'s fill one step further) refuses', () => {
    for (const cap of S2_CAPS) {
      const rs = s2PercentOrderReadingsOf(cap)
      const flipped = new Set([...new Set(designated.filter(x => x.block.cap === cap).map(x => x.block.id))].map(id => designated.find(x => x.block.id === id && x.role === 'decisive')!.ref))
      const fills = plantedFills(c, cell => {
        const a = s2PercentOrderReadingPrediction(cell, rs[0]!)
        if (cell.block.cap !== cap || !flipped.has(cell.ref)) return a
        // A fill no reading predicts: the first reading's fill two P-line steps further (still a twin-line fill? no: a literal mismatch).
        return a === 'FF14F028' ? 'FF16EE2A' : `FF${(parseInt(a.slice(2, 4), 16) + 2).toString(16).toUpperCase().padStart(2, '0')}${a.slice(4)}`
      })
      const out = judge(c, fills)
      expect(out[cap]?.status, cap).toBe('refused')
    }
  })
})

/**
 * The committed Stage 2 captures under every world. The worlds consistent with stage2-scale-formula-bounds (d2660d46) give 0
 * violations; the v2 worlds that capture contradicts are caught exactly there (E2..E6 for as-literal+lerp-first; the
 * as-literal-matched cells for the fifteen-digit worlds), so a match of one of them would be a contradiction between
 * captures, never a registration (recipe §15.2). Cells that change from withheld to painted are listed by count.
 */
describe('every world against every committed capture', () => {
  const measured = manifest.cases.filter(x => x.stage === 2 && x.status === 'measured').map(x => x.id)
  const E = ['E2', 'E3', 'E4', 'E5', 'E6'].map(r => `stage2-scale-formula-bounds:${r}`)
  const expectations: Record<string, { violations: (v: string[]) => void; newlyPainted: number }> = {
    'unregistered': { violations: v => expect(v).toEqual([]), newlyPainted: 0 },
    'percent multiply-first': { violations: v => expect(v).toEqual([]), newlyPainted: 5 },
    'percent lerp-first': { violations: v => expect(v).toEqual([]), newlyPainted: 0 },
    'percentile multiply-first': { violations: v => expect(v).toEqual([]), newlyPainted: 0 },
    'percentile lerp-first': { violations: v => expect(v).toEqual([]), newlyPainted: 0 },
    'v2 as-literal+lerp-first': { violations: v => expect(v.map(x => x.slice(0, x.indexOf(':', 'stage2-scale-formula-bounds:'.length)))).toEqual(E), newlyPainted: -1 },
    'v2 as-literal+multiply-first with percent multiply-first': { violations: v => expect(v).toEqual([]), newlyPainted: 47 },
    'v2 as-literal+multiply-first alone (order unregistered: no lift)': { violations: v => expect(v).toEqual([]), newlyPainted: 0 },
    'v2 fifteen-digit+lerp-first': { violations: v => { expect(v.length).toBeGreaterThan(0); expect(v.every(x => x.startsWith('stage2-scale-formula-bounds:'))).toBe(true) }, newlyPainted: -1 },
    'v2 fifteen-digit+multiply-first with percent multiply-first': { violations: v => { expect(v.length).toBeGreaterThan(0); expect(v.every(x => x.startsWith('stage2-scale-formula-bounds:'))).toBe(true) }, newlyPainted: -1 },
    'narrowed percent': { violations: v => expect(v).toEqual([]), newlyPainted: -2 },
    'narrowed percentile': { violations: v => expect(v).toEqual([]), newlyPainted: 0 },
  }
  for (const [name, r] of WORLDS) {
    it(`Stage 2 conformance (${name})`, async () => {
      const base = await stage2Conformance(manifest, REPO_ROOT, measured)
      const got = await withRegistration(r, () => stage2Conformance(manifest, REPO_ROOT, measured))
      const want = expectations[name]!
      want.violations(got.violations)
      expect(got.total).toBe(base.total)
      if (want.newlyPainted >= 0) expect(base.withheld - got.withheld, name).toBe(want.newlyPainted)
      // The narrowing withholds the whole E2:E10 rule of stage2-scale-formula-bounds (its percent thresholds differ): E7..E10,
      // painted today, become withheld, justified (E2..E6 were withheld already).
      if (want.newlyPainted === -2) expect(got.withheld - base.withheld).toBe(4)
    }, HEAVY)
  }
})

describe('every world against the committed Stage 5 captures', () => {
  it('Stage 5 records: every measured case agrees with Excel in every world; no record changes, except under the scale percentile narrowing (stage5-cfvo-formulas G / T withheld, justified)', async () => {
    const { stage5CaseRecords } = await import('../scripts/cf-refs/audit')
    const base = await stage5CaseRecords(manifest, REPO_ROOT)
    expect(base.filter(r => r.compared && !r.agrees)).toEqual([])
    const k = (r: (typeof base)[number]): string => `${r.caseId}:${r.ref}`
    const was = new Map(base.map(r => [k(r), r]))
    const worlds: ReadonlyArray<[string, Registration]> = [
      ['icon percent multiply-first', { readings: { [CF_ICON_PERCENT_ORDER_CAP]: 'multiply-first' } }],
      ['icon percentile multiply-first', { readings: { [CF_ICON_PERCENTILE_ORDER_CAP]: 'multiply-first' } }],
      ['icon narrowed', { narrowed: [CF_ICON_PERCENT_ORDER_CAP, CF_ICON_PERCENTILE_ORDER_CAP] }],
      ['scale multiply-first', { readings: { [CF_SCALE_PERCENT_ORDER_CAP]: 'multiply-first', [CF_SCALE_PERCENTILE_ORDER_CAP]: 'multiply-first' } }],
      ['scale percentile narrowed', { narrowed: [CF_SCALE_PERCENTILE_ORDER_CAP] }],
    ]
    for (const [name, r] of worlds) {
      const held = new Set([...Object.keys(r.readings ?? {})])
      const rows = await withRegistration(r, () => stage5CaseRecords(manifest, REPO_ROOT, { published: id => VERIFIED_CAPABILITIES.has(id) || held.has(id) }))
      expect(rows.length, name).toBe(base.length)
      expect(rows.filter(x => x.compared && !x.agrees).map(x => `${k(x)} ${x.disagreement}`), name).toEqual([])
      const changed = rows.filter(x => JSON.stringify(x.production) !== JSON.stringify(was.get(k(x))!.production)).map(k)
      if (name === 'scale percentile narrowed') expect(changed).toEqual([...['G2', 'G3', 'G4', 'G5', 'G6', 'G7', 'G8', 'G9'], ...['T2', 'T3', 'T4', 'T5', 'T6', 'T7', 'T8', 'T9']].map(x => `stage5-cfvo-formulas:${x}`))
      else expect(changed, name).toEqual([])
    }
  }, 900_000)
})

describe('stage5-icon-percent-order: the case', () => {
  const blocks = s5IconPoBlocks()
  it('is appended last; owns its three ids; both questions unregistered and in the registry', async () => {
    expect(CF_STAGE5_CASES.at(-1)!.id).toBe(S5_ICON_PO_ID)
    for (const id of [S5_ICON_PERCENT_ORDER_CAP, S5_ICON_PERCENTILE_ORDER_CAP, S5_ICON_PO_TWIN]) {
      expect(CF_STAGE5_NONDISCRIMINATING[id], id).toMatch(/./)
      expect(VERIFIED_CAPABILITIES.has(id), id).toBe(false)
    }
    expect([S5_ICON_PERCENT_ORDER_CAP, S5_ICON_PERCENTILE_ORDER_CAP]).toEqual([CF_ICON_PERCENT_ORDER_CAP, CF_ICON_PERCENTILE_ORDER_CAP])
    expect([...S5_ICON_PO_READINGS]).toEqual([...CF_PERCENT_ORDER_READINGS])
    const { CF_STAGE5_PRODUCTION_READINGS } = await import('../scripts/cf-refs/audit')
    for (const id of [S5_ICON_PERCENT_ORDER_CAP, S5_ICON_PERCENTILE_ORDER_CAP]) expect(Object.prototype.hasOwnProperty.call(CF_STAGE5_PRODUCTION_READINGS, id), id).toBe(false)
    const c = caseOf(S5_ICON_PO_ID)
    expect(c.range).toBe('A1:AK27')
    expect(c.expected.cells.length).toBe(134)
    expect(c.expected.cells.length).toBeLessThan(300)
  })
  it('decisive cells hold the smaller threshold, one bucket apart; both signs for percent, one for percentile; controls agree', () => {
    const decisive = (id: string): Array<[number, number]> => { const b = blocks.find(x => x.id === id)!; return b.designated.filter(d => d.role === 'decisive').map(d => [s5IconPoBucket(b, b.values[d.at]!, 'lerp-first'), s5IconPoBucket(b, b.values[d.at]!, 'multiply-first')]) }
    expect(decisive('ipc-up')).toEqual([[0, 1], [0, 1], [1, 2], [1, 2]])
    expect(decisive('ipc-down')).toEqual([[1, 0], [1, 0], [2, 1], [2, 1]])
    expect(decisive('ipl-up')).toEqual(Array(5).fill([0, 1]))
    expect(decisive('ipc-tol')).toEqual([])
    for (const b of blocks) for (const v of b.values) expect(String(Number(String(v))), b.id).toBe(String(v))
  })
  it('the text scope is well under the cap', async () => {
    const c = caseOf(S5_ICON_PO_ID)
    const { sheet } = await sheetOf(c)
    const bytes = estimateCfPrepareBytes(sheet.conditionalFormatting!, [{ firstRow: 0, lastRow: 26, firstCol: 0, lastCol: 36 }])
    expect(bytes!).toBe(5958496)
    expect(bytes!).toBeLessThan(8388608 * 0.75)
  })
  it('production thresholds and buckets equal the model in every world (the workbook\'s own rules; icons compared on the bucket)', async () => {
    const c = caseOf(S5_ICON_PO_ID)
    const { sheet } = await sheetOf(c)
    const rules = sheet.conditionalFormatting!.rules.filter(x => x.definition.kind === 'iconSet').slice(0, blocks.length)
    const worlds: ReadonlyArray<[Registration, (b: typeof blocks[number]) => 'lerp-first' | 'multiply-first' | 'withheld']> = [
      [{}, () => 'lerp-first'],
      [{ readings: { [CF_ICON_PERCENT_ORDER_CAP]: 'multiply-first' } }, b => (b.type === 'percent' ? 'multiply-first' : 'lerp-first')],
      [{ readings: { [CF_ICON_PERCENTILE_ORDER_CAP]: 'multiply-first' } }, b => (b.type === 'percentile' ? 'multiply-first' : 'lerp-first')],
      [{ readings: { [CF_SCALE_PERCENT_ORDER_CAP]: 'multiply-first', [CF_SCALE_PERCENTILE_ORDER_CAP]: 'multiply-first' } }, () => 'lerp-first'],
      [{ narrowed: [CF_ICON_PERCENT_ORDER_CAP] }, b => (b.type === 'percent' && b.id !== 'ipc-tol' ? 'withheld' : 'lerp-first')],
      [{ narrowed: [CF_ICON_PERCENTILE_ORDER_CAP] }, b => (b.type === 'percentile' ? 'withheld' : 'lerp-first')],
    ]
    for (const [r, want] of worlds) {
      await withRegistration(r, async () => {
        blocks.forEach((b, i) => {
          const def = rules[i]!.definition as Extract<CfDefinition, { kind: 'iconSet' }>
          const t = resolveIconThresholds(def, statsOf(b.values), { hasEvidence: cap => hasEvidence(cap) })
          const w = want(b)
          if (w === 'withheld') { expect(t, `${b.id} ${JSON.stringify(r)}`).toEqual({ kind: 'unresolved', reason: 'cf-unmeasured-percent-order' }); return }
          expect(t.kind, b.id).toBe('ready')
          const th = (t as { value: readonly number[] }).value
          for (const d of b.designated) expect(iconBucket(b.values[d.at]!, th, def.thresholds.map(x => x.gte)), `${b.id} ${d.at} ${JSON.stringify(r)}`).toBe(s5IconPoBucket(b, b.values[d.at]!, w))
        })
      })
    }
  })
})
