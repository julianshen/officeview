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
import { copyFileSync, existsSync, mkdirSync, mkdtempSync, readFileSync, writeFileSync } from 'node:fs'
import { tmpdir } from 'node:os'
import { dirname, join, resolve } from 'node:path'
import { createHash } from 'node:crypto'
import { createCanvas, loadImage } from 'canvas'
import JSZip from 'jszip'
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
import { computeWorksheetRangeGeometry, decodePngSync, evaluateFormulaCase, stage5IconKeys, evaluateFormulaCapabilities, formulaControlsOf, readZipPartsSync, stage2ReadingsObservationOf, verifyCfCase } from '../scripts/cf-refs/verify'
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
import { hasEvidence, VERIFIED_CAPABILITIES, VERIFIED_STAGE2_CAPABILITIES, withExactCfEvidence } from '../src/xlsx/cf/evidence'
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
    expect(bytes).toBe(3428094)
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

// ---- a synthetic measured capture of stage2-scale-percent-order through the real verifier, audit, conformance and cf-compare ----
const sha = (b: Uint8Array): string => createHash('sha256').update(b).digest('hex')
function syntheticRoot(fills: Map<string, string>): { root: string; manifest: CfManifest } {
  const root = mkdtempSync(join(tmpdir(), 'cf-s2-percent-order-'))
  const copy = (rel: string): void => { if (!existsSync(resolve(REPO_ROOT, rel))) return; mkdirSync(dirname(join(root, rel)), { recursive: true }); copyFileSync(resolve(REPO_ROOT, rel), join(root, rel)) }
  for (const x of manifest.cases.filter(k => k.stage === 2)) { copy(x.input); copy(x.saved); for (const cap of x.captures) { copy(cap.pngPath); copy(cap.sidecarPath) } }
  const c = caseOf(S2_PERCENT_ORDER_ID)
  const saved = new Uint8Array(readFileSync(resolve(REPO_ROOT, c.saved)))
  const input = new Uint8Array(readFileSync(resolve(REPO_ROOT, c.input)))
  const geom = computeWorksheetRangeGeometry(readZipPartsSync(saved)!, c.range)!
  const width = Math.round(geom.expectedWidthPt) + 1, height = Math.round(geom.expectedHeightPt) + 1
  const canvas = createCanvas(width, height)
  const ctx = canvas.getContext('2d')
  ctx.fillStyle = '#FFFFFF'
  ctx.fillRect(0, 0, width, height)
  for (const e of c.expected.cells) {
    const { row, col } = rc(e.ref)
    const ci = col - geom.minCol, ri = row - geom.minRow
    const x0 = Math.round(geom.colWidthsPt.slice(0, ci).reduce((a, b) => a + b, 0)), x1 = Math.round(geom.colWidthsPt.slice(0, ci + 1).reduce((a, b) => a + b, 0))
    const y0 = Math.round(geom.rowHeightsPt.slice(0, ri).reduce((a, b) => a + b, 0)), y1 = Math.round(geom.rowHeightsPt.slice(0, ri + 1).reduce((a, b) => a + b, 0))
    ctx.fillStyle = `#${fills.get(e.ref)!.slice(2)}`
    ctx.fillRect(x0 + 1, y0 + 1, x1 - x0 - 1, y1 - y0 - 1)
  }
  const png = new Uint8Array(canvas.toBuffer('image/png'))
  const cap = c.captures[0]!
  mkdirSync(dirname(join(root, cap.pngPath)), { recursive: true })
  writeFileSync(join(root, cap.pngPath), png)
  const at = '2026-11-26T00:00:00.000Z'
  writeFileSync(join(root, cap.sidecarPath), JSON.stringify({
    excelVersion: '16.115', route: 'copy-picture', appearance: 'screen', scale: 1, pxPerPt: 1, width, height,
    colWidthsPt: geom.colWidthsPt, rowHeightsPt: geom.rowHeightsPt, geometrySource: 'excel-measured', resavedByExcel: true, resaveSavedSha256: sha(saved),
    sheet: c.sheet, range: c.range, hostTimeZone: 'Asia/Taipei', savedSha256: sha(saved), pngSha256: sha(png), capturedAt: at,
    observedCells: c.expected.cells.map(e => ({ ref: e.ref, controlRef: e.controlRef, fillSource: 'excel-display-format', fillRgb: fills.get(e.ref) })),
  }, null, 2))
  const measured: CfOracleCase = {
    ...c, status: 'measured', reason: undefined, resavedByExcel: true,
    resaveReceipt: { openedRepairFree: true, excelVersion: '16.115', resavedAt: at, inputSha256: sha(input), savedSha256: sha(saved) } as CfOracleCase['resaveReceipt'],
    captures: [{ ...cap, pngSha256: sha(png) }],
    expected: { ...c.expected, provenance: 'excel-observed', cells: c.expected.cells.map(e => (e.matched ? { ...e, fillSource: 'excel-display-format' as const } : e)) },
  }
  return { root, manifest: { ...manifest, cases: manifest.cases.map(x => (x.id === S2_PERCENT_ORDER_ID ? measured : x)) } }
}

describe('stage2-scale-percent-order on a synthetic measured capture (real verifyCfCase, auditStage2Manifest, stage2Conformance, compareCfCase)', () => {
  const disposition = (r: { perCapability: ReadonlyArray<{ capabilityId: string; disposition: string }> }, id: string) => r.perCapability.find(p => p.capabilityId === id)!.disposition
  const removed = Object.fromEntries(Object.entries(CF_STAGE2_NONDISCRIMINATING).filter(([id]) => !(S2_CAPS as readonly string[]).includes(id)))
  const MULT: S2PercentOrderWorld = { percent: 'multiply-first', percentile: 'multiply-first', formula: 'as-literal' }
  const multRegistration: Registration = { readings: { [CF_SCALE_PERCENT_ORDER_CAP]: 'multiply-first', [CF_SCALE_PERCENTILE_ORDER_CAP]: 'multiply-first' }, v2: 'as-literal+multiply-first' }
  const ID = S2_PERCENT_ORDER_ID

  it('the real manifest: Stage 2 audits ok; the four ids unverified-nondiscriminating; the verified set unchanged', async () => {
    const { auditStage2Manifest } = await import('../scripts/cf-refs/audit')
    const r = auditStage2Manifest(manifest, { rootDir: REPO_ROOT })
    expect([r.ok, r.stageComplete]).toEqual([true, true])
    for (const id of [...S2_CAPS, S2_PERCENT_ORDER_TWIN]) expect(disposition(r, id), id).toBe('unverified-nondiscriminating')
    expect(r.perCapability.filter(p => p.disposition === 'verified').map(p => p.capabilityId).sort()).toEqual([...VERIFIED_STAGE2_CAPABILITIES].sort())
  }, HEAVY)

  it('a planted multiply-first world (as-literal for formulas): each id matches it; registered it verifies; unregistered production is caught exactly at the painted-today decisive cells; registered, 0 violations and cf-compare passes', async () => {
    const { auditStage2Manifest, CF_STAGE2_PRODUCTION_READINGS } = await import('../scripts/cf-refs/audit')
    const { compareCfCase } = await import('../scripts/cf-compare')
    const { root, manifest: m } = syntheticRoot(plantedFills(caseOf(ID), cell => s2PercentOrderPrediction(cell, MULT)))
    const c = m.cases.find(x => x.id === ID)!
    const res = verifyCfCase(c, { rootDir: root })
    expect(res.kind, res.kind === 'refusal' ? `${res.code}: ${res.reason}` : '').toBe('verified')
    if (res.kind !== 'verified') return
    expect(res.matchedHypotheses).toEqual({ [S2_PERCENT_ORDER_CAP]: 'multiply-first', [S2_PERCENTILE_ORDER_CAP]: 'multiply-first', [S2_FORMULA_BOUNDS_V2_CAP]: 'as-literal+multiply-first' })
    for (const id of S2_CAPS) expect(disposition(auditStage2Manifest(m, { rootDir: root }), id), id).toBe('unverified-nondiscriminating')
    const registered = auditStage2Manifest(m, { rootDir: root, nonDiscriminating: removed, productionReadings: { ...CF_STAGE2_PRODUCTION_READINGS, [S2_PERCENT_ORDER_CAP]: 'multiply-first', [S2_PERCENTILE_ORDER_CAP]: 'multiply-first', [S2_FORMULA_BOUNDS_V2_CAP]: 'as-literal+multiply-first' } })
    for (const id of S2_CAPS) expect(disposition(registered, id), id).toBe('verified')
    expect(disposition(auditStage2Manifest(m, { rootDir: root, nonDiscriminating: removed }), S2_PERCENT_ORDER_CAP)).toBe('supported-structural')
    // Unregistered: the decisive cells production paints today are lerp-first's.
    const paintedDecisive = designated.filter(x => x.role === 'decisive' && !s2PercentOrderWithheldToday(x) && s2PercentOrderPrediction(x, S2_PERCENT_ORDER_DEFAULT_WORLD) !== s2PercentOrderPrediction(x, MULT)).map(x => x.ref)
    expect(paintedDecisive).toEqual([...['A', 'B', 'C', 'D', 'E', 'F'].flatMap(col => [2, 3, 4, 5, 6].map(r => `${col}${r}`)), 'H7', 'H8', 'H9'])
    const before = await stage2Conformance(m, root, [ID])
    expect(before.violations.map(v => v.split(':')[1]).sort()).toEqual([...paintedDecisive].sort())
    const after = await withRegistration(multRegistration, () => stage2Conformance(m, root, [ID]))
    expect(after.violations).toEqual([])
    expect([after.total, after.withheld]).toEqual([before.total, 0])
    for (const mode of ['interactive', 'paged'] as const) {
      const b = await withRegistration(multRegistration, () => compareCfCase(c, mode, { rootDir: root }))
      expect(`${mode}: ${b.outcome} ${b.outcome === 'pass' ? '' : b.reason}`).toBe(`${mode}: pass `)
    }
  }, 900_000)

  it('a planted lerp-first world: each id matches lerp-first (as-literal+lerp-first); unregistered production agrees (the formula-bound shapes withheld, justified) and cf-compare passes', async () => {
    const { compareCfCase } = await import('../scripts/cf-compare')
    const { root, manifest: m } = syntheticRoot(plantedFills(caseOf(ID), cell => s2PercentOrderPrediction(cell, S2_PERCENT_ORDER_DEFAULT_WORLD)))
    const c = m.cases.find(x => x.id === ID)!
    const res = verifyCfCase(c, { rootDir: root })
    expect(res.kind).toBe('verified')
    if (res.kind !== 'verified') return
    expect(res.matchedHypotheses).toEqual({ [S2_PERCENT_ORDER_CAP]: 'lerp-first', [S2_PERCENTILE_ORDER_CAP]: 'lerp-first', [S2_FORMULA_BOUNDS_V2_CAP]: 'as-literal+lerp-first' })
    const r = await stage2Conformance(m, root, [ID])
    expect(r.violations).toEqual([])
    expect(r.withheld).toBe(28)
    for (const mode of ['interactive', 'paged'] as const) {
      const b = await compareCfCase(c, mode, { rootDir: root })
      expect(`${mode}: ${b.outcome} ${b.outcome === 'pass' ? '' : b.reason}`).toBe(`${mode}: pass `)
    }
  }, 900_000)

  it('a planted "neither" world for scale:percent-order (one decisive cell per literal group off both readings): refused; under the pre-committed narrowing production has 0 violations (every differing percent rule withheld, justified)', async () => {
    const flipped = new Set(['po-lo-pq', 'po-lo-qp', 'po-hi-qp', 'po-hi-pq'].map(id => designated.find(x => x.block.id === id && x.role === 'decisive')!.ref))
    const off = (f: string): string => `FF${f.slice(2, 4)}${(parseInt(f.slice(4, 6), 16) ^ 0x10).toString(16).toUpperCase().padStart(2, '0')}${f.slice(6)}`
    const { root, manifest: m } = syntheticRoot(plantedFills(caseOf(ID), cell => (flipped.has(cell.ref) ? off(s2PercentOrderPrediction(cell, S2_PERCENT_ORDER_DEFAULT_WORLD)) : s2PercentOrderPrediction(cell, S2_PERCENT_ORDER_DEFAULT_WORLD))))
    const res = verifyCfCase(m.cases.find(x => x.id === ID)!, { rootDir: root })
    expect(res.kind).toBe('verified')
    if (res.kind !== 'verified') return
    expect((res.refusedCapabilities ?? []).map(x => [x.capabilityId, x.code])).toEqual([[S2_PERCENT_ORDER_CAP, 'formula-hypothesis-none']])
    const unregistered = await stage2Conformance(m, root, [ID])
    expect(unregistered.violations.map(v => v.split(':')[1]).sort()).toEqual([...flipped].sort())
    const narrowed = await withRegistration({ narrowed: [CF_SCALE_PERCENT_ORDER_CAP] }, () => stage2Conformance(m, root, [ID]))
    expect(narrowed.violations).toEqual([])
    expect(narrowed.withheld).toBeGreaterThan(unregistered.withheld)
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
    expect(c.range).toBe('A1:AC27')
    expect(c.expected.cells.length).toBe(150)
    expect(c.expected.cells.length).toBeLessThan(300)
  })
  it('decisive cells hold the smaller threshold, one bucket apart; both signs for percent, one for percentile; controls agree', () => {
    const decisive = (id: string): Array<[number, number]> => { const b = blocks.find(x => x.id === id)!; return b.designated.filter(d => d.role === 'decisive').map(d => [s5IconPoBucket(b, b.values[d.at]!, 'lerp-first'), s5IconPoBucket(b, b.values[d.at]!, 'multiply-first')]) }
    expect(decisive('ipc-up')).toEqual([[0, 1], [0, 1], [1, 2], [1, 2]])
    expect(decisive('ipc-down')).toEqual([[1, 0], [1, 0], [2, 1], [2, 1]])
    expect(decisive('ipl-up')).toEqual(Array(5).fill([0, 1]))
    expect(decisive('ipc-tol')).toEqual([])
    expect(decisive('ipl-tol')).toEqual([])
    for (const b of blocks) for (const v of b.values) expect(String(Number(String(v))), b.id).toBe(String(v))
  })
  it('the text scope admits the whole range in one batch, under the 8,388,608-byte cap (twin blocks two to a column)', async () => {
    const c = caseOf(S5_ICON_PO_ID)
    const { sheet } = await sheetOf(c)
    const bytes = estimateCfPrepareBytes(sheet.conditionalFormatting!, [{ firstRow: 0, lastRow: 26, firstCol: 0, lastCol: 28 }])
    expect(bytes).toBe(5740304)
    expect(bytes!).toBeLessThan(8388608 * 0.7)
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
      [{ narrowed: [CF_ICON_PERCENTILE_ORDER_CAP] }, b => (b.type === 'percentile' && b.id !== 'ipl-tol' ? 'withheld' : 'lerp-first')],
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

// ---- a synthetic measured capture of stage5-icon-percent-order: real icon crops through verifyCfCase, the Stage 5 audit and
// records, and compareCfCase ----
/**
 * Icon pixels of each 3Arrows class, cut from the committed capture of stage5-cfvo-relative-a1 (Excel 16.115, the same 84 pt
 * columns and 20 pt centred rows): B11 shows 3Arrows:2, D11 3Arrows:0, D12 3Arrows:1 (its formulaEvaluation). The left
 * ICON_CROP_WIDTH_PX + 8 px of the cell are copied (the right-aligned value text lies outside them); its same-content control's
 * same region is the control's pixels.
 */
const S5_SOURCE = 'stage5-cfvo-relative-a1'
const S5_CLASS_SOURCE: Readonly<Record<number, string>> = { 0: 'D11', 1: 'D12', 2: 'B11' }
const S5_CUT_PX = 36
type IconWorld = { percent: CfPercentOrderReading; percentile: CfPercentOrderReading; flip?: ReadonlySet<string> }
function s5CellBox(sidecar: { colWidthsPt: readonly number[]; rowHeightsPt: readonly number[] }, ref: string): { x0: number; y0: number; x1: number; y1: number } {
  const { row, col } = rc(ref)
  const sum = (xs: readonly number[], n: number): number => Math.round(xs.slice(0, n).reduce((a, b) => a + b, 0))
  return { x0: sum(sidecar.colWidthsPt, col), x1: sum(sidecar.colWidthsPt, col + 1), y0: sum(sidecar.rowHeightsPt, row), y1: sum(sidecar.rowHeightsPt, row + 1) }
}
async function s5SyntheticRoot(w: IconWorld): Promise<{ root: string; manifest: CfManifest }> {
  const root = mkdtempSync(join(tmpdir(), 'cf-s5-percent-order-'))
  const copy = (rel: string): void => { if (!existsSync(resolve(REPO_ROOT, rel))) return; mkdirSync(dirname(join(root, rel)), { recursive: true }); copyFileSync(resolve(REPO_ROOT, rel), join(root, rel)) }
  for (const x of manifest.cases.filter(k => k.stage >= 4)) { copy(x.input); copy(x.saved); for (const cap of x.captures) { copy(cap.pngPath); copy(cap.sidecarPath) } }
  // The glyph keys are read through the root's own manifest (the committed one: the key cases are unchanged).
  copy('tests/fixtures/xlsx-cf/manifest.json')
  const src = caseOf(S5_SOURCE)
  const srcSide = JSON.parse(readFileSync(resolve(REPO_ROOT, src.captures[0]!.sidecarPath), 'utf8')) as { colWidthsPt: number[]; rowHeightsPt: number[] }
  const srcImg = await loadImage(readFileSync(resolve(REPO_ROOT, src.captures[0]!.pngPath)))
  const srcControl = (ref: string): string => src.expected.cells.find(e => e.ref === ref)!.controlRef
  const c = caseOf(S5_ICON_PO_ID)
  const input = new Uint8Array(readFileSync(resolve(REPO_ROOT, c.input)))
  // Excel's resave (as it saved stage5-cfvo-relative-a1): the default cell style vertically centred, so the saved workbook
  // selects the centred glyph key (iconKeyPlacementReader). Everything else is the generated workbook.
  const zip = await JSZip.loadAsync(readFileSync(resolve(REPO_ROOT, c.saved)))
  const styles = await zip.file('xl/styles.xml')!.async('string')
  const centred = styles.replace(/<cellXfs([^>]*)>([\s\S]*?)<\/cellXfs>/, (_m, a: string, body: string) => `<cellXfs${a}>${body.replace(/<xf([^>]*?)\/>/g, '<xf$1><alignment vertical="center"/></xf>')}</cellXfs>`)
  if (centred === styles) throw new Error('styles unchanged')
  zip.file('xl/styles.xml', centred)
  const saved = new Uint8Array(await zip.generateAsync({ type: 'uint8array', compression: 'DEFLATE' }))
  mkdirSync(dirname(join(root, c.saved)), { recursive: true })
  writeFileSync(join(root, c.saved), saved)
  const geom = computeWorksheetRangeGeometry(readZipPartsSync(saved)!, c.range)!
  const side = { colWidthsPt: geom.colWidthsPt, rowHeightsPt: geom.rowHeightsPt }
  const width = Math.round(geom.expectedWidthPt) + 1, height = Math.round(geom.expectedHeightPt) + 1
  const canvas = createCanvas(width, height)
  const ctx = canvas.getContext('2d')
  ctx.fillStyle = '#FFFFFF'
  ctx.fillRect(0, 0, width, height)
  // Every cell of the case: designated cells and twins draw their class's icon, controls the source control's pixels.
  const values = new Map<string, number>()
  { const { sheet } = await sheetOf(c); for (const r of sheet.rows) for (const cl of r.cells) if (typeof cl.value === 'number') values.set(`${colName(cl.col)}${r.index + 1}`, cl.value) }
  const classOf = new Map<string, number>()
  for (const b of s5IconPoBlocks()) {
    const r = b.type === 'percent' ? w.percent : w.percentile
    for (const d of b.designated) {
      const ref = `${b.col}${2 + d.at}`
      const k = s5IconPoBucket(b, b.values[d.at]!, r)
      classOf.set(ref, w.flip?.has(ref) ? (k === 2 ? 0 : 2) : k)
    }
  }
  for (const m of c.markers.filter(x => x.capabilityId === S5_ICON_PO_TWIN)) {
    const [a, z] = m.sqref.split(':') as [string, string]
    const [t1, t2] = m.icon!.cfvo.slice(1).map(x => Number(x.val ?? x.f)) as [number, number]
    for (let r = rowOf(a); r <= rowOf(z); r++) { const ref = `${colOf(a)}${r}`; const v = values.get(ref)!; classOf.set(ref, v >= t2 ? 2 : v >= t1 ? 1 : 0) }
  }
  const paste = (from: string, to: string): void => {
    const s0 = s5CellBox(srcSide, from), d0 = s5CellBox(side, to)
    ctx.drawImage(srcImg, s0.x0 + 1, s0.y0 + 1, S5_CUT_PX, s0.y1 - s0.y0 - 1, d0.x0 + 1, d0.y0 + 1, S5_CUT_PX, d0.y1 - d0.y0 - 1)
  }
  const observations: Record<string, { displayFormatFill: string; baseFill: string; strikethrough: boolean }> = {}
  for (const e of c.expected.cells) {
    observations[e.ref] = { displayFormatFill: 'FFFFFFFF', baseFill: 'FFFFFFFF', strikethrough: false }
    const k = classOf.get(e.ref)
    if (k !== undefined) paste(S5_CLASS_SOURCE[k]!, e.ref)
    else paste(srcControl(S5_CLASS_SOURCE[1]!), e.ref)
  }
  const png = new Uint8Array(canvas.toBuffer('image/png'))
  const cap = c.captures[0]!
  mkdirSync(dirname(join(root, cap.pngPath)), { recursive: true })
  writeFileSync(join(root, cap.pngPath), png)
  const at = '2026-11-26T00:00:00.000Z'
  const recorded: Record<string, unknown> = {
    excelVersion: '16.115', route: 'copy-picture', appearance: 'screen', scale: 1, pxPerPt: 1, width, height,
    colWidthsPt: geom.colWidthsPt, rowHeightsPt: geom.rowHeightsPt, geometrySource: 'excel-measured', resavedByExcel: true, resaveSavedSha256: sha(saved),
    sheet: c.sheet, range: c.range, hostTimeZone: 'Asia/Taipei', savedSha256: sha(saved), pngSha256: sha(png), capturedAt: at,
    formulaObservations: observations,
  }
  // As capture.ts records it: the evaluation of the observed crops (the centred glyph key from the committed key capture).
  const formulaEvaluation = evaluateFormulaCase({ ...c, status: 'measured' }, recorded, decodePngSync(png) ?? undefined, geom, readZipPartsSync(saved)!, stage5IconKeys(root))
  writeFileSync(join(root, cap.sidecarPath), JSON.stringify({ lockBefore: 'unlocked', lockAfter: 'unlocked', ...recorded, formulaEvaluation }, null, 2))
  const measured: CfOracleCase = {
    ...c, status: 'measured', reason: undefined, resavedByExcel: true,
    resaveReceipt: { openedRepairFree: true, excelVersion: '16.115', resavedAt: at, inputSha256: sha(input), savedSha256: sha(saved) } as CfOracleCase['resaveReceipt'],
    captures: [{ ...cap, pngSha256: sha(png) }],
    expected: { ...c.expected, provenance: 'excel-observed', inputSha256: sha(input), savedSha256: sha(saved) },
  }
  return { root, manifest: { ...manifest, cases: manifest.cases.map(x => (x.id === S5_ICON_PO_ID ? measured : x)) } }
}
const colName = (col: number): string => { let s = ''; for (let x = col + 1; x > 0; x = Math.floor((x - 1) / 26)) s = String.fromCharCode(65 + ((x - 1) % 26)) + s; return s }

describe('stage5-icon-percent-order on a synthetic measured capture (real icon crops; verifyCfCase, the Stage 5 audit and records, compareCfCase)', () => {
  const ICON_IDS = [CF_ICON_PERCENT_ORDER_CAP, CF_ICON_PERCENTILE_ORDER_CAP] as const
  /**
   * After the capture, a matched id's registry entry is removed (recipe §15.2); the verifier and the audit read the registry
   * directly, so the rehearsal removes the two entries around `fn` and restores them.
   */
  async function withoutRegistryEntries<T>(fn: () => Promise<T>): Promise<T> {
    const reg = CF_STAGE5_NONDISCRIMINATING as Record<string, string>
    const kept = ICON_IDS.map(id => [id, reg[id]!] as const)
    for (const [id] of kept) delete reg[id]
    try { return await fn() } finally { for (const [id, text] of kept) reg[id] = text }
  }
  const decisiveRefs = (id: string): string[] => { const b = s5IconPoBlocks().find(x => x.id === id)!; return b.designated.filter(d => d.role === 'decisive').map(d => `${b.col}${2 + d.at}`) }
  const evaluationOf = (root: string): Record<string, { status: string; matched?: string; code?: string }> =>
    JSON.parse(readFileSync(join(root, caseOf(S5_ICON_PO_ID).captures[0]!.sidecarPath), 'utf8')).formulaEvaluation
  const worlds: ReadonlyArray<[CfPercentOrderReading, IconWorld]> = [
    ['lerp-first', { percent: 'lerp-first', percentile: 'lerp-first' }],
    ['multiply-first', { percent: 'multiply-first', percentile: 'multiply-first' }],
  ]

  for (const [reading, w] of worlds) {
    it(`a planted ${reading} world: the crops classify, both ids select exactly ${reading}; registered it verifies and publishes; production agrees in the Stage 5 records and cf-compare (interactive, paged)`, async () => {
      const { root, manifest: m } = await s5SyntheticRoot(w)
      const ev = evaluationOf(root)
      for (const id of ICON_IDS) expect([ev[id]?.status, ev[id]?.matched], `${id}: ${JSON.stringify(ev[id]).slice(0, 400)}`).toEqual(['discriminating', reading])
      const c = m.cases.find(x => x.id === S5_ICON_PO_ID)!
      // Registry kept (pre-capture): verified, both ids non-discriminating.
      const kept = verifyCfCase(c, { rootDir: root })
      expect(kept.kind).toBe('verified')
      if (kept.kind === 'verified') expect(kept.nonDiscriminatingCapabilities).toEqual([...ICON_IDS].sort())
      const registration: Registration = { readings: { [CF_ICON_PERCENT_ORDER_CAP]: reading, [CF_ICON_PERCENTILE_ORDER_CAP]: reading } }
      await withoutRegistryEntries(async () => {
        const res = verifyCfCase(c, { rootDir: root })
        expect(res.kind).toBe('verified')
        if (res.kind !== 'verified') return
        expect(res.matchedHypotheses).toEqual({ [CF_ICON_PERCENT_ORDER_CAP]: reading, [CF_ICON_PERCENTILE_ORDER_CAP]: reading })
        expect([...res.verifiedCapabilities].filter(x => (ICON_IDS as readonly string[]).includes(x)).sort()).toEqual([...ICON_IDS].sort())
        const { auditStage5Manifest, stage5Publication, CF_STAGE5_PRODUCTION_READINGS } = await import('../scripts/cf-refs/audit')
        const report = auditStage5Manifest(m, { rootDir: root })
        expect(report.ok).toBe(true)
        const pub = stage5Publication(report, { ...CF_STAGE5_PRODUCTION_READINGS, [CF_ICON_PERCENT_ORDER_CAP]: reading, [CF_ICON_PERCENTILE_ORDER_CAP]: reading })
        for (const id of ICON_IDS) expect(pub.published, id).toContain(id)
        // The other reading registered: never published.
        const other = reading === 'lerp-first' ? 'multiply-first' : 'lerp-first'
        const wrong = stage5Publication(report, { ...CF_STAGE5_PRODUCTION_READINGS, [CF_ICON_PERCENT_ORDER_CAP]: other, [CF_ICON_PERCENTILE_ORDER_CAP]: other })
        for (const id of ICON_IDS) expect(wrong.published, id).not.toContain(id)
      })
      const { stage5CaseRecords } = await import('../scripts/cf-refs/audit')
      const held = new Set<string>(ICON_IDS)
      const rows = await withRegistration(registration, () => stage5CaseRecords(m, root, { published: id => VERIFIED_CAPABILITIES.has(id) || held.has(id) }))
      const mine = rows.filter(r => r.caseId === S5_ICON_PO_ID)
      expect(mine.length).toBeGreaterThan(0)
      expect(rows.filter(r => r.compared && !r.agrees).map(r => `${r.caseId}:${r.ref} ${r.disagreement}`)).toEqual([])
      const { compareCfCase } = await import('../scripts/cf-compare')
      for (const mode of ['interactive', 'paged'] as const) {
        const b = await withRegistration(registration, () => compareCfCase(c, mode, { rootDir: root }))
        expect(`${mode}: ${b.outcome} ${b.outcome === 'pass' ? '' : b.reason}`).toBe(`${mode}: pass `)
      }
    }, 900_000)
  }

  it('a planted multiply-first world under production unregistered (lerp-first): the records and cf-compare do not agree silently (caught, or the icons withheld and justified)', async () => {
    const { root, manifest: m } = await s5SyntheticRoot({ percent: 'multiply-first', percentile: 'multiply-first' })
    const { stage5CaseRecords } = await import('../scripts/cf-refs/audit')
    const held = new Set<string>(ICON_IDS)
    const rows = await stage5CaseRecords(m, root, { published: id => VERIFIED_CAPABILITIES.has(id) || held.has(id) })
    const bad = rows.filter(r => r.caseId === S5_ICON_PO_ID && r.compared && !r.agrees).map(r => r.ref).sort()
    expect(bad).toEqual([...decisiveRefs('ipc-up'), ...decisiveRefs('ipc-down'), ...decisiveRefs('ipl-up')].sort())
  }, 900_000)

  it('a planted "neither" world (one decisive cell per evidence group off both readings): both ids refuse (formula-hypothesis-none); under the pre-committed narrowing production withholds the differing rules, justified, with no disagreement, and cf-compare passes', async () => {
    const flip = new Set(['ipc-up', 'ipc-down', 'ipl-up'].map(id => decisiveRefs(id)[0]!))
    const { root, manifest: m } = await s5SyntheticRoot({ percent: 'lerp-first', percentile: 'lerp-first', flip })
    const ev = evaluationOf(root)
    for (const id of ICON_IDS) expect([ev[id]?.status, ev[id]?.code], id).toEqual(['refused', 'formula-hypothesis-none'])
    const c = m.cases.find(x => x.id === S5_ICON_PO_ID)!
    await withoutRegistryEntries(async () => {
      const res = verifyCfCase(c, { rootDir: root })
      expect(res.kind).toBe('verified')
      if (res.kind !== 'verified') return
      expect((res.refusedCapabilities ?? []).map(x => [x.capabilityId, x.code]).sort()).toEqual(ICON_IDS.map(id => [id, 'formula-hypothesis-none']).sort())
      expect(res.matchedHypotheses).toEqual({})
    })
    const narrowing: Registration = { narrowed: [...ICON_IDS] }
    const { stage5CaseRecords } = await import('../scripts/cf-refs/audit')
    const held = new Set<string>(ICON_IDS)
    const rows = await withRegistration(narrowing, () => stage5CaseRecords(m, root, { published: id => VERIFIED_CAPABILITIES.has(id) || held.has(id) }))
    expect(rows.filter(r => r.compared && !r.agrees).map(r => `${r.caseId}:${r.ref} ${r.disagreement}`)).toEqual([])
    // The differing rules (ipc-up, ipc-down, ipl-up) are withheld with the narrowing's justification; the tolerance rules paint.
    for (const ref of [...decisiveRefs('ipc-up'), ...decisiveRefs('ipc-down'), ...decisiveRefs('ipl-up')]) {
      const r = rows.find(x => x.caseId === S5_ICON_PO_ID && x.ref === ref)!
      expect(r.justification.map(j => j.refused), ref).toContain('cf-unmeasured-percent-order')
    }
    const { compareCfCase } = await import('../scripts/cf-compare')
    for (const mode of ['interactive', 'paged'] as const) {
      const b = await withRegistration(narrowing, () => compareCfCase(c, mode, { rootDir: root }))
      expect(`${mode}: ${b.outcome} ${b.outcome === 'pass' ? '' : b.reason}`).toBe(`${mode}: pass `)
    }
  }, 900_000)
})
