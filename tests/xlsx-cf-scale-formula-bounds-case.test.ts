// @vitest-environment node
/**
 * The formula-bounds follow-up (brief cf-formula-bounds), Phase 1 tasks 3-5: stage2-scale-formula-bounds, the pre-registered
 * and registrable case for colour-scale segments bounded by a formula-resolved cfvo in the shapes whose rule a capture settled,
 * and the rehearsal of each reading's registration. Checked here:
 * 1. the case design: ownership, registries, size, values, the formula cfvo and their literal cells, every designated cell's
 *    role, >= 3 decisive cells per evidence configuration, the twins' IEEE-integer products, the text scope;
 * 2. the independent model is production, unregistered (the formula-bounded cells of the settled shapes withheld) and under
 *    each registration (with and without the near-top id);
 * 3. the evaluator on synthetic observations: each planted reading comes back as the one match; worlds no reading explains are
 *    refused; it fails closed on broken controls;
 * 4. the audit gate end to end on a synthetic measured capture (real verifyCfCase and auditStage2Manifest);
 * 5. each registration against every committed capture: Stage 2 unchanged; Stage 5 agrees with Excel, only E4 / E5 of
 *    stage5-cfvo-nonnumeric-v2 newly painted at Excel's fill; the conformance pass and cf-compare (interactive, paged) pass on
 *    each planted world, unregistered and under its own registration; the other reading is caught.
 * This checks the case and production, not Excel.
 */
import { copyFileSync, existsSync, mkdirSync, mkdtempSync, readFileSync, writeFileSync } from 'node:fs'
import { tmpdir } from 'node:os'
import { dirname, join, resolve } from 'node:path'
import { createHash } from 'node:crypto'
import { createCanvas } from 'canvas'
import { afterEach, describe, expect, it } from 'vitest'
import {
  CF_STAGE2_CAPABILITIES,
  CF_STAGE2_CASES,
  CF_STAGE2_NONDISCRIMINATING,
  S2_FORMULA_BOUNDS,
  S2_FORMULA_BOUNDS_CAP,
  S2_FORMULA_BOUNDS_CELLS,
  S2_FORMULA_BOUNDS_ID,
  S2_FORMULA_BOUNDS_NEAR_TOP_CAP,
  S2_FORMULA_BOUNDS_READINGS,
  S2_FORMULA_BOUNDS_TWIN,
  s2FormulaBoundsCells,
  s2FormulaBoundsLayout,
  s2FormulaBoundsNearTop,
  s2FormulaBoundsPrediction,
  s2FormulaBoundsResult,
  s2FormulaBoundsRound,
  s2FormulaBoundsSegment,
  s2FormulaBoundsWithheldToday,
  s2UwLine,
  type CfManifest,
  type CfOracleCase,
  type S2FormulaBoundsCell,
  type S2FormulaBoundsReading,
} from '../scripts/cf-refs/cases'
import { CF_MANIFEST_PATH } from '../scripts/cf-refs/generate'
import { computeWorksheetRangeGeometry, evaluateFormulaCapabilities, formulaControlsOf, isCellInSqref, readZipPartsSync, stage2ReadingsObservationOf, verifyCfCase } from '../scripts/cf-refs/verify'
import { stage2Conformance } from './helpers/cf-stage2-conformance'
import { OfficePackage } from '../src/core/zip'
import { parseXlsx } from '../src/xlsx/parse'
import { cfDocumentEvaluationOf, prepareCfPaint } from '../src/xlsx/cf/render'
import { cfRuntimeEvidence } from '../src/xlsx/cf/predicates'
import { CF_FORMULA_BOUNDS_CAP, CF_FORMULA_BOUNDS_NEAR_TOP_CAP, CF_FORMULA_BOUNDS_READINGS, cfFormulaBoundFifteenDigit, injectCfFormulaBoundsReadings, safeFraction } from '../src/xlsx/cf/colorScale'
import { estimateCfPrepareBytes } from '../src/xlsx/cf/service'
import { VERIFIED_CAPABILITIES, VERIFIED_STAGE2_CAPABILITIES, withExactCfEvidence } from '../src/xlsx/cf/evidence'

const REPO_ROOT = resolve(__dirname, '..')
const ID = S2_FORMULA_BOUNDS_ID
const CAP = S2_FORMULA_BOUNDS_CAP
const TWIN = S2_FORMULA_BOUNDS_TWIN
const NEAR = S2_FORMULA_BOUNDS_NEAR_TOP_CAP
const HEAVY = 300_000

afterEach(() => { injectCfFormulaBoundsReadings(null) })

const manifest = JSON.parse(readFileSync(CF_MANIFEST_PATH, 'utf8')) as CfManifest
const caseOf = (): CfOracleCase => {
  const c = manifest.cases.find(x => x.id === ID)
  if (!c) throw new Error(`${ID} is not in the manifest (generate it with generate.ts --case ${ID})`)
  return c
}
const colOf = (ref: string): string => /^[A-Z]+/.exec(ref)![0]
const rowOf = (ref: string): number => Number(/\d+$/.exec(ref)![0])
const rc = (ref: string): { row: number; col: number } => ({ row: rowOf(ref) - 1, col: [...colOf(ref)].reduce((n, ch) => n * 26 + ch.charCodeAt(0) - 64, 0) - 1 })
const partsOf = (c: CfOracleCase, file = c.saved): Map<string, Uint8Array> => readZipPartsSync(new Uint8Array(readFileSync(resolve(REPO_ROOT, file))))!
const sheetXml = (c: CfOracleCase): string => new TextDecoder().decode(partsOf(c, c.input).get('xl/worksheets/sheet1.xml')!)
const valuesOf = (c: CfOracleCase): Map<string, number> => {
  const out = new Map<string, number>()
  for (const m of sheetXml(c).matchAll(/<c r="([A-Z]+\d+)"[^>]*><v>([^<]+)<\/v><\/c>/g)) out.set(m[1]!, Number(m[2]))
  return out
}
const layout = s2FormulaBoundsLayout()
const twinFill = new Map(layout.twins.map(t => [t.ref, t.fill]))
const designated = s2FormulaBoundsCells()
const cellAt = new Map(designated.map(c => [c.ref, c]))
const decisive = (c: S2FormulaBoundsCell): boolean => s2FormulaBoundsPrediction(c, 'as-literal') !== s2FormulaBoundsPrediction(c, 'fifteen-digit')
const EVIDENCE = ['f2-num-hi', 'f2-formula-lo', 'f2-painted', 'f2-expr', 'f2-percent', 'f3-lower', 'f3-upper', 'f3-upper-mixed']
const COINCIDE = ['f2-coincide-expr', 'f3-coincide-ref']
const withheldRefs = designated.filter(s2FormulaBoundsWithheldToday).map(c => c.ref)
/** The cells near-top under a reading (withheld under it while the near-top id is not registered at the same reading). */
const nearTopUnder = (r: S2FormulaBoundsReading): string[] => designated.filter(c => s2FormulaBoundsNearTop(c, r)).map(c => c.ref)

/**
 * Rehearse a registration: the reading selected through the seam for the main id (and, `nearTop`, for the near-top id: by
 * default the same reading, or another, or none), with the evidence of each registered id held, around `fn`.
 */
async function withRegistration<T>(reading: S2FormulaBoundsReading, fn: () => Promise<T>, nearTop: S2FormulaBoundsReading | false = reading): Promise<T> {
  injectCfFormulaBoundsReadings({ [CF_FORMULA_BOUNDS_CAP]: reading, ...(nearTop === false ? {} : { [CF_FORMULA_BOUNDS_NEAR_TOP_CAP]: nearTop }) })
  const caps = [...VERIFIED_CAPABILITIES, CF_FORMULA_BOUNDS_CAP, ...(nearTop === false ? [] : [CF_FORMULA_BOUNDS_NEAR_TOP_CAP])]
  try { return await withExactCfEvidence(caps, fn) } finally { injectCfFormulaBoundsReadings(null) }
}

/** A synthetic world: each designated cell's fill by `world` (a reading or a function), twins as fixed, positive controls as fixed, controls white. */
function plantedFills(c: CfOracleCase, world: S2FormulaBoundsReading | ((cell: S2FormulaBoundsCell) => string)): Map<string, string> {
  const out = new Map<string, string>()
  for (const e of c.expected.cells) out.set(e.ref, e.matched ? e.fillRgb ?? '' : 'FFFFFFFF')
  for (const [ref, fill] of twinFill) out.set(ref, fill)
  for (const cell of designated) out.set(cell.ref, typeof world === 'string' ? s2FormulaBoundsPrediction(cell, world) : world(cell))
  return out
}
const plantedSidecar = (c: CfOracleCase, fills: Map<string, string>): { observedCells: Array<Record<string, unknown>> } => ({
  observedCells: c.expected.cells.filter(e => fills.get(e.ref) !== undefined).map(e => ({ ref: e.ref, controlRef: e.controlRef, fillSource: 'excel-display-format', fillRgb: fills.get(e.ref) })),
})
const judge = (c: CfOracleCase, fills: Map<string, string>) => evaluateFormulaCapabilities(c.scaleReadings!, formulaControlsOf(c), stage2ReadingsObservationOf(plantedSidecar(c, fills)))
/** A fill moved by `by` steps on every channel toward `to` (the per-channel step direction). */
const stepped = (fill: string, from: string, to: string, by: number): string => [0, 2, 4, 6].map(i => {
  const x = parseInt(fill.slice(i, i + 2), 16), a = parseInt(from.slice(i, i + 2), 16), b = parseInt(to.slice(i, i + 2), 16)
  return (a === b ? x : x + (b > a ? by : -by)).toString(16).toUpperCase().padStart(2, '0')
}).join('')

async function sheetOf(c: CfOracleCase, root = REPO_ROOT) {
  const doc = await parseXlsx(await OfficePackage.load(new Uint8Array(readFileSync(resolve(root, c.saved)))))
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

describe('stage2-scale-formula-bounds: the case', () => {
  it('is appended last; owns its three ids; both questions unregistered and in the registry (captured d2660d46: no single reading matched, so nothing is registered), the twins recorded', async () => {
    // Appended last at its time; the percent-order follow-up is appended after it.
    expect(CF_STAGE2_CASES.map(c => c.id).slice(-3)).toEqual(['stage2-scale-upper-segment', ID, 'stage2-scale-percent-order'])
    for (const id of [CAP, NEAR, TWIN]) {
      expect(CF_STAGE2_CAPABILITIES.filter(e => e.capabilityId === id).map(e => e.caseId), id).toEqual([ID])
      expect(CF_STAGE2_NONDISCRIMINATING[id], id).toMatch(/./)
    }
    expect(CAP).toBe(CF_FORMULA_BOUNDS_CAP)
    expect(NEAR).toBe(CF_FORMULA_BOUNDS_NEAR_TOP_CAP)
    expect([...S2_FORMULA_BOUNDS_READINGS]).toEqual([...CF_FORMULA_BOUNDS_READINGS])
    const c = caseOf()
    expect(c.status).toBe('measured')
    expect(c.scaleReadings).toEqual(CF_STAGE2_CASES.find(t => t.id === ID)!.scaleReadings)
    expect(c.scaleReadings!.map(x => [x.id, x.capabilityId, Object.keys(x.hypotheses), x.role ?? (x.record ? 'record' : 'evidence')])).toEqual([
      ...[...EVIDENCE, ...COINCIDE].map(id => [id, CAP, ['as-literal', 'fifteen-digit'], 'evidence']),
      ['near-top', NEAR, ['as-literal', 'fifteen-digit'], 'evidence'],
      ['controls', CAP, ['as-literal', 'fifteen-digit'], 'negative-control'],
    ])
    const { CF_STAGE2_PRODUCTION_READINGS } = await import('../scripts/cf-refs/audit')
    for (const id of [CAP, NEAR, TWIN]) {
      expect(Object.prototype.hasOwnProperty.call(CF_STAGE2_PRODUCTION_READINGS, id), id).toBe(false)
      expect(VERIFIED_CAPABILITIES.has(id), id).toBe(false)
    }
  })

  it('200 cells: 90 designated (43 decisive, 18 kept, 14 clear, 15 controls; 9 near-top), 93 twins, 2 positive controls, 15 unformatted controls; A1:W14, 28 colour rules', () => {
    const c = caseOf()
    expect(c.range).toBe('A1:W14')
    expect(c.expected.cells.length).toBe(200)
    expect(c.scaleReadings!.map(x => x.refs.length)).toEqual([7, 8, 7, 6, 7, 7, 9, 6, 5, 4, 9, 15])
    const count = (role: string): number => designated.filter(x => x.role === role).length
    expect([designated.length, count('decisive'), count('kept'), count('clear'), count('control')]).toEqual([90, 43, 18, 14, 15])
    expect(layout.twins.length).toBe(93)
    expect(c.expected.cells.filter(e => e.matched && !e.scaleApplied).length).toBe(2)
    expect(c.expected.negativeRefs.length).toBe(15)
    expect(c.markers.length).toBe(28)
    for (const e of c.expected.cells) expect(c.markers.filter(m => isCellInSqref(e.ref, m.sqref)).length, e.ref).toBeLessThanOrEqual(1)
    for (const ref of c.scaleReadings!.flatMap(x => x.refs)) expect(c.expected.cells.find(e => e.ref === ref)!.controlRef, ref).toBe(`${colOf(ref)}14`)
    for (const ref of c.expected.negativeRefs) expect(c.markers.some(m => isCellInSqref(ref, m.sqref)), ref).toBe(false)
    for (const ref of Object.keys(S2_FORMULA_BOUNDS_CELLS)) expect(c.markers.some(m => isCellInSqref(ref, m.sqref)) || isCellInSqref(ref, c.range), ref).toBe(false)
  })

  it('the text scope admits the whole range in one batch, well under the 8,388,608-byte cap', async () => {
    const { sheet } = await sheetOf(caseOf())
    const bytes = estimateCfPrepareBytes(sheet.conditionalFormatting!, [{ firstRow: 0, lastRow: 13, firstCol: 0, lastCol: 22 }])
    expect(bytes).toBe(TEXT_SCOPE_BYTES)
    expect(bytes!).toBeLessThan(8388608 * 0.6)
  })

  it('the workbook: the catalogue values, the literal bound cells at their doubles, every formula cfvo as written; no cell formula, hidden rows, merges or errors', () => {
    const c = caseOf()
    const values = valuesOf(c)
    for (const cell of designated) expect(values.get(cell.ref), cell.ref).toBe(Number(cell.text))
    for (const t of layout.twins) expect(values.get(t.ref), t.ref).toBe(Number(t.text))
    for (const [ref, text] of Object.entries(S2_FORMULA_BOUNDS_CELLS)) {
      expect(values.get(ref), ref).toBe(Number(text))
      expect(String(Number(text)), ref).toBe(text)
    }
    const xml = sheetXml(c)
    expect(xml).not.toMatch(/hidden="1"|<mergeCell|<autoFilter|t="e"|<f>/)
    for (const frag of ['<cfvo type="num" val="$AH$2"/>', '<cfvo type="formula" val="$AH$3"/>', '<cfvo type="formula" val="0.1*3"/>', '<cfvo type="percent" val="$AH$5"/>', '<cfvo type="formula" val="1/3*630"/>']) expect(xml, frag).toContain(frag)
    for (const e of c.expected.cells) expect(typeof values.get(e.ref), e.ref).toBe('number')
  })

  it('the formula results: none but the coincident controls equals its 15-digit rounding (each run here); the expressions evaluate as stated', () => {
    expect(s2FormulaBoundsResult('0.1*3')).toBe(0.30000000000000004)
    expect(s2FormulaBoundsResult('1/3*630')).toBe(210)
    expect(0.1 * 3).toBe(0.30000000000000004)
    expect(1 / 3 * 630).toBe(210)
    const coincide = new Set(['AH11'])
    for (const [ref, text] of Object.entries(S2_FORMULA_BOUNDS_CELLS)) {
      const x = Number(text)
      expect(s2FormulaBoundsRound(x) === x, ref).toBe(coincide.has(ref))
      expect(s2FormulaBoundsRound(x), ref).toBe(cfFormulaBoundFifteenDigit(x))
    }
    expect(Object.fromEntries(Object.entries(S2_FORMULA_BOUNDS_CELLS).map(([r, t]) => [r, String(s2FormulaBoundsRound(Number(t)))]))).toEqual({
      AH2: '0.3', AH3: '0.3', AH4: '2.1', AH5: '33.3333333333333', AH6: '0.3', AH7: '210', AH8: '420', AH9: '315', AH10: '427', AH11: '0.3',
    })
  })

  it('every designated cell is in its designed role; each evidence configuration separates the readings at >= 5 cells, the near-top one at 3, the coincident ones at none', () => {
    for (const c of designated) {
      const s = s2FormulaBoundsSegment(c, 'as-literal')
      expect(Object.is(s.f, safeFraction(c.value, s.lo, s.hi) ?? NaN) || !s.interior, c.ref).toBe(true)
      if (c.role === 'decisive') expect(decisive(c), c.ref).toBe(true)
      else expect(decisive(c), c.ref).toBe(false)
      if (c.role === 'kept') expect(s2FormulaBoundsWithheldToday(c), c.ref).toBe(true)
      if (c.role === 'clear') expect(s2FormulaBoundsWithheldToday(c), c.ref).toBe(false)
    }
    const dec = (id: string): number => { const cfg = layout.configurations.find(x => x.id === id)!; return cfg.refs.filter(r => cfg.hypotheses['as-literal']![r] !== cfg.hypotheses['fifteen-digit']![r]).length }
    expect(Object.fromEntries([...EVIDENCE, ...COINCIDE, 'near-top', 'controls'].map(id => [id, dec(id)]))).toEqual({
      'f2-num-hi': 5, 'f2-formula-lo': 5, 'f2-painted': 5, 'f2-expr': 5, 'f2-percent': 5, 'f3-lower': 5, 'f3-upper': 5, 'f3-upper-mixed': 5, 'f2-coincide-expr': 0, 'f3-coincide-ref': 0, 'near-top': 3, controls: 0,
    })
    // The painted-today group: production paints its decisive cells now (as-literal's arithmetic); every other decisive cell is withheld today.
    for (const c of designated.filter(x => x.role === 'decisive' && !x.block.nearTop)) expect(s2FormulaBoundsWithheldToday(c), c.ref).toBe(c.block.id !== 'f2-painted')
    // The near-top probes: one per column is an endpoint under one reading and a near-top interior cell under the other.
    expect(designated.filter(c => c.block.nearTop && c.role === 'decisive').map(c => `${c.ref} ${S2_FORMULA_BOUNDS_READINGS.map(r => s2FormulaBoundsSegment(c, r).interior ? 'near-top' : 'endpoint').join('/')}`)).toEqual(['I2 near-top/endpoint', 'J2 endpoint/near-top', 'K2 endpoint/near-top'])
    expect(nearTopUnder('as-literal')).toEqual(['I2', 'I3', 'I4', 'J3', 'J4', 'K3', 'K4'])
    expect(nearTopUnder('fifteen-digit')).toEqual(['I3', 'I4', 'J2', 'J3', 'J4', 'K2', 'K3', 'K4'])
  })

  it('the twins: every live prediction is a twin (or a literal positive control in the controls), and every twin\'s product is an IEEE integer on every channel (recomputed)', () => {
    for (const t of layout.twins) {
      const [from, to] = t.rule.colors
      const v = Number(t.text), f = t.rule.cfvo.length === 2 ? v : 1
      const [a, b] = t.rule.cfvo.length === 2 ? [from, to] : [t.rule.colors[Math.max(0, v - 1)]!, t.rule.colors[Math.max(1, v)]!]
      for (const i of [0, 2, 4, 6]) {
        const d = Math.abs(Number.parseInt(b!.slice(i, i + 2), 16) - Number.parseInt(a!.slice(i, i + 2), 16))
        expect(Number.isInteger(f * d), `${t.ref} ${f * d}`).toBe(true)
      }
      const line = S2_FORMULA_BOUNDS.twinSpans.some(sp => from === s2UwLine(sp.from) && to === s2UwLine(sp.to))
      if (!line) expect(t.fill, t.ref).toBe(t.rule.colors[v])
    }
    for (const cfg of layout.configurations) {
      for (const [name, h] of Object.entries(cfg.hypotheses)) for (const [ref, token] of Object.entries(h)) {
        if (token.startsWith('twin:')) expect(twinFill.get(token.slice(5)), `${cfg.id} ${ref}`).toBe(s2FormulaBoundsPrediction(cellAt.get(ref)!, name as S2FormulaBoundsReading))
        else { expect(cfg.id, ref).toBe('controls'); expect(Object.values(cfg.positiveControls ?? {}), ref).toContain(token) }
      }
    }
  })

  it('the independent model is production unregistered: the formula-bounded cells of the settled shapes are withheld, every other cell paints as-literal\'s prediction', async () => {
    const c = caseOf()
    const fills = await productionFills(c, c.expected.cells.map(e => e.ref))
    const withheld: string[] = []
    for (const e of c.expected.cells) {
      const got = fills.get(e.ref)
      if (got === undefined && e.matched) { withheld.push(e.ref); continue }
      const cell = cellAt.get(e.ref)
      expect(got, e.ref).toBe(!e.matched ? undefined : cell ? s2FormulaBoundsPrediction(cell, 'as-literal') : twinFill.get(e.ref) ?? e.fillRgb)
    }
    expect(withheld.sort()).toEqual([...withheldRefs].sort())
    expect(withheldRefs.length).toBe(WITHHELD_TODAY)
    // A proposed reading without the ids' evidence: the same.
    injectCfFormulaBoundsReadings({ [CF_FORMULA_BOUNDS_CAP]: 'as-literal', [CF_FORMULA_BOUNDS_NEAR_TOP_CAP]: 'as-literal' })
    const ev = await productionFills(c, c.expected.cells.map(e => e.ref))
    expect(ev).toEqual(fills)
  }, HEAVY)

  for (const reading of S2_FORMULA_BOUNDS_READINGS) {
    it(`the independent model is production under the ${reading} registration: with the near-top id at the same reading every designated cell paints it; with the main id alone, or the near-top id at the other reading, the cells near-top under ${reading} stay withheld`, async () => {
      const c = caseOf()
      const other = S2_FORMULA_BOUNDS_READINGS.find(x => x !== reading)!
      for (const nearTop of [reading, false, other] as const) {
        const fills = await withRegistration(reading, () => productionFills(c, c.expected.cells.map(e => e.ref)), nearTop)
        for (const e of c.expected.cells) {
          const cell = cellAt.get(e.ref)
          const want = !e.matched || (nearTop !== reading && nearTopUnder(reading).includes(e.ref)) ? undefined : cell ? s2FormulaBoundsPrediction(cell, reading) : twinFill.get(e.ref) ?? e.fillRgb
          expect(fills.get(e.ref), `${e.ref} near-top ${nearTop}`).toBe(want)
        }
      }
    }, HEAVY)
  }
})

describe('stage2-scale-formula-bounds: the evaluator on synthetic observations', () => {
  for (const reading of S2_FORMULA_BOUNDS_READINGS) {
    it(`a planted ${reading} world matches exactly ${reading} for both questions`, () => {
      const c = caseOf()
      const r = judge(c, plantedFills(c, reading))
      expect([r[CAP]!.status, r[CAP]!.matched, r[NEAR]!.status, r[NEAR]!.matched]).toEqual(['discriminating', reading, 'discriminating', reading])
    })
  }

  it('no reading matches: one decisive cell of any configuration at the other reading, or one kept or coincident cell drawn one step off, is refused', () => {
    const c = caseOf()
    for (const [base, flip] of [['as-literal', 'fifteen-digit'], ['fifteen-digit', 'as-literal']] as const) {
      for (const cfg of EVIDENCE) {
        const target = designated.find(x => x.configuration === cfg && decisive(x))!
        const r = judge(c, plantedFills(c, x => s2FormulaBoundsPrediction(x, x === target ? flip : base)))
        expect([r[CAP]!.status, r[CAP]!.code], `${base} with ${target.ref} at ${flip}`).toEqual(['refused', 'formula-hypothesis-none'])
      }
    }
    for (const cfg of [...EVIDENCE, ...COINCIDE]) {
      const target = designated.find(x => x.configuration === cfg && x.role === 'kept')
      if (!target) continue
      const s = s2FormulaBoundsSegment(target, 'as-literal')
      const r = judge(c, plantedFills(c, x => (x === target ? stepped(s2FormulaBoundsPrediction(x, 'as-literal'), s.from, s.to, 1) : s2FormulaBoundsPrediction(x, 'as-literal'))))
      expect([r[CAP]!.status, r[CAP]!.code], `${target.ref} one step off`).toEqual(['refused', 'formula-hypothesis-none'])
    }
  })

  it('the near-top question is judged on its own: as-literal everywhere but the near-top probes at fifteen-digit splits the two ids; one kept near-top probe one step off refuses only the near-top id', () => {
    const c = caseOf()
    const split = judge(c, plantedFills(c, x => s2FormulaBoundsPrediction(x, x.configuration === 'near-top' ? 'fifteen-digit' : 'as-literal')))
    expect([split[CAP]!.status, split[CAP]!.matched, split[NEAR]!.status, split[NEAR]!.matched]).toEqual(['discriminating', 'as-literal', 'discriminating', 'fifteen-digit'])
    const target = designated.find(x => x.configuration === 'near-top' && x.role === 'kept')!
    const s = s2FormulaBoundsSegment(target, 'as-literal')
    const r = judge(c, plantedFills(c, x => (x === target ? stepped(s2FormulaBoundsPrediction(x, 'as-literal'), s.from, s.to, 1) : s2FormulaBoundsPrediction(x, 'as-literal'))))
    expect([r[CAP]!.status, r[CAP]!.matched, r[NEAR]!.status, r[NEAR]!.code]).toEqual(['discriminating', 'as-literal', 'refused', 'formula-hypothesis-none'])
  })

  it('fails closed: a formatted control, a missing or wrong positive control, a twin rendering nothing, a moved negative control', () => {
    const c = caseOf()
    for (const reading of S2_FORMULA_BOUNDS_READINGS) {
      const fills = plantedFills(c, reading)
      const edit = (ref: string, fill: string | undefined): Map<string, string> => { const m = new Map(fills); if (fill === undefined) m.delete(ref); else m.set(ref, fill); return m }
      const pc = layout.positiveControls[0]!.ref
      expect(judge(c, edit('A14', 'FF14F028'))[CAP]!.status, reading).toBe('refused')
      expect(judge(c, edit(pc, undefined))[CAP]!.code, reading).toBe('fill-unobserved')
      expect(judge(c, edit(pc, 'FF000000'))[CAP]!.code, reading).toBe('formula-positive-control-failed')
      expect(judge(c, edit([...twinFill.keys()][0]!, 'FFFFFFFF'))[CAP]!.status, reading).toBe('refused')
      // An endpoint control (A9, the value 0 of f2-num-hi: P) drawn one step off P: nothing matches.
      const moved = judge(c, edit('A9', 'FF15EF29'))
      expect([moved[CAP]!.status, moved[CAP]!.code], reading).toEqual(['refused', 'formula-hypothesis-none'])
    }
  })
})

// ---- a synthetic measured capture through the real verifier, audit, conformance pass and cf-compare (temp root) ----
const sha = (b: Uint8Array): string => createHash('sha256').update(b).digest('hex')

function syntheticRoot(fills: Map<string, string>): { root: string; manifest: CfManifest } {
  const root = mkdtempSync(join(tmpdir(), 'cf-s2-formula-bounds-'))
  const copy = (rel: string): void => { if (!existsSync(resolve(REPO_ROOT, rel))) return; mkdirSync(dirname(join(root, rel)), { recursive: true }); copyFileSync(resolve(REPO_ROOT, rel), join(root, rel)) }
  for (const x of manifest.cases.filter(k => k.stage === 2)) {
    copy(x.input); copy(x.saved)
    for (const cap of x.captures) { copy(cap.pngPath); copy(cap.sidecarPath) }
  }
  const c = caseOf()
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
  const at = '2026-11-25T00:00:00.000Z'
  writeFileSync(join(root, cap.sidecarPath), JSON.stringify({
    excelVersion: '16.115', route: 'copy-picture', appearance: 'screen', scale: 1, pxPerPt: 1, width, height,
    colWidthsPt: geom.colWidthsPt, rowHeightsPt: geom.rowHeightsPt, geometrySource: 'excel-measured', resavedByExcel: true, resaveSavedSha256: sha(saved),
    sheet: c.sheet, range: c.range, hostTimeZone: 'Asia/Taipei', savedSha256: sha(saved), pngSha256: sha(png), capturedAt: at,
    observedCells: c.expected.cells.map(e => ({ ref: e.ref, controlRef: e.controlRef, fillSource: 'excel-display-format', fillRgb: fills.get(e.ref) })),
  }, null, 2))
  const measured: CfOracleCase = {
    ...c,
    status: 'measured',
    reason: undefined,
    resavedByExcel: true,
    resaveReceipt: { openedRepairFree: true, excelVersion: '16.115', resavedAt: at, inputSha256: sha(input), savedSha256: sha(saved) } as CfOracleCase['resaveReceipt'],
    captures: [{ ...cap, pngSha256: sha(png) }],
    expected: { ...c.expected, provenance: 'excel-observed', cells: c.expected.cells.map(e => (e.matched ? { ...e, fillSource: 'excel-display-format' as const } : e)) },
  }
  return { root, manifest: { ...manifest, cases: manifest.cases.map(x => (x.id === ID ? measured : x)) } }
}

describe('the audit gate on a synthetic measured capture (real verifyCfCase and auditStage2Manifest)', () => {
  const disposition = (r: { perCapability: ReadonlyArray<{ capabilityId: string; disposition: string }> }, id: string) => r.perCapability.find(p => p.capabilityId === id)!.disposition
  const before = [...VERIFIED_STAGE2_CAPABILITIES].sort()
  const kept = CF_STAGE2_NONDISCRIMINATING
  const removed = Object.fromEntries(Object.entries(CF_STAGE2_NONDISCRIMINATING).filter(([id]) => id !== CAP && id !== NEAR))

  it('the real manifest: Stage 2 audits ok and complete; the three ids unverified-nondiscriminating (recipe pending); the verified set unchanged', async () => {
    const { auditStage2Manifest } = await import('../scripts/cf-refs/audit')
    const r = auditStage2Manifest(manifest, { rootDir: REPO_ROOT })
    expect([r.ok, r.stageComplete]).toEqual([true, true])
    for (const id of [CAP, NEAR, TWIN]) expect(disposition(r, id), id).toBe('unverified-nondiscriminating')
    expect(r.perCapability.filter(p => p.disposition === 'verified').map(p => p.capabilityId).sort()).toEqual(before)
  }, HEAVY)

  for (const reading of S2_FORMULA_BOUNDS_READINGS) {
    it(`planted ${reading}: matched itself; registry kept: unverified; entries removed, unregistered: withheld; registered: verified; the main id alone: the near-top id withheld; the other reading registered: withheld`, async () => {
      const { auditStage2Manifest, CF_STAGE2_PRODUCTION_READINGS } = await import('../scripts/cf-refs/audit')
      const { root, manifest: m } = syntheticRoot(plantedFills(caseOf(), reading))
      const res = verifyCfCase(m.cases.find(x => x.id === ID)!, { rootDir: root })
      expect(res.kind, res.kind === 'refusal' ? `${res.code}: ${res.reason}` : '').toBe('verified')
      if (res.kind !== 'verified') return
      expect(res.matchedHypotheses).toEqual({ [CAP]: reading, [NEAR]: reading })
      const k = auditStage2Manifest(m, { rootDir: root, nonDiscriminating: kept })
      expect(k.ok).toBe(true)
      for (const id of [CAP, NEAR, TWIN]) expect(disposition(k, id), id).toBe('unverified-nondiscriminating')
      expect(disposition(auditStage2Manifest(m, { rootDir: root, nonDiscriminating: removed }), CAP)).toBe('supported-structural')
      const registered = auditStage2Manifest(m, { rootDir: root, nonDiscriminating: removed, productionReadings: { ...CF_STAGE2_PRODUCTION_READINGS, [CAP]: reading, [NEAR]: reading } })
      for (const id of [CAP, NEAR]) expect(disposition(registered, id), id).toBe('verified')
      expect(registered.perCapability.filter(p => p.disposition === 'verified').map(p => p.capabilityId).sort()).toEqual([...before, CAP, NEAR].sort())
      const main = auditStage2Manifest(m, { rootDir: root, nonDiscriminating: removed, productionReadings: { ...CF_STAGE2_PRODUCTION_READINGS, [CAP]: reading } })
      expect([disposition(main, CAP), disposition(main, NEAR)]).toEqual(['verified', 'supported-structural'])
      const other = S2_FORMULA_BOUNDS_READINGS.find(x => x !== reading)!
      expect(disposition(auditStage2Manifest(m, { rootDir: root, nonDiscriminating: removed, productionReadings: { ...CF_STAGE2_PRODUCTION_READINGS, [CAP]: other, [NEAR]: other } }), CAP)).toBe('supported-structural')
    }, HEAVY)
  }

  it('planted "no reading matches" (as-literal with one decisive cell per evidence configuration at fifteen-digit): formula-hypothesis-none, never verified, and production keeps the withhold', async () => {
    const { auditStage2Manifest, CF_STAGE2_PRODUCTION_READINGS } = await import('../scripts/cf-refs/audit')
    const flipped = new Set(EVIDENCE.map(cfg => designated.find(x => x.configuration === cfg && decisive(x) && s2FormulaBoundsWithheldToday(x))?.ref).filter((r): r is string => r !== undefined))
    expect(flipped.size).toBe(EVIDENCE.length - 1)
    const { root, manifest: m } = syntheticRoot(plantedFills(caseOf(), x => s2FormulaBoundsPrediction(x, flipped.has(x.ref) ? 'fifteen-digit' : 'as-literal')))
    const res = verifyCfCase(m.cases.find(x => x.id === ID)!, { rootDir: root })
    expect(res.kind).toBe('verified')
    if (res.kind !== 'verified') return
    expect((res.refusedCapabilities ?? []).map(x => [x.capabilityId, x.code])).toEqual([[CAP, 'formula-hypothesis-none']])
    expect(res.matchedHypotheses).toEqual({ [NEAR]: 'as-literal' })
    for (const reading of S2_FORMULA_BOUNDS_READINGS) expect(disposition(auditStage2Manifest(m, { rootDir: root, nonDiscriminating: removed, productionReadings: { ...CF_STAGE2_PRODUCTION_READINGS, [CAP]: reading } }), CAP), reading).toBe('supported-structural')
    // Unregistered production on that world: every formula-bounded cell of the settled shapes withheld and justified; nothing disagrees.
    const r = await stage2Conformance(m, root, [ID])
    expect(r.violations).toEqual([])
    expect([r.total, r.withheld, r.numSegmentWithheld, r.upperWholeWithheld, r.guardBoundaryWithheld]).toEqual([200, WITHHELD_TODAY, 0, UPPER_WITHHELD, WITHHELD_TODAY - UPPER_WITHHELD])
  }, HEAVY)
})

/** The measured Stage 2 cases before this capture (d2660d46): the rehearsal's earlier captures. */
const STAGE2_MEASURED = manifest.cases.filter(x => x.stage === 2 && x.status === 'measured' && x.id !== ID).map(x => x.id)

describe('the rehearsal: each registration against every committed capture, and this case on a synthetic capture', () => {
  it('each registration: no violation over the measured Stage 2 cases and nothing moves (no committed Stage 2 capture holds a formula bound that resolves)', async () => {
    expect(STAGE2_MEASURED).not.toContain(ID)
    const base = await stage2Conformance(manifest, REPO_ROOT, STAGE2_MEASURED)
    expect(base.violations).toEqual([])
    for (const reading of S2_FORMULA_BOUNDS_READINGS) {
      const r = await withRegistration(reading, () => stage2Conformance(manifest, REPO_ROOT, STAGE2_MEASURED))
      expect(r.violations, reading).toEqual([])
      expect([r.total, r.exact, r.withheld], reading).toEqual([base.total, base.exact, base.withheld])
    }
  }, 900_000)

  it('the conformance pass and cf-compare (interactive, paged) on this case: each planted world passes unregistered and under its own registration; the other reading is caught at its decisive cells', async () => {
    const { compareCfCase } = await import('../scripts/cf-compare')
    for (const reading of S2_FORMULA_BOUNDS_READINGS) {
      const { root, manifest: m } = syntheticRoot(plantedFills(caseOf(), reading))
      const c = m.cases.find(x => x.id === ID)!
      const before = await stage2Conformance(m, root, [ID])
      // Unregistered, the decisive cells production paints today (the painted-today group, and the near-top probes J2 / K2, an
      // endpoint under as-literal) are as-literal's: a fifteen-digit world is caught there, and only there.
      const painted = designated.filter(x => decisive(x) && !s2FormulaBoundsWithheldToday(x)).map(x => `${ID}:${x.ref}`)
      expect(painted.map(r => r.split(':')[1])).toEqual(['C2', 'C3', 'C4', 'C5', 'C6', 'J2', 'K2'])
      expect(before.violations.map(v => v.split(':').slice(0, 2).join(':')), reading).toEqual(reading === 'as-literal' ? [] : painted)
      expect([before.total, before.withheld], reading).toEqual([200, WITHHELD_TODAY])
      const registered = await withRegistration(reading, () => stage2Conformance(m, root, [ID]))
      expect(registered.violations, reading).toEqual([])
      expect([registered.total, registered.withheld], `${reading} registered`).toEqual([200, 0])
      // The main id alone: the cells near-top under this reading withheld and justified.
      const narrowed = await withRegistration(reading, () => stage2Conformance(m, root, [ID]), false)
      expect([narrowed.violations.length, narrowed.withheld], `${reading} main only`).toEqual([0, nearTopUnder(reading).length])
      for (const mode of ['interactive', 'paged'] as const) {
        const b = await withRegistration(reading, () => compareCfCase(c, mode, { rootDir: root }))
        expect(`${reading} registered ${mode}: ${b.outcome} ${b.outcome === 'pass' ? '' : b.reason}`).toBe(`${reading} registered ${mode}: pass `)
        if (reading === 'as-literal') {
          const a = await compareCfCase(c, mode, { rootDir: root })
          expect(`${reading} ${mode}: ${a.outcome} ${a.outcome === 'pass' ? '' : a.reason}`).toBe(`${reading} ${mode}: pass `)
        }
      }
      const other = S2_FORMULA_BOUNDS_READINGS.find(x => x !== reading)!
      const wrong = await withRegistration(other, () => stage2Conformance(m, root, [ID]))
      expect(wrong.violations.length, `${other} on a ${reading} world`).toBe(designated.filter(decisive).length)
    }
  }, 900_000)

  it('Stage 5: every measured case agrees with Excel under each registration; the only records that change are E4 and E5 of stage5-cfvo-nonnumeric-v2 (the formula-bounded upper integer class), newly painted at Excel\'s fill', async () => {
    const { stage5CaseRecords } = await import('../scripts/cf-refs/audit')
    const base = await stage5CaseRecords(manifest, REPO_ROOT)
    expect(base.filter(r => r.compared && !r.agrees)).toEqual([])
    const k = (r: (typeof base)[number]): string => `${r.caseId}:${r.ref}`
    const was = new Map(base.map(r => [k(r), r]))
    for (const reading of S2_FORMULA_BOUNDS_READINGS) {
      const rows = await withRegistration(reading, () => stage5CaseRecords(manifest, REPO_ROOT, { published: id => VERIFIED_CAPABILITIES.has(id) || id === CAP || id === NEAR }))
      expect(rows.length, reading).toBe(base.length)
      expect(rows.filter(r => r.compared && !r.agrees).map(r => `${r.caseId}:${r.ref} ${r.disagreement}`), reading).toEqual([])
      const changed = rows.filter(r => JSON.stringify(r.production) !== JSON.stringify(was.get(k(r))!.production))
      expect(changed.map(k), reading).toEqual(['stage5-cfvo-nonnumeric-v2:E4', 'stage5-cfvo-nonnumeric-v2:E5'])
      for (const r of changed) {
        expect(was.get(k(r))!.production.fill, k(r)).toBe('none')
        expect(r.production.fill, k(r)).toBe(r.excel?.fill)
      }
    }
  }, 900_000)
})

const sidecarFills = (): Map<string, string> => {
  const sc = JSON.parse(readFileSync(resolve(REPO_ROOT, caseOf().captures[0]!.sidecarPath), 'utf8')) as { observedCells: Array<{ ref: string; fillRgb: string }> }
  return new Map(sc.observedCells.map(o => [o.ref, o.fillRgb.toUpperCase()]))
}
const PERCENT = ['E2', 'E3', 'E4', 'E5', 'E6']

describe('the capture (Excel 16.115, d2660d46): no single reading matched the main question; nothing is registered', () => {
  it('verifyCfCase: verified; scale:formula-bounds refused (formula-hypothesis-none); scale:formula-bounds-near-top matched as-literal', () => {
    const res = verifyCfCase(caseOf(), { rootDir: REPO_ROOT })
    expect(res.kind).toBe('verified')
    if (res.kind !== 'verified') return
    expect(res.matchedHypotheses).toEqual({ [NEAR]: 'as-literal' })
    expect((res.refusedCapabilities ?? []).map(x => [x.capabilityId, x.code])).toEqual([[CAP, 'formula-hypothesis-none']])
  })

  it('sound: every pixel equals its DisplayFormat fill (200/200, minimum contrast 129); controls white; positive controls and all 93 twins as predicted', () => {
    const c = caseOf()
    const sc = JSON.parse(readFileSync(resolve(REPO_ROOT, c.captures[0]!.sidecarPath), 'utf8')) as { observedCells: Array<{ ref: string; fillRgb: string; bgRgb: number[]; bgContrastDelta: number }> }
    expect(sc.observedCells.length).toBe(200)
    for (const o of sc.observedCells) expect(o.bgRgb.map(x => x.toString(16).toUpperCase().padStart(2, '0')).join(''), o.ref).toBe(o.fillRgb.slice(2).toUpperCase())
    expect(Math.min(...sc.observedCells.filter(o => o.fillRgb !== 'FFFFFFFF').map(o => o.bgContrastDelta))).toBe(129)
    const fills = sidecarFills()
    for (const e of c.expected.cells.filter(x => !x.matched)) expect(fills.get(e.ref), e.ref).toBe('FFFFFFFF')
    for (const p of layout.positiveControls) expect(fills.get(p.ref), p.ref).toBe(p.token)
    for (const t of layout.twins) expect(fills.get(t.ref), t.ref).toBe(t.fill)
  })

  it('every designated cell: as-literal at every decisive cell but the five of f2-percent, where Excel drew fifteen-digit; every kept, clear and control cell as both predict', () => {
    const fills = sidecarFills()
    const byGroup: Record<string, string> = {}
    for (const cell of designated) {
      const ex = fills.get(cell.ref)!, a = s2FormulaBoundsPrediction(cell, 'as-literal'), b = s2FormulaBoundsPrediction(cell, 'fifteen-digit')
      const v = ex === a && ex === b ? 'both' : ex === a ? 'as-literal' : ex === b ? 'fifteen-digit' : 'none'
      expect(v === 'none', cell.ref).toBe(false)
      if (cell.role !== 'decisive') { expect(v, cell.ref).toBe('both'); continue }
      byGroup[cell.configuration] = byGroup[cell.configuration] === undefined || byGroup[cell.configuration] === v ? v : 'mixed'
    }
    expect(byGroup).toEqual({
      'f2-num-hi': 'as-literal', 'f2-formula-lo': 'as-literal', 'f2-painted': 'as-literal', 'f2-expr': 'as-literal', 'f2-percent': 'fifteen-digit',
      'f3-lower': 'as-literal', 'f3-upper': 'as-literal', 'f3-upper-mixed': 'as-literal', 'near-top': 'as-literal',
    })
    // The percent group (record only, never registered): its threshold as Excel drew it is also min + (max − min)·p/100 = 100
    // exactly (multiplication first), not production's lerp min + (p/100)·(max − min) = 100.00000000000001.
    expect([0 + (33.333333333333336 / 100) * 300, 0 + (300 * 33.333333333333336) / 100]).toEqual([100.00000000000001, 100])
  })

  it('production as committed (nothing registered): no violation over every measured Stage 2 case, this capture included; the 54 formula-bounded cells of the settled shapes withheld and justified', async () => {
    const r = await stage2Conformance(manifest, REPO_ROOT, [ID])
    expect(r.violations).toEqual([])
    expect([r.total, r.exact, r.withheld, r.upperWholeWithheld, r.guardBoundaryWithheld]).toEqual([200, 146, WITHHELD_TODAY, UPPER_WITHHELD, WITHHELD_TODAY - UPPER_WITHHELD])
    const all = await stage2Conformance(manifest, REPO_ROOT, [...STAGE2_MEASURED, ID])
    expect(all.violations).toEqual([])
  }, HEAVY)

  it('either reading registered would be caught: as-literal exactly at the five f2-percent cells, fifteen-digit at the 38 other decisive cells', async () => {
    const lit = await withRegistration('as-literal', () => stage2Conformance(manifest, REPO_ROOT, [ID]))
    expect(lit.violations.map(v => v.split(':')[1]).sort()).toEqual(PERCENT)
    const fif = await withRegistration('fifteen-digit', () => stage2Conformance(manifest, REPO_ROOT, [ID]))
    expect(fif.violations.map(v => v.split(':')[1]).sort()).toEqual(designated.filter(c => decisive(c) && !PERCENT.includes(c.ref)).map(c => c.ref).sort())
  }, HEAVY)
})

describe('a committed input survives regeneration', () => {
  it('keepCommittedArchive keeps the committed bytes when the inflated parts match', async () => {
    const { keepCommittedArchive, generateCaseArchive } = await import('../scripts/cf-refs/generate')
    const committed = new Uint8Array(readFileSync(resolve(REPO_ROOT, caseOf().input)))
    expect(keepCommittedArchive(committed, await generateCaseArchive(ID))).toBe(committed)
  }, HEAVY)
})

/** Pinned: the cells production withholds today (guard-boundary and upper shapes with a formula bound), and the text scope. */
const WITHHELD_TODAY = 54
const UPPER_WITHHELD = 23
const TEXT_SCOPE_BYTES = 4214842
