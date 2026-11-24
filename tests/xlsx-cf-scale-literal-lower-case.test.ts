// @vitest-environment node
/**
 * The literal-lower follow-up (brief cf-literal-lower), Phase 1 tasks 3-5: stage2-scale-literal-lower, the pre-registered and
 * registrable case for interior cells of a 3-colour LOWER segment bounded by two literal num cfvo where f·2^20 is not an
 * integer (withheld today as the num-segment shape), and the rehearsal of each reading's registration. Checked here:
 * 1. the case design: ownership, registries, size, values, every designated cell's class, >= 3 decisive cells per evidence
 *    configuration (span × colour order × palette kind), T4/V3's doubles in the lower segment, the upper probes' doubles, the
 *    twins' IEEE-integer products, the text scope's batch estimate;
 * 2. the independent model is production, today (the shape withheld) and under each registration (only the lower-segment
 *    evidence cells move; the upper probes stay withheld);
 * 3. the evaluator on synthetic observations: each planted reading comes back as the one match; the upper probes never
 *    decide; worlds no reading explains are refused; it fails closed on broken controls;
 * 4. the audit gate end to end on a synthetic measured capture (real verifyCfCase and auditStage2Manifest);
 * 5. each registration against every committed capture: lower-arithmetic paints the census cells exactly and keeps every
 *    other captured cell; step-plus-one is caught at the census cells; Stage 5 records unchanged; the conformance pass and
 *    cf-compare (interactive, paged) pass on each planted world, unregistered and under its own registration.
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
  S2_LITERAL_LOWER,
  S2_LITERAL_LOWER_CAP,
  S2_LITERAL_LOWER_ID,
  S2_LITERAL_LOWER_READINGS,
  S2_LITERAL_LOWER_TWIN,
  s2LiteralLowerCells,
  s2LiteralLowerClasses,
  s2LiteralLowerFraction,
  s2LiteralLowerInShape,
  s2LiteralLowerLayout,
  s2LiteralLowerPrediction,
  s2UwLine,
  type CfManifest,
  type CfOracleCase,
  type S2LiteralLowerCell,
  type S2LiteralLowerReading,
} from '../scripts/cf-refs/cases'
import { CF_MANIFEST_PATH } from '../scripts/cf-refs/generate'
import { computeWorksheetRangeGeometry, evaluateFormulaCapabilities, formulaControlsOf, isCellInSqref, readZipPartsSync, stage2ReadingsObservationOf, verifyCfCase } from '../scripts/cf-refs/verify'
import { literalLowerCensusOf } from '../scripts/cf-refs/literal-lower-census'
import { stage2Conformance } from './helpers/cf-stage2-conformance'
import { OfficePackage } from '../src/core/zip'
import { parseXlsx } from '../src/xlsx/parse'
import { cfDocumentEvaluationOf, prepareCfPaint } from '../src/xlsx/cf/render'
import { cfRuntimeEvidence } from '../src/xlsx/cf/predicates'
import { CF_LITERAL_LOWER_CAP, CF_LITERAL_LOWER_READINGS, injectCfLiteralLowerReadings, safeFraction } from '../src/xlsx/cf/colorScale'
import { estimateCfPrepareBytes } from '../src/xlsx/cf/service'
import { VERIFIED_CAPABILITIES, VERIFIED_STAGE2_CAPABILITIES, withExactCfEvidence } from '../src/xlsx/cf/evidence'
import { recordedBeforeUpperSegmentRegistration } from './helpers/cf-upper-segment-recorded'

// These tests record production as of their own captures, before the upper-segment registration (2ee3735a).
recordedBeforeUpperSegmentRegistration()

const REPO_ROOT = resolve(__dirname, '..')
const ID = S2_LITERAL_LOWER_ID
const CAP = S2_LITERAL_LOWER_CAP
const TWIN = S2_LITERAL_LOWER_TWIN
const HEAVY = 300_000

afterEach(() => injectCfLiteralLowerReadings(null))

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
const valuesOf = (c: CfOracleCase): Map<string, number> => {
  const xml = new TextDecoder().decode(partsOf(c, c.input).get('xl/worksheets/sheet1.xml')!)
  const out = new Map<string, number>()
  for (const m of xml.matchAll(/<c r="([A-Z]+\d+)"[^>]*><v>([^<]+)<\/v><\/c>/g)) out.set(m[1]!, Number(m[2]))
  return out
}
const layout = s2LiteralLowerLayout()
const twinFill = new Map(layout.twins.map(t => [t.ref, t.fill]))
const designated = s2LiteralLowerCells()
const cellAt = new Map(designated.map(c => [c.ref, c]))
/** Stage 2 conformance over every measured case, without the id's evidence and as committed (pinned from the run). */
const BEFORE = [1490, 1077, 413, 324, 59, 0]
const AFTER = [1490, 1191, 299, 210, 59, 0]
const decisive = (c: S2LiteralLowerCell): boolean => s2LiteralLowerPrediction(c, 'lower-arithmetic') !== s2LiteralLowerPrediction(c, 'step-plus-one')
const EVIDENCE = ['pl-pq-210', 'pl-qp-420', 'pl-pq-105', 'pl-qp-150', 'pl-resort-210', 'ds-yg-30', 'ds-dd-210', 'ds-da-105']
const evidenceRefs = designated.filter(c => c.role === 'evidence').map(c => c.ref)
const upperRefs = designated.filter(c => c.role === 'upper').map(c => c.ref)
/** Every designated cell of the num-segment shape (withheld unregistered): evidence, the 4 P-line below-integer controls and the upper probes. */
const inShapeRefs = designated.filter(c => s2LiteralLowerInShape(c)).map(c => c.ref)

/** Rehearse a registration: the reading selected through the seam and the id's evidence held, around `fn`. */
async function withRegistration<T>(reading: S2LiteralLowerReading, fn: () => Promise<T>): Promise<T> {
  injectCfLiteralLowerReadings({ [CF_LITERAL_LOWER_CAP]: reading })
  try { return await withExactCfEvidence([...VERIFIED_CAPABILITIES, CF_LITERAL_LOWER_CAP], fn) } finally { injectCfLiteralLowerReadings(null) }
}

/** The unpublished state: the registration in place but the id's evidence withheld, so the shape is withheld as before the capture. */
async function withoutEvidence<T>(fn: () => Promise<T>): Promise<T> {
  return withExactCfEvidence([...VERIFIED_CAPABILITIES].filter(id => id !== CF_LITERAL_LOWER_CAP), fn)
}

/** A synthetic world: each designated cell's fill by `pick` (default the named reading), twins as fixed, positive controls as fixed, controls white. */
function plantedFills(c: CfOracleCase, world: S2LiteralLowerReading | ((cell: S2LiteralLowerCell) => string)): Map<string, string> {
  const out = new Map<string, string>()
  for (const e of c.expected.cells) out.set(e.ref, e.matched ? e.fillRgb ?? '' : 'FFFFFFFF')
  for (const [ref, fill] of twinFill) out.set(ref, fill)
  for (const cell of designated) out.set(cell.ref, typeof world === 'string' ? s2LiteralLowerPrediction(cell, world) : world(cell))
  return out
}
const plantedSidecar = (c: CfOracleCase, fills: Map<string, string>): { observedCells: Array<Record<string, unknown>> } => ({
  observedCells: c.expected.cells.filter(e => fills.get(e.ref) !== undefined).map(e => ({ ref: e.ref, controlRef: e.controlRef, fillSource: 'excel-display-format', fillRgb: fills.get(e.ref) })),
})
const judge = (c: CfOracleCase, fills: Map<string, string>) => evaluateFormulaCapabilities(c.scaleReadings!, formulaControlsOf(c), stage2ReadingsObservationOf(plantedSidecar(c, fills)))

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

describe('stage2-scale-literal-lower: the case', () => {
  it('is appended last; owns its two ids; the question is registered as lower-arithmetic (its entry removed), the twins stay recorded', async () => {
    // Appended last at its time; the upper-segment follow-up is appended after it.
    expect(CF_STAGE2_CASES.map(c => c.id).slice(-3)).toEqual(['stage2-scale-raw-floor', ID, 'stage2-scale-upper-segment'])
    for (const id of [CAP, TWIN]) expect(CF_STAGE2_CAPABILITIES.filter(e => e.capabilityId === id).map(e => e.caseId), id).toEqual([ID])
    expect(CF_STAGE2_NONDISCRIMINATING[CAP]).toBeUndefined()
    expect(CF_STAGE2_NONDISCRIMINATING[TWIN]).toMatch(/./)
    expect(CAP).toBe(CF_LITERAL_LOWER_CAP)
    expect([...S2_LITERAL_LOWER_READINGS]).toEqual([...CF_LITERAL_LOWER_READINGS])
    const c = caseOf()
    expect(c.status).toBe('measured')
    expect(c.scaleReadings).toEqual(CF_STAGE2_CASES.find(t => t.id === ID)!.scaleReadings)
    expect(c.scaleReadings!.map(x => [x.id, x.capabilityId, Object.keys(x.hypotheses), x.role ?? (x.record ? 'record' : 'evidence')])).toEqual([
      ...EVIDENCE.map(id => [id, CAP, ['lower-arithmetic', 'step-plus-one'], 'evidence']),
      ['controls', CAP, ['lower-arithmetic', 'step-plus-one'], 'negative-control'], ['upper', CAP, ['lower-arithmetic', 'step-plus-one'], 'record'],
    ])
    const { CF_STAGE2_PRODUCTION_READINGS } = await import('../scripts/cf-refs/audit')
    // Registered: the question as lower-arithmetic (matched by a3f85309, published); the twins never.
    expect(CF_STAGE2_PRODUCTION_READINGS[CAP]).toBe('lower-arithmetic')
    expect(VERIFIED_CAPABILITIES.has(CAP)).toBe(true)
    expect(Object.prototype.hasOwnProperty.call(CF_STAGE2_PRODUCTION_READINGS, TWIN)).toBe(false)
    expect(VERIFIED_CAPABILITIES.has(TWIN)).toBe(false)
  })

  it('180 cells: 82 designated (46 evidence, 20 controls, 16 upper probes), 86 twins, 2 positive controls, 10 controls; A1:R13, 31 colour rules', () => {
    const c = caseOf()
    expect(c.range).toBe('A1:R13')
    expect(c.expected.cells.length).toBe(180)
    expect(c.scaleReadings!.map(x => x.refs.length)).toEqual([6, 6, 6, 6, 6, 5, 5, 6, 20, 16])
    expect([designated.length, evidenceRefs.length, upperRefs.length]).toEqual([82, 46, 16])
    expect(layout.twins.length).toBe(86)
    expect(c.expected.cells.filter(e => e.matched && !e.scaleApplied).length).toBe(2)
    expect(c.expected.negativeRefs.length).toBe(10)
    expect(c.markers.length).toBe(31)
    for (const e of c.expected.cells) expect(c.markers.filter(m => isCellInSqref(e.ref, m.sqref)).length, e.ref).toBeLessThanOrEqual(1)
    for (const ref of c.scaleReadings!.flatMap(x => x.refs)) expect(c.expected.cells.find(e => e.ref === ref)!.controlRef, ref).toBe(`${colOf(ref)}13`)
    for (const ref of c.expected.negativeRefs) expect(c.markers.some(m => isCellInSqref(ref, m.sqref)), ref).toBe(false)
    for (const m of c.markers) expect(m.sqref, m.sqref).not.toMatch(/^([A-Z]+\d+):\1$/)
    // Every block rule is three literal num cfvo (the shape); every twin rule is 2-colour num 0 / 1.
    for (const m of c.markers) {
      if (m.capabilityId === CAP) expect(m.scale!.cfvo.map(v => v.type), m.sqref).toEqual(['num', 'num', 'num'])
      else expect([m.capabilityId, m.scale!.cfvo], m.sqref).toEqual([TWIN, [{ type: 'num', val: '0' }, { type: 'num', val: '1' }]])
    }
  })

  it('the text scope admits the whole range in one batch, well under the 8,388,608-byte cap', async () => {
    const { sheet } = await sheetOf(caseOf())
    const bytes = estimateCfPrepareBytes(sheet.conditionalFormatting!, [{ firstRow: 0, lastRow: 12, firstCol: 0, lastCol: 17 }])
    expect(bytes).toBe(3339576)
    expect(bytes!).toBeLessThan(8388608 * 0.45)
  })

  it('the workbook: the catalogue values, twins at their values, positive controls on their dxf fills; no hidden rows, merges or errors', () => {
    const c = caseOf()
    const values = valuesOf(c)
    for (const cell of designated) expect(values.get(cell.ref), cell.ref).toBe(Number(cell.text))
    for (const t of layout.twins) expect(values.get(t.ref), t.ref).toBe(Number(t.text))
    const xml = new TextDecoder().decode(partsOf(c, c.input).get('xl/worksheets/sheet1.xml')!)
    expect(xml).not.toMatch(/hidden="1"|<mergeCell|<autoFilter|t="e"/)
    const styles = new TextDecoder().decode(partsOf(c, c.input).get('xl/styles.xml')!)
    for (const p of layout.positiveControls) {
      expect(xml, p.ref).toContain(`<conditionalFormatting sqref="${p.ref}"><cfRule type="cellIs" dxfId="`)
      expect(values.get(p.ref), p.ref).toBe(1)
      expect(styles, p.token).toContain(`<bgColor rgb="${p.token}"/>`)
    }
    for (const e of c.expected.cells) expect(typeof values.get(e.ref), e.ref).toBe('number')
  })

  it('every designated cell is in its designed role; each evidence configuration separates the readings at >= 5 cells; the classes and palette kinds (pinned)', () => {
    for (const c of designated) {
      const f = s2LiteralLowerFraction(c)
      // The model's fraction is production's double.
      expect(Object.is(f, safeFraction(c.value, c.lo, c.hi)), c.ref).toBe(true)
      if (c.role === 'evidence') { expect([c.segment, s2LiteralLowerInShape(c), decisive(c)], c.ref).toEqual([0, true, true]) }
      else if (c.role === 'control') { expect([c.segment, decisive(c)], c.ref).toEqual([0, false]) }
      else { expect([c.segment, s2LiteralLowerInShape(c)], c.ref).toEqual([1, true]) }
      for (const r of S2_LITERAL_LOWER_READINGS) if (c.role !== 'upper') expect(s2LiteralLowerPrediction(c, r), `${c.ref} ${r}`).not.toBe('none')
    }
    const dec = (id: string): number => { const cfg = layout.configurations.find(x => x.id === id)!; return cfg.refs.filter(r => cfg.hypotheses['lower-arithmetic']![r] !== cfg.hypotheses['step-plus-one']![r]).length }
    expect(Object.fromEntries([...EVIDENCE, 'controls'].map(id => [id, dec(id)]))).toEqual({ 'pl-pq-210': 6, 'pl-qp-420': 6, 'pl-pq-105': 6, 'pl-qp-150': 6, 'pl-resort-210': 6, 'ds-yg-30': 5, 'ds-dd-210': 5, 'ds-da-105': 6, controls: 0 })
    // The worst channel class of each evidence cell, per palette kind: P-line above / integer / clear (its below-integer cells
    // coincide under both readings: controls); distinct below / above / integer / clear.
    const RANK = { below: 3, above: 2, integer: 1, clear: 0 } as const
    const worst = (c: S2LiteralLowerCell): string => s2LiteralLowerClasses(c).reduce((a, b) => (RANK[b] > RANK[a] ? b : a), 'clear' as keyof typeof RANK)
    const byKind = (kind: string, role: string): Record<string, number> => {
      const out: Record<string, number> = {}
      for (const c of designated.filter(x => x.block.kind === kind && x.role === role)) out[worst(c)] = (out[worst(c)] ?? 0) + 1
      return out
    }
    expect(byKind('p-line', 'evidence')).toEqual({ above: 12, integer: 11, clear: 7 })
    expect(byKind('distinct', 'evidence')).toEqual({ below: 3, above: 4, integer: 6, clear: 3 })
    expect(byKind('p-line', 'control')).toEqual({ below: 4, integer: 10, clear: 1 })
    expect(byKind('distinct', 'control')).toEqual({ integer: 5 })
    // Spans, bounds and colour orders of the evidence configurations.
    const spans = [...new Set(designated.filter(c => c.role === 'evidence').map(c => `${c.block.id} ${c.lo}..${c.hi}`))]
    expect(spans).toEqual(['pl-pq-210 7..217', 'pl-qp-420 -210..210', 'pl-pq-105 -52.5..52.5', 'pl-qp-150 0.5..150.5', 'pl-resort-210 0..210', 'ds-yg-30 0..30', 'ds-dd-210 -105..105', 'ds-da-105 -15..90'])
    const dirs = (c: S2LiteralLowerCell): string => [2, 4, 6].map(i => (Number.parseInt(c.to.slice(i, i + 2), 16) > Number.parseInt(c.from.slice(i, i + 2), 16) ? '+' : '-')).join('')
    expect(Object.fromEntries(EVIDENCE.map(id => [id, dirs(designated.find(c => c.configuration === id)!)]))).toEqual({
      'pl-pq-210': '+-+', 'pl-qp-420': '-+-', 'pl-pq-105': '+-+', 'pl-qp-150': '-+-', 'pl-resort-210': '+-+', 'ds-yg-30': '---', 'ds-dd-210': '---', 'ds-da-105': '+++',
    })
    // Every f·2^20-integer and endpoint control: both readings at production's truncation.
    expect(designated.filter(c => c.role === 'control' && !s2LiteralLowerInShape(c)).length).toBe(16)
  })

  it('T4/V3 in the lower segment: ds-yg-30 F2 and F3 are T4 and V3\'s doubles and colours, the lower-arithmetic and step-plus-one fills their truncation and their Excel fills; the upper probes F10, F11 share T4/V3\'s values and doubles but run the other way (G -> Y)', () => {
    const f = (c: S2LiteralLowerCell): number => safeFraction(c.value, c.lo, c.hi)!
    // T4 and V3 (e448d831): upper segment of num 0 / 30 / 60 at 40 and 50, FFFFEB84 -> FF63BE7B.
    const t4 = safeFraction(40, 30, 60)!, v3 = safeFraction(50, 30, 60)!
    const [F2, F3, F10, F11] = ['F2', 'F3', 'F10', 'F11'].map(r => cellAt.get(r)!)
    expect([Object.is(f(F2!), t4), Object.is(f(F3!), v3)]).toEqual([true, true])
    expect([F2!.from, F2!.to]).toEqual(['FFFFEB84', 'FF63BE7B'])
    expect(S2_LITERAL_LOWER_READINGS.map(r => s2LiteralLowerPrediction(F2!, r))).toEqual(['FFCBDC81', 'FFCCDD82'])
    expect(S2_LITERAL_LOWER_READINGS.map(r => s2LiteralLowerPrediction(F3!, r))).toEqual(['FF97CD7E', 'FF98CE7F'])
    // The upper probes at 40 and 50 have T4/V3's values and doubles, but their segment runs G -> Y (T4/V3's ran Y -> G): not replicas.
    expect([F10!.value, F11!.value, F10!.segment, F11!.segment]).toEqual([40, 50, 1, 1])
    expect([Object.is(f(F10!), t4), Object.is(f(F11!), v3), F10!.from, F10!.to]).toEqual([true, true, 'FF63BE7B', 'FFFFEB84'])
    // 1/3 and 2/3 also on the P line (pl-pq-210 A5 and pl-resort-210 E5) and ascending (ds-da-105 H6 at 70/105).
    expect(Object.is(f(cellAt.get('A5')!), 70 / 210)).toBe(true)
    expect(Object.is(f(cellAt.get('E5')!), 70 / 210)).toBe(true)
  })

  it('the upper probes: each has a lower-segment evidence partner in its block with the same f double bit for bit; recorded only', () => {
    const f = (c: S2LiteralLowerCell): number => safeFraction(c.value, c.lo, c.hi)!
    const partners = upperRefs.map(ref => {
      const u = cellAt.get(ref)!
      const p = designated.find(c => c.role === 'evidence' && c.block === u.block && Object.is(f(c), f(u)))
      return `${ref}:${p?.ref ?? '-'}`
    })
    expect(partners).toEqual(['A11:A2', 'A12:A5', 'B11:B2', 'B12:B5', 'C11:C2', 'C12:C6', 'D11:D2', 'D12:D5', 'E11:E2', 'E12:E5', 'F10:F2', 'F11:F3', 'G8:G2', 'G9:G5', 'H9:H2', 'H10:H5'])
    expect(layout.configurations.find(x => x.id === 'upper')!.record).toMatch(/never evidence/)
  })

  it('the twins: every live prediction is a twin (or a literal positive control in the controls), and every twin\'s product is an IEEE integer on every channel (recomputed)', () => {
    for (const t of layout.twins) {
      const [from, to] = t.rule.colors
      const v = Number(t.text), f = (v - 0) / (1 - 0)
      for (const i of [0, 2, 4, 6]) {
        const d = Math.abs(Number.parseInt(to.slice(i, i + 2), 16) - Number.parseInt(from.slice(i, i + 2), 16))
        expect(Number.isInteger(f * d), `${t.ref} ${f * d}`).toBe(true)
      }
      // The twin shows its fill: at f = 0 its rule's first colour, at f = 1 its second, on the line the P-line step.
      const line = S2_LITERAL_LOWER.twinSpans.some(sp => from === s2UwLine(sp.from) && to === s2UwLine(sp.to))
      if (!line) expect(t.fill, t.ref).toBe(v === 0 ? from : to)
    }
    for (const cfg of layout.configurations.filter(x => x.record === undefined)) {
      for (const h of Object.values(cfg.hypotheses)) for (const [ref, token] of Object.entries(h)) {
        if (token.startsWith('twin:')) expect(twinFill.get(token.slice(5)), `${cfg.id} ${ref}`).toBe(s2LiteralLowerPrediction(cellAt.get(ref)!, Object.keys(cfg.hypotheses).find(k => cfg.hypotheses[k] === h) as S2LiteralLowerReading))
        else { expect(cfg.id, ref).toBe('controls'); expect(Object.values(cfg.positiveControls ?? {}), ref).toContain(token) }
      }
    }
    expect(S2_LITERAL_LOWER.twinSpans).toEqual([{ from: 0, to: 128 }, { from: 82, to: 210 }])
  })

  it('the independent model is production without the evidence (as before the capture): painted cells are the truncation; every cell of the shape (evidence, the 4 below-integer controls, upper probes) is withheld', async () => {
    const c = caseOf()
    const fills = await withoutEvidence(() => productionFills(c, c.expected.cells.map(e => e.ref)))
    const withheld: string[] = []
    for (const e of c.expected.cells) {
      const got = fills.get(e.ref)
      if (got === undefined && e.matched) { withheld.push(e.ref); continue }
      const cell = cellAt.get(e.ref)
      expect(got, e.ref).toBe(!e.matched ? undefined : cell ? s2LiteralLowerPrediction(cell, 'lower-arithmetic') : twinFill.get(e.ref) ?? e.fillRgb)
    }
    expect(withheld.sort()).toEqual(inShapeRefs.sort())
    expect(inShapeRefs.length).toBe(46 + 4 + 16)
    // With the id's evidence but no valid reading registered (the pre-capture table): the same.
    injectCfLiteralLowerReadings({ [CF_LITERAL_LOWER_CAP]: 'unregistered' as S2LiteralLowerReading })
    try { expect(await productionFills(c, c.expected.cells.map(e => e.ref))).toEqual(fills) } finally { injectCfLiteralLowerReadings(null) }
  }, HEAVY)

  it('the independent model is production as committed (lower-arithmetic registered and published): Excel\'s fill at every cell production paints; only the 16 upper probes are withheld', async () => {
    const c = caseOf()
    const fills = await productionFills(c, c.expected.cells.map(e => e.ref))
    const sidecar = JSON.parse(readFileSync(resolve(REPO_ROOT, c.captures[0]!.sidecarPath), 'utf8')) as { observedCells: Array<{ ref: string; fillRgb: string }> }
    const excel = new Map(sidecar.observedCells.map(o => [o.ref, o.fillRgb]))
    const withheld: string[] = []
    for (const e of c.expected.cells) {
      const got = fills.get(e.ref)
      if (got === undefined && e.matched) { withheld.push(e.ref); continue }
      const cell = cellAt.get(e.ref)
      expect(got, e.ref).toBe(!e.matched ? undefined : cell ? s2LiteralLowerPrediction(cell, 'lower-arithmetic') : twinFill.get(e.ref) ?? e.fillRgb)
      if (e.matched) expect(got, `${e.ref} Excel`).toBe(excel.get(e.ref))
    }
    expect(withheld.sort()).toEqual([...upperRefs].sort())
  }, HEAVY)

  for (const reading of S2_LITERAL_LOWER_READINGS) {
    it(`the independent model is production under the ${reading} registration: every evidence cell paints the reading; the upper probes stay withheld; nothing else moves`, async () => {
      const c = caseOf()
      const fills = await withRegistration(reading, () => productionFills(c, c.expected.cells.map(e => e.ref)))
      for (const e of c.expected.cells) {
        const cell = cellAt.get(e.ref)
        const want = !e.matched || cell?.role === 'upper' ? undefined : cell ? s2LiteralLowerPrediction(cell, reading) : twinFill.get(e.ref) ?? e.fillRgb
        expect(fills.get(e.ref), e.ref).toBe(want)
      }
    }, HEAVY)
  }
})

describe('stage2-scale-literal-lower: the evaluator on synthetic observations', () => {
  for (const reading of S2_LITERAL_LOWER_READINGS) {
    it(`a planted ${reading} world matches exactly ${reading}; the upper record never decides`, () => {
      const c = caseOf()
      const r = judge(c, plantedFills(c, reading))
      expect([r[CAP]!.status, r[CAP]!.matched]).toEqual(['discriminating', reading])
      const other = S2_LITERAL_LOWER_READINGS.find(x => x !== reading)!
      for (const up of [(x: S2LiteralLowerCell) => s2LiteralLowerPrediction(x, other), () => 'FF000000']) {
        const w = judge(c, plantedFills(c, x => (x.role === 'upper' ? up(x) : s2LiteralLowerPrediction(x, reading))))
        expect([w[CAP]!.status, w[CAP]!.matched], reading).toEqual(['discriminating', reading])
      }
    })
  }

  it('no reading matches: a world at one reading except one decisive cell of one configuration at the other is refused', () => {
    const c = caseOf()
    for (const [base, flip] of [['lower-arithmetic', 'step-plus-one'], ['step-plus-one', 'lower-arithmetic']] as const) {
      for (const cfg of EVIDENCE) {
        const target = designated.find(x => x.configuration === cfg && decisive(x))!
        const r = judge(c, plantedFills(c, x => s2LiteralLowerPrediction(x, x === target ? flip : base)))
        expect([r[CAP]!.status, r[CAP]!.code], `${base} with ${target.ref} at ${flip}`).toEqual(['refused', 'formula-hypothesis-none'])
      }
    }
  })

  it('fails closed: a formatted control, a missing or wrong positive control, a twin rendering nothing, a moved negative control', () => {
    const c = caseOf()
    for (const reading of S2_LITERAL_LOWER_READINGS) {
      const fills = plantedFills(c, reading)
      const edit = (ref: string, fill: string | undefined): Map<string, string> => { const m = new Map(fills); if (fill === undefined) m.delete(ref); else m.set(ref, fill); return m }
      const pc = layout.positiveControls[0]!.ref
      expect(judge(c, edit('A13', 'FF14F028'))[CAP]!.status, reading).toBe('refused')
      expect(judge(c, edit(pc, undefined))[CAP]!.code, reading).toBe('fill-unobserved')
      expect(judge(c, edit(pc, 'FF000000'))[CAP]!.code, reading).toBe('formula-positive-control-failed')
      expect(judge(c, edit([...twinFill.keys()][0]!, 'FFFFFFFF'))[CAP]!.status, reading).toBe('refused')
      // A below-integer P-line control (A8, 34 in 7 .. 217) drawn one step larger (guarded truncation): nothing matches.
      const moved = judge(c, edit('A8', 'FF2FD543'))
      expect([moved[CAP]!.status, moved[CAP]!.code], reading).toEqual(['refused', 'formula-hypothesis-none'])
    }
  })
})

// ---- a synthetic measured capture through the real verifier, audit, conformance pass and cf-compare (temp root) ----
const sha = (b: Uint8Array): string => createHash('sha256').update(b).digest('hex')

function syntheticRoot(fills: Map<string, string>): { root: string; manifest: CfManifest } {
  const root = mkdtempSync(join(tmpdir(), 'cf-s2-literal-lower-'))
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
  const at = '2026-11-24T00:00:00.000Z'
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
  // The verified set without the question, and the registry with its pre-capture entry put back (the rehearsal's starting point).
  const before = [...VERIFIED_STAGE2_CAPABILITIES].filter(id => id !== CAP).sort()
  const removed = CF_STAGE2_NONDISCRIMINATING
  const kept = { ...CF_STAGE2_NONDISCRIMINATING, [CAP]: 'pre-registered, awaiting its capture' }

  it('the real manifest: Stage 2 audits ok and complete; the question verified (lower-arithmetic), the twins unverified-nondiscriminating; the verified set gains exactly the question', async () => {
    const { auditStage2Manifest } = await import('../scripts/cf-refs/audit')
    const r = auditStage2Manifest(manifest, { rootDir: REPO_ROOT })
    expect([r.ok, r.stageComplete]).toEqual([true, true])
    expect(disposition(r, CAP)).toBe('verified')
    expect(disposition(r, TWIN)).toBe('unverified-nondiscriminating')
    expect(r.perCapability.filter(p => p.disposition === 'verified').map(p => p.capabilityId).sort()).toEqual([...before, CAP].sort())
    const k = auditStage2Manifest(manifest, { rootDir: REPO_ROOT, nonDiscriminating: kept })
    expect(disposition(k, CAP)).toBe('unverified-nondiscriminating')
    expect(k.perCapability.filter(p => p.disposition === 'verified').map(p => p.capabilityId).sort()).toEqual(before)
  }, HEAVY)

  for (const reading of S2_LITERAL_LOWER_READINGS) {
    it(`planted ${reading}: matched itself; registry kept: unverified; entry removed, unregistered: withheld; registered: verified; the other reading registered: withheld`, async () => {
      const { auditStage2Manifest, CF_STAGE2_PRODUCTION_READINGS } = await import('../scripts/cf-refs/audit')
      const { root, manifest: m } = syntheticRoot(plantedFills(caseOf(), reading))
      const res = verifyCfCase(m.cases.find(x => x.id === ID)!, { rootDir: root })
      expect(res.kind, res.kind === 'refusal' ? `${res.code}: ${res.reason}` : '').toBe('verified')
      if (res.kind !== 'verified') return
      expect(res.matchedHypotheses).toEqual({ [CAP]: reading })
      const k = auditStage2Manifest(m, { rootDir: root, nonDiscriminating: kept })
      expect(k.ok).toBe(true)
      for (const id of [CAP, TWIN]) expect(disposition(k, id), id).toBe('unverified-nondiscriminating')
      const unregisteredReadings = Object.fromEntries(Object.entries(CF_STAGE2_PRODUCTION_READINGS).filter(([id]) => id !== CAP))
      expect(disposition(auditStage2Manifest(m, { rootDir: root, nonDiscriminating: removed, productionReadings: unregisteredReadings }), CAP)).toBe('supported-structural')
      const registered = auditStage2Manifest(m, { rootDir: root, nonDiscriminating: removed, productionReadings: { ...CF_STAGE2_PRODUCTION_READINGS, [CAP]: reading } })
      expect(disposition(registered, CAP)).toBe('verified')
      expect(registered.perCapability.filter(p => p.disposition === 'verified').map(p => p.capabilityId).sort()).toEqual([...before, CAP].sort())
      const other = S2_LITERAL_LOWER_READINGS.find(x => x !== reading)!
      expect(disposition(auditStage2Manifest(m, { rootDir: root, nonDiscriminating: removed, productionReadings: { ...CF_STAGE2_PRODUCTION_READINGS, [CAP]: other } }), CAP)).toBe('supported-structural')
    }, HEAVY)
  }

  it('planted "no reading matches" (lower-arithmetic with one cell per evidence configuration at step-plus-one): formula-hypothesis-none, never verified, and production keeps the withhold', async () => {
    const { auditStage2Manifest, CF_STAGE2_PRODUCTION_READINGS } = await import('../scripts/cf-refs/audit')
    const flipped = new Set(EVIDENCE.map(cfg => designated.find(x => x.configuration === cfg && decisive(x))!.ref))
    const { root, manifest: m } = syntheticRoot(plantedFills(caseOf(), x => s2LiteralLowerPrediction(x, flipped.has(x.ref) ? 'step-plus-one' : 'lower-arithmetic')))
    const res = verifyCfCase(m.cases.find(x => x.id === ID)!, { rootDir: root })
    expect(res.kind).toBe('verified')
    if (res.kind !== 'verified') return
    expect((res.refusedCapabilities ?? []).map(x => [x.capabilityId, x.code])).toEqual([[CAP, 'formula-hypothesis-none']])
    for (const reading of S2_LITERAL_LOWER_READINGS) expect(disposition(auditStage2Manifest(m, { rootDir: root, nonDiscriminating: removed, productionReadings: { ...CF_STAGE2_PRODUCTION_READINGS, [CAP]: reading } }), CAP), reading).toBe('supported-structural')
    // Unpublished production on that world: every cell of the shape withheld and justified (num-segment); nothing disagrees.
    const r = await withoutEvidence(() => stage2Conformance(m, root, [ID]))
    expect(r.violations).toEqual([])
    expect([r.total, r.withheld, r.numSegmentWithheld, r.upperWholeWithheld, r.guardBoundaryWithheld]).toEqual([180, 66, 66, 0, 0])
    // As committed (lower-arithmetic registered and published), the conformance pass catches the 8 flipped cells.
    const committed = await stage2Conformance(m, root, [ID])
    expect(committed.violations.map(v => v.split(':').slice(0, 2).join(':')).sort()).toEqual([...flipped].map(x => `${ID}:${x}`).sort())
  }, HEAVY)
})

const STAGE2_MEASURED = manifest.cases.filter(x => x.stage === 2 && x.status === 'measured' && x.id !== 'stage2-scale-upper-segment').map(x => x.id)

describe('the rehearsal: each registration against every committed capture, and this case on a synthetic capture', () => {
  it('without the evidence (as before the registration): no violation over the measured Stage 2 cases, this capture included; the literal lower-segment cells of the shape withheld', async () => {
    expect(STAGE2_MEASURED).toContain(ID)
    const r = await withoutEvidence(() => stage2Conformance(manifest, REPO_ROOT, STAGE2_MEASURED))
    expect(r.violations).toEqual([])
    expect([r.total, r.exact, r.withheld, r.numSegmentWithheld, r.upperWholeWithheld, r.guardBoundaryWithheld]).toEqual(BEFORE)
  }, HEAVY)

  it('as committed (lower-arithmetic registered and published): every committed capture agrees; the 114 census cells of the shape (64 earlier, 50 of this capture) painted at Excel\'s fill; nothing else moves', async () => {
    const before = await withoutEvidence(() => stage2Conformance(manifest, REPO_ROOT, STAGE2_MEASURED))
    const r = await stage2Conformance(manifest, REPO_ROOT, STAGE2_MEASURED)
    expect(r.violations).toEqual([])
    expect([r.total, r.exact, r.withheld, r.numSegmentWithheld, r.upperWholeWithheld, r.guardBoundaryWithheld]).toEqual(AFTER)
    expect([r.exact - before.exact, before.numSegmentWithheld - r.numSegmentWithheld]).toEqual([114, 114])
    expect(await withRegistration('lower-arithmetic', () => stage2Conformance(manifest, REPO_ROOT, STAGE2_MEASURED))).toEqual(r)
    const { rows } = await literalLowerCensusOf()
    const census = rows.filter(x => !x.f20)
    expect([census.length, census.filter(x => x.caseId === ID).length]).toEqual([114, 50])
    const byCase = new Map<string, string[]>()
    for (const x of census) byCase.set(x.caseId, [...(byCase.get(x.caseId) ?? []), x.ref])
    for (const [caseId, refs] of byCase) {
      const c = manifest.cases.find(x => x.id === caseId)!
      const fills = await productionFills(c, refs)
      for (const x of census.filter(y => y.caseId === caseId)) expect(fills.get(x.ref), `${caseId}:${x.ref}`).toBe(x.excel)
    }
  }, 900_000)

  it('step-plus-one registered: the committed captures refute it at the census cells where the readings differ (caught, never published)', async () => {
    const r = await withRegistration('step-plus-one', () => stage2Conformance(manifest, REPO_ROOT, STAGE2_MEASURED))
    const { rows } = await literalLowerCensusOf()
    const census = rows.filter(x => !x.f20)
    // The below-integer P-line census cells (raw-floor F6..F9, this capture's 4 controls) coincide under both readings.
    const differs = census.filter(x => !(x.kind === 'p-line' && x.cls === 'below'))
    expect([census.length, differs.length]).toEqual([114, 106])
    expect(r.violations.map(v => v.split(':').slice(0, 2).join(':')).sort()).toEqual(differs.map(x => `${x.caseId}:${x.ref}`).sort())
  }, HEAVY)

  it('the conformance pass and cf-compare (interactive, paged) on this case: each planted world passes unregistered and under its own registration; the other reading is caught', async () => {
    const { compareCfCase } = await import('../scripts/cf-compare')
    for (const reading of S2_LITERAL_LOWER_READINGS) {
      const { root, manifest: m } = syntheticRoot(plantedFills(caseOf(), reading))
      const c = m.cases.find(x => x.id === ID)!
      const unregistered = await withoutEvidence(() => stage2Conformance(m, root, [ID]))
      expect(unregistered.violations, reading).toEqual([])
      expect([unregistered.total, unregistered.withheld, unregistered.numSegmentWithheld], reading).toEqual([180, 66, 66])
      const registered = await withRegistration(reading, () => stage2Conformance(m, root, [ID]))
      expect(registered.violations, reading).toEqual([])
      expect([registered.total, registered.withheld, registered.numSegmentWithheld, registered.upperWholeWithheld, registered.guardBoundaryWithheld], `${reading} registered`).toEqual([180, 16, 16, 0, 0])
      for (const mode of ['interactive', 'paged'] as const) {
        const a = await withoutEvidence(() => compareCfCase(c, mode, { rootDir: root }))
        expect(`${reading} ${mode}: ${a.outcome} ${a.outcome === 'pass' ? '' : a.reason}`).toBe(`${reading} ${mode}: pass `)
        const b = await withRegistration(reading, () => compareCfCase(c, mode, { rootDir: root }))
        expect(`${reading} registered ${mode}: ${b.outcome} ${b.outcome === 'pass' ? '' : b.reason}`).toBe(`${reading} registered ${mode}: pass `)
      }
      const other = S2_LITERAL_LOWER_READINGS.find(x => x !== reading)!
      const wrong = await withRegistration(other, () => stage2Conformance(m, root, [ID]))
      expect(wrong.violations.length, `${other} on a ${reading} world`).toBe(46)
    }
  }, 900_000)

  it('Stage 5: every measured case agrees with Excel, records unchanged, under each registration (no Stage 5 cell is a literal-num lower-segment cell)', async () => {
    const { stage5CaseRecords } = await import('../scripts/cf-refs/audit')
    const base = await stage5CaseRecords(manifest, REPO_ROOT)
    expect(base.filter(r => r.compared && !r.agrees)).toEqual([])
    const key = (rows: typeof base) => rows.map(r => `${r.caseId}:${r.ref}:${r.agrees}:${JSON.stringify(r.production)}`).join('\n')
    for (const reading of S2_LITERAL_LOWER_READINGS) {
      const rows = await withRegistration(reading, () => stage5CaseRecords(manifest, REPO_ROOT, { published: id => VERIFIED_CAPABILITIES.has(id) || id === CAP }))
      expect(rows.filter(r => r.compared && !r.agrees).map(r => `${r.caseId}:${r.ref} ${r.disagreement}`), reading).toEqual([])
      expect(key(rows) === key(base), `${reading}: Stage 5 records unchanged`).toBe(true)
    }
  }, 900_000)
})

describe('a committed input survives regeneration', () => {
  it('keepCommittedArchive keeps the committed bytes when the inflated parts match', async () => {
    const { keepCommittedArchive, generateCaseArchive } = await import('../scripts/cf-refs/generate')
    const committed = new Uint8Array(readFileSync(resolve(REPO_ROOT, caseOf().input)))
    expect(keepCommittedArchive(committed, await generateCaseArchive(ID))).toBe(committed)
  }, HEAVY)
})
