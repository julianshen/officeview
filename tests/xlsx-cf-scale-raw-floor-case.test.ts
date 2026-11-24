// @vitest-environment node
/**
 * The raw-floor follow-up (brief cf-rawfloor), Phase 1 tasks 3-5: stage2-scale-raw-floor, the pre-registered case for the
 * guard-boundary shape of 2-colour scales and 3-colour lower segments (W1, docs/validation/xlsx-cf.md §14.z.17), and the
 * rehearsal of each reading's registration. Checked here:
 * 1. the case design: ownership, registries, size, values, statistical thresholds bit for bit, every designated cell's class,
 *    the 14 doubles of ed0fccb6 replicated, >= 3 decisive cells per evidence configuration, the W3 probes' doubles, the twin
 *    line's IEEE-integer products, the text scope's batch estimate;
 * 2. the independent model is production, today (the shape withheld) and under each registration;
 * 3. the evaluator on synthetic observations: each planted reading comes back as the one match; the W3 probes never decide;
 *    worlds no reading explains are refused; it fails closed on broken controls;
 * 4. the audit gate end to end on a synthetic measured capture (real verifyCfCase and auditStage2Manifest);
 * 5. each registration against every committed capture: raw-floor paints the 14 cells of ed0fccb6 exactly and keeps every
 *    other captured cell exact; guarded is refuted at exactly those 14; Stage 5 records unchanged; the conformance pass and
 *    cf-compare (interactive, paged) pass on each planted world, unregistered and under its own registration.
 * This checks the case and production, not Excel. After the capture (ce3f701d, Excel 16.115) matched raw-floor, raw-floor is
 * registered (docs/validation/xlsx-cf.md §14.z.18): the pins below are the registered state, and the unregistered state is
 * rehearsed by withholding the id's evidence (`withoutEvidence`).
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
  S2_RAW_FLOOR,
  S2_RAW_FLOOR_CAP,
  S2_RAW_FLOOR_ID,
  S2_RAW_FLOOR_LITERAL,
  S2_RAW_FLOOR_READINGS,
  S2_RAW_FLOOR_TWIN,
  s2RawFloorBlocks,
  s2RawFloorCells,
  s2RawFloorClasses,
  s2RawFloorLayout,
  s2RawFloorPrediction,
  s2UwLine,
  type CfManifest,
  type CfOracleCase,
  type S2RawFloorCell,
  type S2RawFloorReading,
} from '../scripts/cf-refs/cases'
import { CF_MANIFEST_PATH } from '../scripts/cf-refs/generate'
import { computeWorksheetRangeGeometry, evaluateFormulaCapabilities, formulaControlsOf, isCellInSqref, readZipPartsSync, stage2ReadingsObservationOf, verifyCfCase } from '../scripts/cf-refs/verify'
import { stage2Conformance } from './helpers/cf-stage2-conformance'
import { OfficePackage } from '../src/core/zip'
import { parseXlsx } from '../src/xlsx/parse'
import { cfDocumentEvaluationOf, prepareCfPaint } from '../src/xlsx/cf/render'
import { cfRuntimeEvidence } from '../src/xlsx/cf/predicates'
import { CF_GUARD_BOUNDARY_CAP, CF_GUARD_BOUNDARY_READINGS, injectCfGuardBoundaryReadings, resolveScaleThresholds, safeFraction } from '../src/xlsx/cf/colorScale'
import { estimateCfPrepareBytes } from '../src/xlsx/cf/service'
import { VERIFIED_CAPABILITIES, VERIFIED_STAGE2_CAPABILITIES, withExactCfEvidence } from '../src/xlsx/cf/evidence'
import type { CfScalePoint } from '../src/xlsx/cf/model'
import { recordedBeforeUpperSegmentRegistration } from './helpers/cf-upper-segment-recorded'

// These tests record production as of their own captures, before the upper-segment registration (2ee3735a).
recordedBeforeUpperSegmentRegistration()

const REPO_ROOT = resolve(__dirname, '..')
const ID = S2_RAW_FLOOR_ID
const CAP = S2_RAW_FLOOR_CAP
const LIT = S2_RAW_FLOOR_LITERAL
const TWIN = S2_RAW_FLOOR_TWIN
const HEAVY = 300_000
/** The guard-boundary cells of ed0fccb6 (stage2-scale-upper-window), withheld on main. */
const UW = 'stage2-scale-upper-window'
const MOVED = ['A3', 'A4', 'D2', 'D4', 'E7', 'E8', 'E9', 'E10', 'E12', 'E13', 'E14', 'H8', 'H18', 'I5']

afterEach(() => injectCfGuardBoundaryReadings(null))

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
const layout = s2RawFloorLayout()
const twinK = new Map(layout.twins.map(t => [t.ref, t.k]))
const designated = s2RawFloorCells().filter(c => c.configuration !== undefined)
const cellAt = new Map(designated.map(c => [c.ref, c]))
const decisive = (c: S2RawFloorCell): boolean => s2RawFloorPrediction(c, 'raw-floor') !== s2RawFloorPrediction(c, 'guarded')

/** Rehearse a registration: the reading selected through the seam and the id's evidence held, around `fn`. */
async function withRegistration<T>(reading: S2RawFloorReading, fn: () => Promise<T>): Promise<T> {
  injectCfGuardBoundaryReadings({ [CF_GUARD_BOUNDARY_CAP]: reading })
  try { return await withExactCfEvidence([...VERIFIED_CAPABILITIES, CF_GUARD_BOUNDARY_CAP], fn) } finally { injectCfGuardBoundaryReadings(null) }
}

/** The unpublished state: the registration in place but the id's evidence withheld, so the shape is withheld as before the capture. */
async function withoutEvidence<T>(fn: () => Promise<T>): Promise<T> {
  return withExactCfEvidence([...VERIFIED_CAPABILITIES].filter(id => id !== CF_GUARD_BOUNDARY_CAP), fn)
}

/** A synthetic world: each designated cell's fill by `pick` (default the named reading), twins on the line, positive controls as fixed, controls white. */
function plantedFills(c: CfOracleCase, world: S2RawFloorReading | ((cell: S2RawFloorCell) => string)): Map<string, string> {
  const out = new Map<string, string>()
  for (const e of c.expected.cells) out.set(e.ref, e.matched ? e.fillRgb ?? '' : 'FFFFFFFF')
  for (const [ref, k] of twinK) out.set(ref, s2UwLine(k))
  for (const cell of designated) out.set(cell.ref, typeof world === 'string' ? s2RawFloorPrediction(cell, world) : world(cell))
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
const EVIDENCE = ['2c-num', '2c-minmax', '2c-percent', '2c-percentile', '2c-inverted', '3c-pmp', '3c-percent', '3c-mixed', '3c-resort']

describe('stage2-scale-raw-floor: the case', () => {
  it('is appended last; owns its three ids; the question is registered as raw-floor (its entry removed), the literal id and twin line stay recorded', async () => {
    // Appended last at its time; the literal-lower follow-up stage2-scale-literal-lower is appended after it.
    expect(CF_STAGE2_CASES.map(c => c.id).slice(-4)).toEqual(['stage2-scale-upper-window', ID, 'stage2-scale-literal-lower', 'stage2-scale-upper-segment'])
    for (const id of [CAP, LIT, TWIN]) expect(CF_STAGE2_CAPABILITIES.filter(e => e.capabilityId === id).map(e => e.caseId), id).toEqual([ID])
    expect(CF_STAGE2_NONDISCRIMINATING[CAP]).toBeUndefined()
    for (const id of [LIT, TWIN]) expect(CF_STAGE2_NONDISCRIMINATING[id], id).toMatch(/./)
    expect(CAP).toBe(CF_GUARD_BOUNDARY_CAP)
    expect([...S2_RAW_FLOOR_READINGS]).toEqual([...CF_GUARD_BOUNDARY_READINGS])
    const c = caseOf()
    expect(c.status).toBe('measured')
    expect(c.scaleReadings).toEqual(CF_STAGE2_CASES.find(t => t.id === ID)!.scaleReadings)
    expect(c.scaleReadings!.map(x => [x.id, x.capabilityId, Object.keys(x.hypotheses), x.role ?? (x.record ? 'record' : 'evidence')])).toEqual([
      ...EVIDENCE.map(id => [id, CAP, ['raw-floor', 'guarded'], 'evidence']),
      ['controls', CAP, ['raw-floor', 'guarded'], 'negative-control'], ['w3', CAP, ['raw-floor', 'guarded'], 'record'], ['literal', LIT, ['raw-floor', 'guarded'], 'evidence'],
    ])
    // Registered: the question as raw-floor (published, VERIFIED_CAPABILITIES); the literal id and the twin line never.
    const { CF_STAGE2_PRODUCTION_READINGS } = await import('../scripts/cf-refs/audit')
    expect(CF_STAGE2_PRODUCTION_READINGS[CAP]).toBe('raw-floor')
    expect(VERIFIED_CAPABILITIES.has(CAP)).toBe(true)
    for (const id of [LIT, TWIN]) { expect(Object.prototype.hasOwnProperty.call(CF_STAGE2_PRODUCTION_READINGS, id), id).toBe(false); expect(VERIFIED_CAPABILITIES.has(id), id).toBe(false) }
  })

  it('201 cells: 114 designated (67 evidence, 34 controls, 9 W3, 4 literal), 72 twins, 2 positive controls, 13 controls; A1:Q17, 22 colour rules', () => {
    const c = caseOf()
    expect(c.range).toBe('A1:Q17')
    expect(c.expected.cells.length).toBe(201)
    expect(c.scaleReadings!.map(x => x.refs.length)).toEqual([15, 9, 8, 4, 7, 8, 5, 7, 4, 34, 9, 4])
    expect(designated.length).toBe(114)
    expect(twinK.size).toBe(72)
    expect(c.expected.cells.filter(e => e.matched && !e.scaleApplied).length).toBe(2)
    expect(c.expected.negativeRefs.length).toBe(13)
    expect(c.markers.length).toBe(22)
    for (const e of c.expected.cells) expect(c.markers.filter(m => isCellInSqref(e.ref, m.sqref)).length, e.ref).toBeLessThanOrEqual(1)
    for (const ref of c.scaleReadings!.flatMap(x => x.refs)) expect(c.expected.cells.find(e => e.ref === ref)!.controlRef, ref).toBe(`${colOf(ref)}17`)
    for (const ref of c.expected.negativeRefs) expect(c.markers.some(m => isCellInSqref(ref, m.sqref)), ref).toBe(false)
    // No degenerate single-cell range (Excel collapses X2:X2 on resave).
    for (const m of c.markers) expect(m.sqref, m.sqref).not.toMatch(/^([A-Z]+\d+):\1$/)
  })

  it('the text scope admits the whole range in one batch, far under the 8,388,608-byte cap', async () => {
    const { sheet } = await sheetOf(caseOf())
    const bytes = estimateCfPrepareBytes(sheet.conditionalFormatting!, [{ firstRow: 0, lastRow: 16, firstCol: 0, lastCol: 16 }])
    expect(bytes).toBe(3072876)
    expect(bytes!).toBeLessThan(8388608 * 0.5)
  })

  it('the workbook: the catalogue values, twins of value (k − from)/128, positive controls on their dxf fills; no hidden rows, merges or errors', () => {
    const c = caseOf()
    const values = valuesOf(c)
    for (const cell of s2RawFloorCells()) expect(values.get(cell.ref), cell.ref).toBe(Number(cell.text))
    for (const t of layout.twins) expect(values.get(t.ref), t.ref).toBe((t.k - t.span.from) / 128)
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

  it('every statistical block resolves (production\'s resolver) to its thresholds bit for bit, and they hold in every operation order', () => {
    const values = valuesOf(caseOf())
    for (const b of s2RawFloorBlocks()) {
      const scope = b.values.map((_, i) => values.get(`${b.col}${b.firstRow + i}`)!)
      const sorted = Float64Array.from([...scope].sort((x, y) => x - y))
      const stats = { scopeKey: b.id, numericCount: scope.length, sum: 0, mean: undefined, min: sorted[0], max: sorted[sorted.length - 1], stddevPop: undefined, sortedAsc: sorted, frequencies: new Map(), declaredCells: scope.length, spanCells: scope.length, storedCells: scope.length }
      const points: CfScalePoint[] = b.cfvo.map(v => ({ bound: (v.type === 'min' || v.type === 'max' ? { type: v.type } : { type: v.type, value: Number(v.val) }) as CfScalePoint['bound'], argb: 0 }))
      const got = resolveScaleThresholds(points, stats, { hasEvidence: () => true })
      expect(got.kind, b.id).toBe('ready')
      const t = got.kind === 'ready' ? got.value : []
      for (let i = 0; i < b.thresholds.length; i++) expect(Object.is(t[i], Number(b.thresholds[i])), `${b.id} ${i}`).toBe(true)
    }
    // percent 50 / 100 over a zero minimum: max/2 and max in every operation order (no rounding anywhere).
    for (const max of [2.2, 420]) {
      for (const alt of [max / 2, max * 0.5, (max * 50) / 100, max * (50 / 100), (0 + max) / 2, 0 * 0.5 + max * 0.5, 0 + (max - 0) * 0.5]) expect(Object.is(alt, max / 2), `${max}`).toBe(true)
      for (const alt of [max * 1, (max * 100) / 100, 0 + (max - 0) * 1, 0 * 0 + max * 1]) expect(Object.is(alt, max), `${max}`).toBe(true)
    }
    // Percentile 50 over an odd count is the middle value (rank (n − 1)/2 exactly), percentile 100 the maximum.
    for (const b of s2RawFloorBlocks().filter(x => x.cfvo.some(v => v.type === 'percentile'))) {
      const n = b.values.length
      expect(n % 2, b.id).toBe(1)
      const sorted = b.values.map(x => Number(x.text)).sort((x, y) => x - y)
      expect(sorted[(n - 1) / 2], b.id).toBe(Number(b.thresholds[b.cfvo.findIndex(v => v.type === 'percentile' && v.val === '50')]))
    }
  })

  it('every designated cell is in its designed class; each evidence configuration separates the readings at >= 3 cells (pinned)', () => {
    for (const c of designated) {
      const cls = s2RawFloorClasses(c)
      if (c.role === 'evidence') { expect(c.colours === 2 || c.segment === 0, c.ref).toBe(true); expect(cls, c.ref).toContain('below'); expect(decisive(c), c.ref).toBe(true) }
      else if (c.role === 'control') { expect(cls, c.ref).not.toContain('below'); expect(decisive(c), c.ref).toBe(false) }
      else { expect(c.role).toBe('w3'); expect(c.segment, c.ref).toBe(1) }
    }
    const { configurations } = layout
    const dec = (id: string): number => { const cfg = configurations.find(x => x.id === id)!; return cfg.refs.filter(r => cfg.hypotheses['raw-floor']![r] !== cfg.hypotheses.guarded![r]).length }
    expect(Object.fromEntries([...EVIDENCE, 'literal', 'controls'].map(id => [id, dec(id)]))).toEqual({ '2c-num': 15, '2c-minmax': 9, '2c-percent': 8, '2c-percentile': 4, '2c-inverted': 7, '3c-pmp': 8, '3c-percent': 5, '3c-mixed': 7, '3c-resort': 4, literal: 4, controls: 0 })
    for (const id of EVIDENCE) expect(dec(id), id).toBeGreaterThanOrEqual(3)
    // Controls hold every class but below: above an integer, an integer and not whole.
    const classes = new Set(designated.filter(c => c.role === 'control').flatMap(c => s2RawFloorClasses(c)))
    expect([...classes].sort()).toEqual(['above', 'clear', 'integer'])
    // Both colour orders: every evidence configuration has cells whose R channel ascends and whose R channel descends, or the configuration's colours run both ways across its blocks.
    const rAsc = (c: S2RawFloorCell): boolean => Number.parseInt(c.to.slice(2, 4), 16) > Number.parseInt(c.from.slice(2, 4), 16)
    const evidence = designated.filter(c => c.role === 'evidence')
    expect(new Set(evidence.map(rAsc))).toEqual(new Set([true, false]))
    // Spans: 210, 150, 105 and 30.
    expect([...new Set(evidence.map(c => Math.abs(Number.parseInt(c.to.slice(2, 4), 16) - Number.parseInt(c.from.slice(2, 4), 16))))].sort((a, b) => a - b)).toEqual([30, 105, 150, 210])
  })

  it('the 14 doubles of ed0fccb6 are replicated: each moved cell has an evidence cell with the same f double and span, in its own and other bound kinds', () => {
    // ed0fccb6's 14 cells as (value, lo, hi, from step, to step).
    const uw: Record<string, [number, number, number, number, number]> = {
      A3: [0.13, 0.1, 0.4, 0, 210], A4: [0.21, 0.1, 0.4, 0, 210], D2: [0.11, 0.1, 0.2, 0, 210], D4: [0.15, 0.1, 0.2, 0, 210],
      E7: [0.6, 0.4, 1, 210, 0], E8: [0.44, 0.4, 1, 210, 0], E9: [0.5, 0.4, 1, 210, 0], E10: [0.58, 0.4, 1, 210, 0],
      E12: [1.21, 1.1, 2.2, 210, 0], E13: [1.32, 1.1, 2.2, 210, 0], E14: [1.43, 1.1, 2.2, 210, 0],
      H8: [57, 0, 210, 0, 210], H18: [108, 0, 210, 0, 210], I5: [127, 0, 210, 0, 210],
    }
    const f = (v: number, lo: number, hi: number): number => safeFraction(v, lo, hi)!
    const replicas: Record<string, string[]> = {}
    for (const [ref, [v, lo, hi, a, b]] of Object.entries(uw)) {
      const want = f(v, lo, hi)
      replicas[ref] = designated.filter(c => c.role === 'evidence' && Object.is(f(c.value, c.lo, c.hi), want) && c.from === s2UwLine(a) && c.to === s2UwLine(b)).map(c => `${c.configuration}:${c.ref}`)
    }
    expect(replicas).toEqual({
      A3: ['3c-pmp:G3'], A4: ['3c-pmp:G4'], D2: ['3c-mixed:J2'], D4: ['3c-mixed:J3'],
      E7: ['2c-num:A2', '2c-inverted:D10'], E8: ['2c-num:A3', '2c-inverted:D11'], E9: ['2c-num:A4', '2c-inverted:D12'], E10: ['2c-num:A5', '2c-inverted:D13'],
      E12: ['2c-minmax:B8', '2c-percentile:E9'], E13: ['2c-minmax:B9', '2c-percentile:E10'], E14: ['2c-minmax:B10', '2c-percentile:E11'],
      H8: ['2c-num:A10', '2c-inverted:F2', 'literal:F6', '3c-resort:F11', '3c-pmp:H3'], H18: ['2c-num:A11', '2c-inverted:F3', 'literal:F7', '3c-resort:F12', '3c-pmp:H4'], I5: ['2c-num:A12', '2c-inverted:F4', 'literal:F8', '3c-resort:F13', '3c-pmp:H5'],
    })
    // The same doubles in the other colour order (2-colour percent over 1.1 .. 2.2 runs P to Q; the percent 3-colour lower segment Q to P).
    for (const ref of ['E12', 'E13', 'E14'] as const) { const [v, lo, hi] = uw[ref]!; expect(designated.some(c => c.configuration === '2c-percent' && Object.is(f(c.value, c.lo, c.hi), f(v, lo, hi)) && c.from === s2UwLine(0)), ref).toBe(true) }
    for (const ref of ['H8', 'H18', 'I5'] as const) { const [v, lo, hi] = uw[ref]!; expect(designated.some(c => c.configuration === '3c-percent' && Object.is(f(c.value, c.lo, c.hi), f(v, lo, hi)) && c.from === s2UwLine(210)), ref).toBe(true) }
  })

  it('W3 probes: each upper-segment probe has a lower-segment evidence partner with the same f double bit for bit and the same span; both colour orders; recorded only', () => {
    const f = (c: S2RawFloorCell): number => safeFraction(c.value, c.lo, c.hi)!
    const w3 = designated.filter(c => c.role === 'w3')
    expect(w3.map(c => c.ref)).toEqual(['H10', 'H11', 'H12', 'H13', 'I9', 'I10', 'I11', 'I12', 'I13'])
    for (const u of w3) {
      const partner = designated.find(c => c.role === 'evidence' && c.block === u.block && c.segment === 0 && Object.is(f(c), f(u)))
      expect(partner, u.ref).toBeDefined()
      // The partner's segment runs the opposite colour direction (P / Q / P: lower P to Q, upper Q to P).
      expect([partner!.from, partner!.to], u.ref).toEqual([u.to, u.from])
    }
    // Both palette orders: P / Q / P (H) and Q / P / Q (I).
    expect(new Set(w3.map(c => c.from))).toEqual(new Set([s2UwLine(210), s2UwLine(0)]))
    expect(layout.configurations.find(x => x.id === 'w3')!.record).toMatch(/never evidence/)
  })

  it('the twin line: every predicted fill is on the P line, and every twin\'s product is an IEEE integer on every channel (recomputed)', () => {
    for (const t of layout.twins) {
      const v = (t.k - t.span.from) / 128, f = (v - 0) / (1 - 0)
      for (const i of [2, 4, 6]) {
        const d = Math.abs(Number.parseInt(s2UwLine(t.span.to).slice(i, i + 2), 16) - Number.parseInt(s2UwLine(t.span.from).slice(i, i + 2), 16))
        expect(d, t.ref).toBe(128)
        expect(Number.isInteger(f * d), `${t.ref} ${f * d}`).toBe(true)
        expect(f * d, t.ref).toBe(t.k - t.span.from)
      }
    }
    const fills = new Set(designated.flatMap(c => S2_RAW_FLOOR_READINGS.map(r => s2RawFloorPrediction(c, r))))
    expect(new Set([...fills].map(x => [...twinK].find(([, k]) => s2UwLine(k) === x)?.[0]))).not.toContain(undefined)
    expect(S2_RAW_FLOOR.twinSpans).toEqual([{ from: 0, to: 128 }, { from: 82, to: 210 }])
  })

  it('the independent model is production without the evidence (as before the capture): painted cells are the guarded prediction; every evidence cell (guard boundary, literal num-segment) and W3 probe is withheld', async () => {
    const c = caseOf()
    const fills = await withoutEvidence(() => productionFills(c, c.expected.cells.map(e => e.ref)))
    const withheld: string[] = []
    for (const e of c.expected.cells) {
      const got = fills.get(e.ref)
      if (got === undefined && e.matched) { withheld.push(e.ref); continue }
      const cell = cellAt.get(e.ref)
      const want = !e.matched ? undefined : cell ? s2RawFloorPrediction(cell, 'guarded') : twinK.has(e.ref) ? s2UwLine(twinK.get(e.ref)!) : e.fillRgb
      expect(got, e.ref).toBe(want)
    }
    expect(withheld.sort()).toEqual(designated.filter(x => x.role === 'evidence' || x.role === 'w3').map(x => x.ref).sort())
    expect(withheld.length).toBe(67 + 4 + 9)
  }, HEAVY)

  it('the independent model is production as committed (raw-floor registered and published): Excel\'s fill at every cell production paints; only the W3 cells are withheld (the literal cells paint since the literal-lower registration, a3f85309)', async () => {
    const c = caseOf()
    const fills = await productionFills(c, c.expected.cells.map(e => e.ref))
    const sidecar = JSON.parse(readFileSync(resolve(REPO_ROOT, c.captures[0]!.sidecarPath), 'utf8')) as { observedCells: Array<{ ref: string; fillRgb: string }> }
    const excel = new Map(sidecar.observedCells.map(o => [o.ref, o.fillRgb]))
    const withheld: string[] = []
    for (const e of c.expected.cells) {
      const got = fills.get(e.ref)
      if (got === undefined && e.matched) { withheld.push(e.ref); continue }
      const cell = cellAt.get(e.ref)
      expect(got, e.ref).toBe(!e.matched ? undefined : cell ? s2RawFloorPrediction(cell, 'raw-floor') : twinK.has(e.ref) ? s2UwLine(twinK.get(e.ref)!) : e.fillRgb)
      if (e.matched) expect(got, `${e.ref} Excel`).toBe(excel.get(e.ref))
    }
    expect(withheld.sort()).toEqual(designated.filter(x => x.role === 'w3').map(x => x.ref).sort())
    expect(withheld.length).toBe(9)
  }, HEAVY)

  for (const reading of S2_RAW_FLOOR_READINGS) {
    it(`the independent model is production under the ${reading} registration: every evidence cell paints the reading (the literal cells too, through the literal-lower registration); W3 cells stay withheld; nothing else moves`, async () => {
      const c = caseOf()
      const fills = await withRegistration(reading, () => productionFills(c, c.expected.cells.map(e => e.ref)))
      for (const e of c.expected.cells) {
        const cell = cellAt.get(e.ref)
        const still = cell !== undefined && cell.role === 'w3'
        const want = !e.matched || still ? undefined : cell ? s2RawFloorPrediction(cell, reading) : twinK.has(e.ref) ? s2UwLine(twinK.get(e.ref)!) : e.fillRgb
        expect(fills.get(e.ref), e.ref).toBe(want)
      }
    }, HEAVY)
  }
})

describe('stage2-scale-raw-floor: the evaluator on synthetic observations', () => {
  for (const reading of S2_RAW_FLOOR_READINGS) {
    it(`a planted ${reading} world matches exactly ${reading}, for the question and for the literal cells; the W3 record never decides`, () => {
      const c = caseOf()
      const r = judge(c, plantedFills(c, reading))
      expect([r[CAP]!.status, r[CAP]!.matched, r[LIT]!.status, r[LIT]!.matched]).toEqual(['discriminating', reading, 'discriminating', reading])
      // The W3 probes drawn either way (or off both readings) leave the verdict unchanged.
      const other = S2_RAW_FLOOR_READINGS.find(x => x !== reading)!
      for (const w3 of [(x: S2RawFloorCell) => s2RawFloorPrediction(x, other), () => 'FF000000']) {
        const world = (x: S2RawFloorCell): string => (x.role === 'w3' ? w3(x) : s2RawFloorPrediction(x, reading))
        const w = judge(c, plantedFills(c, world))
        expect([w[CAP]!.status, w[CAP]!.matched], reading).toEqual(['discriminating', reading])
      }
    })
  }

  it('no reading matches: a world at raw floor except one decisive cell of one configuration at guarded is refused, and so is the reverse', () => {
    const c = caseOf()
    for (const [base, flip] of [['raw-floor', 'guarded'], ['guarded', 'raw-floor']] as const) {
      for (const cfg of EVIDENCE) {
        const target = designated.find(x => x.configuration === cfg && decisive(x))!
        const r = judge(c, plantedFills(c, x => s2RawFloorPrediction(x, x === target ? flip : base)))
        expect([r[CAP]!.status, r[CAP]!.code], `${base} with ${target.ref} at ${flip}`).toEqual(['refused', 'formula-hypothesis-none'])
      }
    }
  })

  it('fails closed: a formatted control, a missing or wrong positive control, a twin rendering nothing, a moved negative control', () => {
    const c = caseOf()
    for (const reading of S2_RAW_FLOOR_READINGS) {
      const fills = plantedFills(c, reading)
      const edit = (ref: string, fill: string | undefined): Map<string, string> => { const m = new Map(fills); if (fill === undefined) m.delete(ref); else m.set(ref, fill); return m }
      const pc = layout.positiveControls[0]!.ref
      expect(judge(c, edit('A17', 'FF14F028'))[CAP]!.status, reading).toBe('refused')
      expect(judge(c, edit(pc, undefined))[CAP]!.code, reading).toBe('fill-unobserved')
      expect(judge(c, edit(pc, 'FF000000'))[CAP]!.code, reading).toBe('formula-positive-control-failed')
      expect(judge(c, edit([...twinK.keys()][0]!, 'FFFFFFFF'))[CAP]!.status, reading).toBe('refused')
      // A control just above an integer (A15, 115 in 0 .. 210) drawn one step smaller: nothing matches.
      const moved = judge(c, edit('A15', s2UwLine(114)))
      expect([moved[CAP]!.status, moved[CAP]!.code], reading).toEqual(['refused', 'formula-hypothesis-none'])
    }
  })
})

// ---- a synthetic measured capture through the real verifier, audit, conformance pass and cf-compare (temp root) ----
const sha = (b: Uint8Array): string => createHash('sha256').update(b).digest('hex')

function syntheticRoot(fills: Map<string, string>): { root: string; manifest: CfManifest } {
  const root = mkdtempSync(join(tmpdir(), 'cf-s2-rawfloor-'))
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
  const at = '2026-11-23T00:00:00.000Z'
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
  const published = CF_STAGE2_NONDISCRIMINATING
  const kept = { ...CF_STAGE2_NONDISCRIMINATING, [CAP]: 'pre-registered, awaiting its capture' }

  it('the real manifest: Stage 2 audits ok and complete; the question verified (raw-floor), the literal id and twin line unverified-nondiscriminating; the verified set gains exactly the question', async () => {
    const { auditStage2Manifest } = await import('../scripts/cf-refs/audit')
    const r = auditStage2Manifest(manifest, { rootDir: REPO_ROOT })
    expect([r.ok, r.stageComplete]).toEqual([true, true])
    expect(disposition(r, CAP)).toBe('verified')
    for (const id of [LIT, TWIN]) expect(disposition(r, id), id).toBe('unverified-nondiscriminating')
    expect(r.perCapability.filter(p => p.disposition === 'verified').map(p => p.capabilityId).sort()).toEqual([...before, CAP].sort())
    // With the pre-capture entry put back, the question is unverified again and the verified set is the one before.
    const k = auditStage2Manifest(manifest, { rootDir: REPO_ROOT, nonDiscriminating: kept })
    expect(disposition(k, CAP)).toBe('unverified-nondiscriminating')
    expect(k.perCapability.filter(p => p.disposition === 'verified').map(p => p.capabilityId).sort()).toEqual(before)
  }, HEAVY)

  for (const reading of S2_RAW_FLOOR_READINGS) {
    it(`planted ${reading}: matched itself; registry kept: unverified; entry removed, unregistered: withheld; registered: verified; the other reading registered: withheld`, async () => {
      const { auditStage2Manifest, CF_STAGE2_PRODUCTION_READINGS } = await import('../scripts/cf-refs/audit')
      const unregisteredReadings = Object.fromEntries(Object.entries(CF_STAGE2_PRODUCTION_READINGS).filter(([id]) => id !== CAP))
      const { root, manifest: m } = syntheticRoot(plantedFills(caseOf(), reading))
      const res = verifyCfCase(m.cases.find(x => x.id === ID)!, { rootDir: root })
      expect(res.kind, res.kind === 'refusal' ? `${res.code}: ${res.reason}` : '').toBe('verified')
      if (res.kind !== 'verified') return
      expect(res.matchedHypotheses).toEqual({ [CAP]: reading, [LIT]: reading })
      const k = auditStage2Manifest(m, { rootDir: root, nonDiscriminating: kept })
      expect(k.ok).toBe(true)
      for (const id of [CAP, LIT, TWIN]) expect(disposition(k, id), id).toBe('unverified-nondiscriminating')
      expect(disposition(auditStage2Manifest(m, { rootDir: root, nonDiscriminating: published, productionReadings: unregisteredReadings }), CAP)).toBe('supported-structural')
      const registered = auditStage2Manifest(m, { rootDir: root, nonDiscriminating: published, productionReadings: { ...CF_STAGE2_PRODUCTION_READINGS, [CAP]: reading } })
      expect(disposition(registered, CAP)).toBe('verified')
      expect(registered.perCapability.filter(p => p.disposition === 'verified').map(p => p.capabilityId).sort()).toEqual([...before, CAP].sort())
      const other = S2_RAW_FLOOR_READINGS.find(x => x !== reading)!
      expect(disposition(auditStage2Manifest(m, { rootDir: root, nonDiscriminating: published, productionReadings: { ...CF_STAGE2_PRODUCTION_READINGS, [CAP]: other } }), CAP)).toBe('supported-structural')
    }, HEAVY)
  }

  it('planted "no reading matches" (raw floor with one cell per evidence configuration at guarded): formula-hypothesis-none, never verified, and production keeps the withhold', async () => {
    const { auditStage2Manifest, CF_STAGE2_PRODUCTION_READINGS } = await import('../scripts/cf-refs/audit')
    const flipped = new Set(EVIDENCE.map(cfg => designated.find(x => x.configuration === cfg && decisive(x))!.ref))
    const { root, manifest: m } = syntheticRoot(plantedFills(caseOf(), x => s2RawFloorPrediction(x, flipped.has(x.ref) ? 'guarded' : 'raw-floor')))
    const res = verifyCfCase(m.cases.find(x => x.id === ID)!, { rootDir: root })
    expect(res.kind).toBe('verified')
    if (res.kind !== 'verified') return
    expect((res.refusedCapabilities ?? []).map(x => [x.capabilityId, x.code])).toEqual([[CAP, 'formula-hypothesis-none']])
    for (const reading of S2_RAW_FLOOR_READINGS) expect(disposition(auditStage2Manifest(m, { rootDir: root, nonDiscriminating: published, productionReadings: { ...CF_STAGE2_PRODUCTION_READINGS, [CAP]: reading } }), CAP), reading).toBe('supported-structural')
    // Unpublished production on that world: every evidence cell withheld and justified (67 guard-boundary; the 4 literal cells
    // too, as guard-boundary cells since the literal-lower registration), the W3 probes as upper whole products; nothing disagrees.
    const r = await withoutEvidence(() => stage2Conformance(m, root, [ID]))
    expect(r.violations).toEqual([])
    expect([r.total, r.withheld, r.guardBoundaryWithheld, r.numSegmentWithheld, r.upperWholeWithheld]).toEqual([201, 80, 71, 0, 9])
    // As committed (raw-floor registered and published), the conformance pass catches the 9 flipped cells.
    const committed = await stage2Conformance(m, root, [ID])
    expect(committed.violations.map(v => v.split(':').slice(0, 2).join(':')).sort()).toEqual([...flipped].map(x => `${ID}:${x}`).sort())
  }, HEAVY)
})

// As recorded at ce3f701d: the cases measured then (stage2-scale-literal-lower, a3f85309, is rehearsed by its own case test).
const STAGE2_MEASURED = manifest.cases.filter(x => x.stage === 2 && x.status === 'measured' && x.id !== 'stage2-scale-literal-lower' && x.id !== 'stage2-scale-upper-segment').map(x => x.id)

describe('the rehearsal: each registration against every committed capture, and this case on a synthetic capture', () => {
  // This capture's 67 evidence cells (the guard-boundary shape outside the literal num-segment withhold).
  const EVIDENCE_REFS = designated.filter(x => x.role === 'evidence' && x.configuration !== 'literal').map(x => x.ref)

  it('without the evidence (as before the capture): no violation over the measured Stage 2 cases; the 14 cells of ed0fccb6 and this capture\'s 67 withheld', async () => {
    expect(STAGE2_MEASURED).toContain(ID)
    const r = await withoutEvidence(() => stage2Conformance(manifest, REPO_ROOT, STAGE2_MEASURED))
    expect(r.violations).toEqual([])
    expect([r.total, r.exact, r.withheld, r.numSegmentWithheld, r.upperWholeWithheld, r.guardBoundaryWithheld]).toEqual([1310, 942, 368, 194, 59, 85])
  }, HEAVY)

  it('as committed (raw-floor registered and published): every committed capture agrees; the 14 cells of ed0fccb6 and this capture\'s 67 are painted at Excel\'s fill exactly; no guard-boundary withhold is left', async () => {
    const r = await stage2Conformance(manifest, REPO_ROOT, STAGE2_MEASURED)
    expect(r.violations).toEqual([])
    expect([r.total, r.exact, r.withheld, r.numSegmentWithheld, r.upperWholeWithheld, r.guardBoundaryWithheld]).toEqual([1310, 1027, 283, 194, 59, 0])
    expect(await withRegistration('raw-floor', () => stage2Conformance(manifest, REPO_ROOT, STAGE2_MEASURED))).toEqual(r)
    const uw = manifest.cases.find(x => x.id === UW)!
    const sidecar = JSON.parse(readFileSync(resolve(REPO_ROOT, uw.captures[0]!.sidecarPath), 'utf8')) as { observedCells: Array<{ ref: string; fillRgb: string }> }
    const excel = new Map(sidecar.observedCells.map(o => [o.ref, o.fillRgb]))
    const fills = await productionFills(uw, MOVED)
    for (const ref of MOVED) expect(fills.get(ref), ref).toBe(excel.get(ref))
  }, HEAVY)

  it('guarded registered: the committed captures refute it at exactly the 14 cells of ed0fccb6 and this capture\'s 67 (caught, never published)', async () => {
    const r = await withRegistration('guarded', () => stage2Conformance(manifest, REPO_ROOT, STAGE2_MEASURED))
    // ... and the 4 literal cells (F6..F9), painted through the guard-boundary reading since the literal-lower registration.
    const literal = designated.filter(x => x.configuration === 'literal').map(x => x.ref)
    expect(r.violations.map(v => v.split(':').slice(0, 2).join(':')).sort()).toEqual([...MOVED.map(x => `${UW}:${x}`), ...EVIDENCE_REFS.map(x => `${ID}:${x}`), ...literal.map(x => `${ID}:${x}`)].sort())
  }, HEAVY)

  it('the conformance pass and cf-compare (interactive, paged) on this case: each planted world passes unregistered and under its own registration; the other reading is caught', async () => {
    const { compareCfCase } = await import('../scripts/cf-compare')
    for (const reading of S2_RAW_FLOOR_READINGS) {
      const { root, manifest: m } = syntheticRoot(plantedFills(caseOf(), reading))
      const c = m.cases.find(x => x.id === ID)!
      const unregistered = await withoutEvidence(() => stage2Conformance(m, root, [ID]))
      expect(unregistered.violations, reading).toEqual([])
      expect([unregistered.total, unregistered.withheld], reading).toEqual([201, 80])
      const registered = await withRegistration(reading, () => stage2Conformance(m, root, [ID]))
      expect(registered.violations, reading).toEqual([])
      expect([registered.total, registered.withheld, registered.guardBoundaryWithheld, registered.numSegmentWithheld, registered.upperWholeWithheld], `${reading} registered`).toEqual([201, 9, 0, 0, 9])
      for (const mode of ['interactive', 'paged'] as const) {
        const a = await withoutEvidence(() => compareCfCase(c, mode, { rootDir: root }))
        expect(`${reading} ${mode}: ${a.outcome} ${a.outcome === 'pass' ? '' : a.reason}`).toBe(`${reading} ${mode}: pass `)
        const b = await withRegistration(reading, () => compareCfCase(c, mode, { rootDir: root }))
        expect(`${reading} registered ${mode}: ${b.outcome} ${b.outcome === 'pass' ? '' : b.reason}`).toBe(`${reading} registered ${mode}: pass `)
      }
      // The other reading registered on this world: every non-literal evidence cell disagrees (caught).
      const other = S2_RAW_FLOOR_READINGS.find(x => x !== reading)!
      const wrong = await withRegistration(other, () => stage2Conformance(m, root, [ID]))
      expect(wrong.violations.length, `${other} on a ${reading} world`).toBe(67 + 4)
    }
  }, 900_000)

  it('Stage 5: every measured case agrees with Excel, records unchanged, under each registration (no Stage 5 cell is in the guard-boundary shape)', async () => {
    const { stage5CaseRecords } = await import('../scripts/cf-refs/audit')
    const base = await stage5CaseRecords(manifest, REPO_ROOT)
    expect(base.filter(r => r.compared && !r.agrees)).toEqual([])
    const key = (rows: typeof base) => rows.map(r => `${r.caseId}:${r.ref}:${r.agrees}:${JSON.stringify(r.production)}`).join('\n')
    for (const reading of S2_RAW_FLOOR_READINGS) {
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
