// @vitest-environment node
/**
 * The upper-segment follow-up (brief cf-upper-segment), Phase 1b tasks 3-5: stage2-scale-upper-segment, the pre-registered and
 * registrable case for interior cells of a 3-colour UPPER segment that production withholds today (the literal num-segment
 * shape in segment 1, the integer and window classes of the whole-product shape), and the rehearsal of each reading's
 * registration. Checked here:
 * 1. the case design: ownership, registries, size, values, every designated cell's role, >= 4 decisive cells per evidence
 *    configuration, the replicated and new doubles, the mixed-channel cells, the twins' IEEE-integer products, the text scope;
 * 2. the independent model is production, unregistered (the shapes withheld) and under each registration (only the
 *    upper-segment cells of those shapes move);
 * 3. the evaluator on synthetic observations: each planted reading comes back as the one match; worlds no reading explains
 *    (one flipped decisive cell, one kept near-whole cell drawn one step smaller) are refused; it fails closed on broken
 *    controls;
 * 4. the audit gate end to end on a synthetic measured capture (real verifyCfCase and auditStage2Manifest);
 * 5. each registration against every committed capture: index-position paints every census cell withheld today at Excel's
 *    fill and keeps every other captured cell; truncate is caught exactly at the census cells it misses; Stage 5 records
 *    unchanged; the conformance pass and cf-compare (interactive, paged) pass on each planted world, unregistered and under
 *    its own registration.
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
  S2_UPPER_NEAR_TOP_CAP,
  S2_UPPER_SEGMENT,
  S2_UPPER_SEGMENT_CAP,
  S2_UPPER_SEGMENT_ID,
  S2_UPPER_SEGMENT_READINGS,
  S2_UPPER_SEGMENT_TWIN,
  s2UpperSegmentCells,
  s2UpperSegmentClasses,
  s2UpperSegmentFraction,
  s2UpperSegmentLayout,
  s2UpperSegmentNearTop,
  s2UpperSegmentPrediction,
  s2UwLine,
  type CfManifest,
  type CfOracleCase,
  type S2UpperSegmentCell,
  type S2UpperSegmentReading,
} from '../scripts/cf-refs/cases'
import { CF_MANIFEST_PATH } from '../scripts/cf-refs/generate'
import { computeWorksheetRangeGeometry, evaluateFormulaCapabilities, formulaControlsOf, isCellInSqref, readZipPartsSync, stage2ReadingsObservationOf, verifyCfCase } from '../scripts/cf-refs/verify'
import { upperSegmentCensusOf } from '../scripts/cf-refs/upper-segment-candidates'
import { stage2Conformance } from './helpers/cf-stage2-conformance'
import { unregisterUpperSegment } from './helpers/cf-upper-segment-recorded'
import { OfficePackage } from '../src/core/zip'
import { parseXlsx } from '../src/xlsx/parse'
import { cfDocumentEvaluationOf, prepareCfPaint } from '../src/xlsx/cf/render'
import { cfRuntimeEvidence } from '../src/xlsx/cf/predicates'
import { CF_UPPER_NEAR_TOP_CAP, CF_UPPER_SEGMENT_CAP, CF_UPPER_SEGMENT_READINGS, injectCfUpperNearTopReadings, injectCfUpperSegmentReadings, safeFraction } from '../src/xlsx/cf/colorScale'
import { estimateCfPrepareBytes } from '../src/xlsx/cf/service'
import { VERIFIED_CAPABILITIES, VERIFIED_STAGE2_CAPABILITIES, withExactCfEvidence } from '../src/xlsx/cf/evidence'

const REPO_ROOT = resolve(__dirname, '..')
const ID = S2_UPPER_SEGMENT_ID
const CAP = S2_UPPER_SEGMENT_CAP
const TWIN = S2_UPPER_SEGMENT_TWIN
const NEAR = S2_UPPER_NEAR_TOP_CAP
const HEAVY = 300_000

afterEach(() => { injectCfUpperSegmentReadings(null); injectCfUpperNearTopReadings(null) })

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
const layout = s2UpperSegmentLayout()
const twinFill = new Map(layout.twins.map(t => [t.ref, t.fill]))
const designated = s2UpperSegmentCells()
const cellAt = new Map(designated.map(c => [c.ref, c]))
const decisive = (c: S2UpperSegmentCell): boolean => s2UpperSegmentPrediction(c, 'index-position') !== s2UpperSegmentPrediction(c, 'truncate')
const EVIDENCE = ['pl-qp-210', 'pl-pq-420', 'pl-pq-dec', 'pl-qp-resort', 'pl-qp-stat', 'ds-dd-desc', 'ds-x-mixed', 'ds-z-mixed', 'ds-z-stat']
const interior = (c: S2UpperSegmentCell): boolean => c.segment === 1 && c.value > c.lo && c.value < c.hi
const literal = (c: S2UpperSegmentCell): boolean => c.block.bounds === 'literal'
/** The upper-segment cells production withholds today: the literal num-segment shape (f·2^20 not an integer), any near-whole product, the near-top sub-shape. */
const withheldToday = (c: S2UpperSegmentCell): boolean =>
  interior(c) && ((literal(c) && !Number.isInteger(s2UpperSegmentFraction(c) * 1048576)) || s2UpperSegmentClasses(c).some(x => x !== 'clear') || s2UpperSegmentNearTop(c))
const nearTopRefs = designated.filter(c => c.configuration === 'near-top').map(c => c.ref)
const shapeRefs = designated.filter(withheldToday).map(c => c.ref)

/**
 * Rehearse a registration: the reading selected through the seam for the main id (and, `nearTop`, for the near-top id: by
 * default the same reading, or another), with the evidence of each registered id held, around `fn`.
 */
async function withRegistration<T>(reading: S2UpperSegmentReading, fn: () => Promise<T>, nearTop: S2UpperSegmentReading | false = reading): Promise<T> {
  injectCfUpperSegmentReadings({ [CF_UPPER_SEGMENT_CAP]: reading })
  injectCfUpperNearTopReadings({ [CF_UPPER_NEAR_TOP_CAP]: nearTop === false ? ('unregistered' as S2UpperSegmentReading) : nearTop })
  const caps = [...VERIFIED_CAPABILITIES, CF_UPPER_SEGMENT_CAP, CF_UPPER_NEAR_TOP_CAP].filter(id => nearTop !== false || id !== CF_UPPER_NEAR_TOP_CAP)
  try { return await withExactCfEvidence(caps, fn) } finally { injectCfUpperSegmentReadings(null); injectCfUpperNearTopReadings(null) }
}
/** The unregistered state (as before the capture): both ids shadowed by an invalid proposal, around `fn`. */
async function unregistered<T>(fn: () => Promise<T>): Promise<T> {
  unregisterUpperSegment()
  try { return await fn() } finally { injectCfUpperSegmentReadings(null); injectCfUpperNearTopReadings(null) }
}

/** A synthetic world: each designated cell's fill by `world` (a reading or a function), twins as fixed, positive controls as fixed, controls white. */
function plantedFills(c: CfOracleCase, world: S2UpperSegmentReading | ((cell: S2UpperSegmentCell) => string)): Map<string, string> {
  const out = new Map<string, string>()
  for (const e of c.expected.cells) out.set(e.ref, e.matched ? e.fillRgb ?? '' : 'FFFFFFFF')
  for (const [ref, fill] of twinFill) out.set(ref, fill)
  for (const cell of designated) out.set(cell.ref, typeof world === 'string' ? s2UpperSegmentPrediction(cell, world) : world(cell))
  return out
}
const plantedSidecar = (c: CfOracleCase, fills: Map<string, string>): { observedCells: Array<Record<string, unknown>> } => ({
  observedCells: c.expected.cells.filter(e => fills.get(e.ref) !== undefined).map(e => ({ ref: e.ref, controlRef: e.controlRef, fillSource: 'excel-display-format', fillRgb: fills.get(e.ref) })),
})
const judge = (c: CfOracleCase, fills: Map<string, string>) => evaluateFormulaCapabilities(c.scaleReadings!, formulaControlsOf(c), stage2ReadingsObservationOf(plantedSidecar(c, fills)))
/** One channel of an ARGB hex moved by `by` steps toward `to` (the per-channel step direction). */
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

describe('stage2-scale-upper-segment: the case', () => {
  it('is appended last; owns its three ids; both questions registered as index-position (their entries removed), the twins stay recorded', async () => {
    // Appended last at its time; the formula-bounds follow-up is appended after it.
    expect(CF_STAGE2_CASES.map(c => c.id).slice(-3)).toEqual(['stage2-scale-literal-lower', ID, 'stage2-scale-formula-bounds'])
    for (const id of [CAP, NEAR, TWIN]) {
      expect(CF_STAGE2_CAPABILITIES.filter(e => e.capabilityId === id).map(e => e.caseId), id).toEqual([ID])
    }
    expect([CF_STAGE2_NONDISCRIMINATING[CAP], CF_STAGE2_NONDISCRIMINATING[NEAR]]).toEqual([undefined, undefined])
    expect(CF_STAGE2_NONDISCRIMINATING[TWIN]).toMatch(/./)
    expect(CAP).toBe(CF_UPPER_SEGMENT_CAP)
    expect(NEAR).toBe(CF_UPPER_NEAR_TOP_CAP)
    expect([...S2_UPPER_SEGMENT_READINGS]).toEqual([...CF_UPPER_SEGMENT_READINGS])
    const c = caseOf()
    expect(c.status).toBe('measured')
    expect(c.scaleReadings).toEqual(CF_STAGE2_CASES.find(t => t.id === ID)!.scaleReadings)
    expect(c.scaleReadings!.map(x => [x.id, x.capabilityId, Object.keys(x.hypotheses), x.role ?? (x.record ? 'record' : 'evidence')])).toEqual([
      ...EVIDENCE.map(id => [id, CAP, ['index-position', 'truncate'], 'evidence']),
      ['near-top', NEAR, ['index-position', 'truncate'], 'evidence'],
      ['controls', CAP, ['index-position', 'truncate'], 'negative-control'],
    ])
    const { CF_STAGE2_PRODUCTION_READINGS } = await import('../scripts/cf-refs/audit')
    // Registered: both questions as index-position (matched by 2ee3735a, published); the twins never.
    for (const id of [CAP, NEAR]) {
      expect(CF_STAGE2_PRODUCTION_READINGS[id], id).toBe('index-position')
      expect(VERIFIED_CAPABILITIES.has(id), id).toBe(true)
    }
    expect(Object.prototype.hasOwnProperty.call(CF_STAGE2_PRODUCTION_READINGS, TWIN)).toBe(false)
    expect(VERIFIED_CAPABILITIES.has(TWIN)).toBe(false)
  })

  it('207 cells: 86 designated (48 decisive, 26 kept, 5 clear, 7 controls; 9 of them near-top), 105 twins, 2 positive controls, 14 unformatted controls; A1:X14, 32 colour rules', () => {
    const c = caseOf()
    expect(c.range).toBe('A1:X14')
    expect(c.expected.cells.length).toBe(207)
    expect(c.scaleReadings!.map(x => x.refs.length)).toEqual([10, 11, 11, 9, 10, 5, 5, 5, 4, 9, 7])
    const count = (role: string): number => designated.filter(x => x.role === role).length
    expect([designated.length, count('decisive'), count('kept'), count('clear'), count('control')]).toEqual([86, 48, 26, 5, 7])
    expect(layout.twins.length).toBe(105)
    expect(c.expected.cells.filter(e => e.matched && !e.scaleApplied).length).toBe(2)
    expect(c.expected.negativeRefs.length).toBe(14)
    expect(c.markers.length).toBe(32)
    for (const e of c.expected.cells) expect(c.markers.filter(m => isCellInSqref(e.ref, m.sqref)).length, e.ref).toBeLessThanOrEqual(1)
    for (const ref of c.scaleReadings!.flatMap(x => x.refs)) expect(c.expected.cells.find(e => e.ref === ref)!.controlRef, ref).toBe(`${colOf(ref)}14`)
    for (const ref of c.expected.negativeRefs) expect(c.markers.some(m => isCellInSqref(ref, m.sqref)), ref).toBe(false)
    for (const m of c.markers) expect(m.sqref, m.sqref).not.toMatch(/^([A-Z]+\d+):\1$/)
    // Every block rule is 3-colour (literal num, or min / percent 50 / max); every twin rule is num 0 / 1 (the line) or num 0 / 1 / 2.
    for (const m of c.markers) {
      if (m.capabilityId === CAP || m.capabilityId === NEAR) expect([['num', 'num', 'num'], ['min', 'percent', 'max']], m.sqref).toContainEqual(m.scale!.cfvo.map(v => v.type))
      else {
        expect(m.capabilityId, m.sqref).toBe(TWIN)
        expect([[{ type: 'num', val: '0' }, { type: 'num', val: '1' }], [{ type: 'num', val: '0' }, { type: 'num', val: '1' }, { type: 'num', val: '2' }]], m.sqref).toContainEqual(m.scale!.cfvo)
      }
    }
  })

  it('the text scope admits the whole range in one batch, well under the 8,388,608-byte cap', async () => {
    const { sheet } = await sheetOf(caseOf())
    const bytes = estimateCfPrepareBytes(sheet.conditionalFormatting!, [{ firstRow: 0, lastRow: 13, firstCol: 0, lastCol: 23 }])
    // 58.7% of the cap since the near-top columns J..L (three rules, their twins); one batch admits it.
    expect(bytes).toBe(4920160)
    expect(bytes!).toBeLessThan(8388608 * 0.6)
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

  it('every designated cell is in its designed role; each evidence configuration separates the readings at >= 4 cells; classes, bounds and colour orders (pinned)', () => {
    for (const c of designated) {
      const f = s2UpperSegmentFraction(c)
      // The model's fraction is production's double.
      expect(Object.is(f, safeFraction(c.value, c.lo, c.hi)), c.ref).toBe(true)
      if (c.role === 'decisive') expect([interior(c), decisive(c)], c.ref).toEqual([true, true])
      else if (c.role === 'kept') expect([interior(c), decisive(c), s2UpperSegmentClasses(c).some(x => x !== 'clear') || s2UpperSegmentNearTop(c)], c.ref).toEqual([true, false, true])
      else if (c.role === 'clear') expect([interior(c), decisive(c), s2UpperSegmentClasses(c).every(x => x === 'clear')], c.ref).toEqual([true, false, true])
      else expect([interior(c), decisive(c)], c.ref).toEqual([false, false])
    }
    const dec = (id: string): number => { const cfg = layout.configurations.find(x => x.id === id)!; return cfg.refs.filter(r => cfg.hypotheses['index-position']![r] !== cfg.hypotheses.truncate![r]).length }
    // The near-top configuration: every cell within 1e-11 below the full span on some channel; no other cell is.
    for (const c of designated.filter(interior)) expect(s2UpperSegmentNearTop(c), c.ref).toBe(c.configuration === 'near-top')
    expect(Object.fromEntries([...EVIDENCE, 'near-top', 'controls'].map(id => [id, dec(id)]))).toEqual({
      'pl-qp-210': 5, 'pl-pq-420': 5, 'pl-pq-dec': 5, 'pl-qp-resort': 4, 'pl-qp-stat': 5, 'ds-dd-desc': 4, 'ds-x-mixed': 4, 'ds-z-mixed': 4, 'ds-z-stat': 4, 'near-top': 8, controls: 0,
    })
    const spans = [...new Set(designated.filter(c => c.role !== 'control').map(c => `${c.block.id} ${c.lo}..${c.hi}`))]
    expect(spans).toEqual(['pl-qp-210 217..427', 'pl-pq-420 210..630', 'pl-pq-dec 2.1..3.5', 'pl-qp-resort 210..420', 'pl-qp-stat 210..420', 'ds-dd-desc 105..315', 'ds-x-mixed 78..156', 'ds-z-mixed 66..132', 'ds-z-stat 66..132', 'nt-qp-210 217..427', 'nt-pq-420 210..630', 'nt-dd-1 0..1'])
    const dirs = (c: S2UpperSegmentCell): string => [2, 4, 6].map(i => (Number.parseInt(c.to.slice(i, i + 2), 16) > Number.parseInt(c.from.slice(i, i + 2), 16) ? '+' : '-')).join('')
    expect(Object.fromEntries(EVIDENCE.map(id => [id, dirs(designated.find(c => c.configuration === id)!)]))).toEqual({
      'pl-qp-210': '-+-', 'pl-pq-420': '+-+', 'pl-pq-dec': '+-+', 'pl-qp-resort': '-+-', 'pl-qp-stat': '-+-', 'ds-dd-desc': '---', 'ds-x-mixed': '+-+', 'ds-z-mixed': '+-+', 'ds-z-stat': '+-+',
    })
  })

  it('the mixed cells: index-position moves some channels of a cell and keeps others (ds-dd, ds-x, ds-z), which no per-double rule does', () => {
    const mixed = designated.filter(c => c.role === 'decisive').filter(c => {
      const a = s2UpperSegmentPrediction(c, 'index-position'), b = s2UpperSegmentPrediction(c, 'truncate')
      const moved = [2, 4, 6].map(i => a.slice(i, i + 2) !== b.slice(i, i + 2))
      const cls = s2UpperSegmentClasses(c)
      return moved.some(Boolean) && cls.some((x, i) => x !== 'clear' && !moved[i])
    }).map(c => c.ref)
    expect(mixed).toEqual(['G2', 'G3', 'G4', 'G5', 'H4', 'I3'])
  })

  it('the doubles: most are new (not among the 75 doubles of the census), some replicate captured ones', async () => {
    const { rows } = await upperSegmentCensusOf()
    const captured = new Set(rows.map(r => r.fHex))
    const bits = (x: number): string => { const b = new Float64Array([x]); return new BigUint64Array(b.buffer)[0]!.toString(16).padStart(16, '0') }
    const near = designated.filter(c => c.role === 'decisive' || c.role === 'kept')
    const fresh = near.filter(c => !captured.has(bits(s2UpperSegmentFraction(c))))
    expect([near.length, fresh.length]).toEqual([74, 64])
  }, HEAVY)

  it('the twins: every live prediction is a twin (or a literal positive control in the controls), and every twin\'s product is an IEEE integer on every channel (recomputed)', () => {
    for (const t of layout.twins) {
      const [from, to] = t.rule.colors
      const v = Number(t.text), f = t.rule.cfvo.length === 2 ? v : 1
      const [a, b] = t.rule.cfvo.length === 2 ? [from, to] : [t.rule.colors[Math.max(0, v - 1)]!, t.rule.colors[Math.max(1, v)]!]
      for (const i of [0, 2, 4, 6]) {
        const d = Math.abs(Number.parseInt(b.slice(i, i + 2), 16) - Number.parseInt(a.slice(i, i + 2), 16))
        expect(Number.isInteger(f * d), `${t.ref} ${f * d}`).toBe(true)
      }
      const line = S2_UPPER_SEGMENT.twinSpans.some(sp => from === s2UwLine(sp.from) && to === s2UwLine(sp.to))
      if (!line) expect(t.fill, t.ref).toBe(t.rule.colors[v])
    }
    for (const cfg of layout.configurations) {
      for (const h of Object.values(cfg.hypotheses)) for (const [ref, token] of Object.entries(h)) {
        if (token.startsWith('twin:')) expect(twinFill.get(token.slice(5)), `${cfg.id} ${ref}`).toBe(s2UpperSegmentPrediction(cellAt.get(ref)!, Object.keys(cfg.hypotheses).find(k => cfg.hypotheses[k] === h) as S2UpperSegmentReading))
        else { expect(cfg.id, ref).toBe('controls'); expect(Object.values(cfg.positiveControls ?? {}), ref).toContain(token) }
      }
    }
    expect(S2_UPPER_SEGMENT.twinSpans).toEqual([{ from: 0, to: 128 }, { from: 82, to: 210 }])
  })

  it('the independent model is production unregistered (as before the capture): painted cells are the truncation; every upper-segment cell of the withheld shapes is withheld', async () => {
    const c = caseOf()
    const fills = await unregistered(() => productionFills(c, c.expected.cells.map(e => e.ref)))
    const withheld: string[] = []
    for (const e of c.expected.cells) {
      const got = fills.get(e.ref)
      if (got === undefined && e.matched) { withheld.push(e.ref); continue }
      const cell = cellAt.get(e.ref)
      expect(got, e.ref).toBe(!e.matched ? undefined : cell ? s2UpperSegmentPrediction(cell, 'truncate') : twinFill.get(e.ref) ?? e.fillRgb)
    }
    expect(withheld.sort()).toEqual([...shapeRefs].sort())
    expect(shapeRefs.length).toBe(78)
    // Registered, but without the two ids' evidence: the same.
    const ev = await withExactCfEvidence([...VERIFIED_CAPABILITIES].filter(id => id !== CAP && id !== NEAR), () => productionFills(c, c.expected.cells.map(e => e.ref)))
    expect(ev).toEqual(fills)
    // As committed (index-position registered and published): Excel's fill at every cell; nothing withheld.
    const committed = await productionFills(c, c.expected.cells.map(e => e.ref))
    const sidecar = JSON.parse(readFileSync(resolve(REPO_ROOT, c.captures[0]!.sidecarPath), 'utf8')) as { observedCells: Array<{ ref: string; fillRgb: string }> }
    const excel = new Map(sidecar.observedCells.map(o => [o.ref, o.fillRgb.toUpperCase()]))
    for (const e of c.expected.cells) expect(committed.get(e.ref), e.ref).toBe(e.matched ? excel.get(e.ref) : undefined)
  }, HEAVY)

  for (const reading of S2_UPPER_SEGMENT_READINGS) {
    it(`the independent model is production under the ${reading} registration: with the near-top id at the same reading every upper-segment cell paints it; with the main id alone, or the near-top id at the other reading, the near-top cells stay withheld; nothing else moves`, async () => {
      const c = caseOf()
      const other = S2_UPPER_SEGMENT_READINGS.find(x => x !== reading)!
      for (const nearTop of [reading, false, other] as const) {
        const fills = await withRegistration(reading, () => productionFills(c, c.expected.cells.map(e => e.ref)), nearTop)
        for (const e of c.expected.cells) {
          const cell = cellAt.get(e.ref)
          const want = !e.matched || (nearTop !== reading && nearTopRefs.includes(e.ref)) ? undefined : cell ? s2UpperSegmentPrediction(cell, reading) : twinFill.get(e.ref) ?? e.fillRgb
          expect(fills.get(e.ref), `${e.ref} near-top ${nearTop}`).toBe(want)
        }
      }
    }, HEAVY)
  }
})

describe('stage2-scale-upper-segment: the evaluator on synthetic observations', () => {
  for (const reading of S2_UPPER_SEGMENT_READINGS) {
    it(`a planted ${reading} world matches exactly ${reading}`, () => {
      const c = caseOf()
      const r = judge(c, plantedFills(c, reading))
      expect([r[CAP]!.status, r[CAP]!.matched]).toEqual(['discriminating', reading])
    })
  }

  it('no reading matches: one decisive cell of any configuration at the other reading, or one kept near-whole cell drawn one step smaller on its near-whole channels, is refused', () => {
    const c = caseOf()
    for (const [base, flip] of [['index-position', 'truncate'], ['truncate', 'index-position']] as const) {
      for (const cfg of EVIDENCE) {
        const target = designated.find(x => x.configuration === cfg && decisive(x))!
        const r = judge(c, plantedFills(c, x => s2UpperSegmentPrediction(x, x === target ? flip : base)))
        expect([r[CAP]!.status, r[CAP]!.code], `${base} with ${target.ref} at ${flip}`).toEqual(['refused', 'formula-hypothesis-none'])
      }
    }
    // A kept cell (both readings agree) drawn one step smaller: a "minus one at every whole product" world. Nothing matches.
    for (const cfg of EVIDENCE) {
      const target = designated.find(x => x.configuration === cfg && x.role === 'kept')
      if (!target) continue
      const r = judge(c, plantedFills(c, x => (x === target ? stepped(s2UpperSegmentPrediction(x, 'index-position'), x.from, x.to, -1) : s2UpperSegmentPrediction(x, 'index-position'))))
      expect([r[CAP]!.status, r[CAP]!.code], `${target.ref} one step smaller`).toEqual(['refused', 'formula-hypothesis-none'])
    }
  })

  it('the near-top question is judged on its own: index-position everywhere but the near-top probes at truncate matches index-position for the main id and truncate for the near-top id (production then keeps the near-top cells withheld); one near-top probe off both readings refuses only the near-top id', () => {
    const c = caseOf()
    const split = judge(c, plantedFills(c, x => s2UpperSegmentPrediction(x, x.configuration === 'near-top' ? 'truncate' : 'index-position')))
    expect([split[CAP]!.status, split[CAP]!.matched, split[NEAR]!.status, split[NEAR]!.matched]).toEqual(['discriminating', 'index-position', 'discriminating', 'truncate'])
    for (const col of ['J', 'K', 'L']) {
      const target = designated.find(x => x.configuration === 'near-top' && x.ref.startsWith(col) && decisive(x))!
      const r = judge(c, plantedFills(c, x => (x === target ? stepped(s2UpperSegmentPrediction(x, 'index-position'), x.from, x.to, -1) : s2UpperSegmentPrediction(x, 'index-position'))))
      expect([r[CAP]!.status, r[CAP]!.matched, r[NEAR]!.status, r[NEAR]!.code], target.ref).toEqual(['discriminating', 'index-position', 'refused', 'formula-hypothesis-none'])
    }
  })

  it('fails closed: a formatted control, a missing or wrong positive control, a twin rendering nothing, a moved negative control', () => {
    const c = caseOf()
    for (const reading of S2_UPPER_SEGMENT_READINGS) {
      const fills = plantedFills(c, reading)
      const edit = (ref: string, fill: string | undefined): Map<string, string> => { const m = new Map(fills); if (fill === undefined) m.delete(ref); else m.set(ref, fill); return m }
      const pc = layout.positiveControls[0]!.ref
      expect(judge(c, edit('A14', 'FF14F028'))[CAP]!.status, reading).toBe('refused')
      expect(judge(c, edit(pc, undefined))[CAP]!.code, reading).toBe('fill-unobserved')
      expect(judge(c, edit(pc, 'FF000000'))[CAP]!.code, reading).toBe('formula-positive-control-failed')
      expect(judge(c, edit([...twinFill.keys()][0]!, 'FFFFFFFF'))[CAP]!.status, reading).toBe('refused')
      // An endpoint control (A13, the maximum 427 of pl-qp-210) drawn one step off P: nothing matches.
      const moved = judge(c, edit('A13', 'FF15EF29'))
      expect([moved[CAP]!.status, moved[CAP]!.code], reading).toEqual(['refused', 'formula-hypothesis-none'])
    }
  })
})

// ---- a synthetic measured capture through the real verifier, audit, conformance pass and cf-compare (temp root) ----
const sha = (b: Uint8Array): string => createHash('sha256').update(b).digest('hex')

function syntheticRoot(fills: Map<string, string>): { root: string; manifest: CfManifest } {
  const root = mkdtempSync(join(tmpdir(), 'cf-s2-upper-segment-'))
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
  // The verified set without the two questions, and the registry with their pre-capture entries put back (the rehearsal's start).
  const before = [...VERIFIED_STAGE2_CAPABILITIES].filter(id => id !== CAP && id !== NEAR).sort()
  const removed = CF_STAGE2_NONDISCRIMINATING
  const kept = { ...CF_STAGE2_NONDISCRIMINATING, [CAP]: 'pre-registered, awaiting its capture', [NEAR]: 'pre-registered, awaiting its capture' }

  it('the real manifest: Stage 2 audits ok and complete; both questions verified (index-position), the twins unverified-nondiscriminating; the verified set gains exactly the two', async () => {
    const { auditStage2Manifest } = await import('../scripts/cf-refs/audit')
    const r = auditStage2Manifest(manifest, { rootDir: REPO_ROOT })
    expect([r.ok, r.stageComplete]).toEqual([true, true])
    for (const id of [CAP, NEAR]) expect(disposition(r, id), id).toBe('verified')
    expect(disposition(r, TWIN)).toBe('unverified-nondiscriminating')
    expect(r.perCapability.filter(p => p.disposition === 'verified').map(p => p.capabilityId).sort()).toEqual([...before, CAP, NEAR].sort())
    const k = auditStage2Manifest(manifest, { rootDir: REPO_ROOT, nonDiscriminating: kept })
    expect(k.perCapability.filter(p => p.disposition === 'verified').map(p => p.capabilityId).sort()).toEqual(before)
  }, HEAVY)

  for (const reading of S2_UPPER_SEGMENT_READINGS) {
    it(`planted ${reading}: matched itself; registry kept: unverified; entry removed, unregistered: withheld; registered: verified; the other reading registered: withheld`, async () => {
      const { auditStage2Manifest, CF_STAGE2_PRODUCTION_READINGS } = await import('../scripts/cf-refs/audit')
      const { root, manifest: m } = syntheticRoot(plantedFills(caseOf(), reading))
      const res = verifyCfCase(m.cases.find(x => x.id === ID)!, { rootDir: root })
      expect(res.kind, res.kind === 'refusal' ? `${res.code}: ${res.reason}` : '').toBe('verified')
      if (res.kind !== 'verified') return
      expect(res.matchedHypotheses).toEqual({ [CAP]: reading, [NEAR]: reading })
      const k = auditStage2Manifest(m, { rootDir: root, nonDiscriminating: kept })
      expect(k.ok).toBe(true)
      for (const id of [CAP, NEAR, TWIN]) expect(disposition(k, id), id).toBe('unverified-nondiscriminating')
      const unregisteredReadings = Object.fromEntries(Object.entries(CF_STAGE2_PRODUCTION_READINGS).filter(([id]) => id !== CAP && id !== NEAR))
      expect(disposition(auditStage2Manifest(m, { rootDir: root, nonDiscriminating: removed, productionReadings: unregisteredReadings }), CAP)).toBe('supported-structural')
      const registered = auditStage2Manifest(m, { rootDir: root, nonDiscriminating: removed, productionReadings: { ...CF_STAGE2_PRODUCTION_READINGS, [CAP]: reading, [NEAR]: reading } })
      for (const id of [CAP, NEAR]) expect(disposition(registered, id), id).toBe('verified')
      expect(registered.perCapability.filter(p => p.disposition === 'verified').map(p => p.capabilityId).sort()).toEqual([...before, CAP, NEAR].sort())
      // The main id alone (the pre-committed narrowing): verified, the near-top id withheld.
      const main = auditStage2Manifest(m, { rootDir: root, nonDiscriminating: removed, productionReadings: { ...unregisteredReadings, [CAP]: reading } })
      expect([disposition(main, CAP), disposition(main, NEAR)]).toEqual(['verified', 'supported-structural'])
      const other = S2_UPPER_SEGMENT_READINGS.find(x => x !== reading)!
      expect(disposition(auditStage2Manifest(m, { rootDir: root, nonDiscriminating: removed, productionReadings: { ...CF_STAGE2_PRODUCTION_READINGS, [CAP]: other, [NEAR]: other } }), CAP)).toBe('supported-structural')
    }, HEAVY)
  }

  it('planted "no reading matches" (index-position with one decisive cell per evidence configuration at truncate): formula-hypothesis-none, never verified, and production keeps the withhold', async () => {
    const { auditStage2Manifest, CF_STAGE2_PRODUCTION_READINGS } = await import('../scripts/cf-refs/audit')
    const flipped = new Set(EVIDENCE.map(cfg => designated.find(x => x.configuration === cfg && decisive(x))!.ref))
    const { root, manifest: m } = syntheticRoot(plantedFills(caseOf(), x => s2UpperSegmentPrediction(x, flipped.has(x.ref) ? 'truncate' : 'index-position')))
    const res = verifyCfCase(m.cases.find(x => x.id === ID)!, { rootDir: root })
    expect(res.kind).toBe('verified')
    if (res.kind !== 'verified') return
    expect((res.refusedCapabilities ?? []).map(x => [x.capabilityId, x.code])).toEqual([[CAP, 'formula-hypothesis-none']])
    expect(res.matchedHypotheses).toEqual({ [NEAR]: 'index-position' })
    for (const reading of S2_UPPER_SEGMENT_READINGS) expect(disposition(auditStage2Manifest(m, { rootDir: root, nonDiscriminating: removed, productionReadings: { ...CF_STAGE2_PRODUCTION_READINGS, [CAP]: reading } }), CAP), reading).toBe('supported-structural')
    // Unregistered production on that world: every upper-segment cell of the shapes withheld and justified; nothing disagrees.
    const r = await unregistered(() => stage2Conformance(m, root, [ID]))
    expect(r.violations).toEqual([])
    expect([r.total, r.withheld, r.numSegmentWithheld, r.upperWholeWithheld, r.guardBoundaryWithheld]).toEqual([207, 78, 64, 14, 0])
    // As committed (index-position registered and published), the conformance pass catches the 9 flipped cells.
    const committed = await stage2Conformance(m, root, [ID])
    expect(committed.violations.map(v => v.split(':').slice(0, 2).join(':')).sort()).toEqual([...flipped].map(x => `${ID}:${x}`).sort())
  }, HEAVY)
})

const STAGE2_MEASURED = manifest.cases.filter(x => x.stage === 2 && x.status === 'measured' && x.id !== 'stage2-scale-formula-bounds').map(x => x.id)

describe('the rehearsal: each registration against every committed capture, and this case on a synthetic capture', () => {
  it('unregistered (as before the registration): no violation over the measured Stage 2 cases, this capture included; as committed: no violation, 347 more cells exact', async () => {
    expect(STAGE2_MEASURED).toContain(ID)
    const r = await unregistered(() => stage2Conformance(manifest, REPO_ROOT, STAGE2_MEASURED))
    expect(r.violations).toEqual([])
    expect([r.total, r.exact, r.withheld, r.numSegmentWithheld, r.upperWholeWithheld, r.guardBoundaryWithheld]).toEqual([1697, 1320, 377, 274, 73, 0])
    const now = await stage2Conformance(manifest, REPO_ROOT, STAGE2_MEASURED)
    expect(now.violations).toEqual([])
    expect([now.total, now.exact, now.withheld, now.numSegmentWithheld, now.upperWholeWithheld, now.guardBoundaryWithheld]).toEqual([1697, 1667, 30, 0, 0, 0])
  }, HEAVY)

  it('index-position registered: every committed capture agrees; every census cell withheld today is painted at Excel\'s fill; nothing else moves', async () => {
    const base = await unregistered(() => stage2Conformance(manifest, REPO_ROOT, STAGE2_MEASURED))
    const r = await withRegistration('index-position', () => stage2Conformance(manifest, REPO_ROOT, STAGE2_MEASURED))
    expect(r.violations).toEqual([])
    expect([r.total, r.exact - base.exact, base.withheld - r.withheld, base.numSegmentWithheld - r.numSegmentWithheld, base.upperWholeWithheld - r.upperWholeWithheld, r.guardBoundaryWithheld]).toEqual([1697, 347, 347, 274, 73, 0])
    const { rows } = await upperSegmentCensusOf(undefined, REPO_ROOT, { include: 'all' })
    const byCase = new Map<string, string[]>()
    for (const x of rows) byCase.set(x.caseId, [...(byCase.get(x.caseId) ?? []), x.ref])
    let painted = 0
    for (const [caseId, refs] of byCase) {
      const c = manifest.cases.find(x => x.id === caseId)!
      if (c.stage !== 2) continue
      const fills = await withRegistration('index-position', () => productionFills(c, refs))
      for (const x of rows.filter(y => y.caseId === caseId)) { expect(fills.get(x.ref), `${caseId}:${x.ref}`).toBe(x.excel); painted++ }
    }
    // The 377 Stage 2 census cells with this capture (381 less the 4 of stage5-cfvo-nonnumeric-v2).
    expect(painted).toBe(377)
  }, 900_000)

  it('truncate registered: the committed captures refute it exactly at the census cells truncation misses (caught, never published)', async () => {
    const r = await withRegistration('truncate', () => stage2Conformance(manifest, REPO_ROOT, STAGE2_MEASURED))
    const { rows } = await upperSegmentCensusOf(undefined, REPO_ROOT, { include: 'all' })
    const misses = rows.filter(x => x.truncation !== x.excel)
    // 155 before this capture, 48 of it (the 40 decisive cells and the 8 decisive near-top probes).
    expect([misses.length, misses.filter(x => x.caseId === ID).length]).toEqual([203, 48])
    expect(r.violations.map(v => v.split(':').slice(0, 2).join(':')).sort()).toEqual(misses.map(x => `${x.caseId}:${x.ref}`).sort())
  }, HEAVY)

  it('the conformance pass and cf-compare (interactive, paged) on this case: each planted world passes unregistered and under its own registration; the other reading is caught', async () => {
    const { compareCfCase } = await import('../scripts/cf-compare')
    for (const reading of S2_UPPER_SEGMENT_READINGS) {
      const { root, manifest: m } = syntheticRoot(plantedFills(caseOf(), reading))
      const c = m.cases.find(x => x.id === ID)!
      const before = await unregistered(() => stage2Conformance(m, root, [ID]))
      expect(before.violations, reading).toEqual([])
      expect([before.total, before.withheld], reading).toEqual([207, shapeRefs.length])
      const registered = await withRegistration(reading, () => stage2Conformance(m, root, [ID]))
      expect(registered.violations, reading).toEqual([])
      expect([registered.total, registered.withheld], `${reading} registered`).toEqual([207, 0])
      // The main id alone: the near-top cells withheld and justified.
      const narrowed = await withRegistration(reading, () => stage2Conformance(m, root, [ID]), false)
      expect([narrowed.violations.length, narrowed.withheld], `${reading} main only`).toEqual([0, nearTopRefs.length])
      for (const mode of ['interactive', 'paged'] as const) {
        const a = await unregistered(() => compareCfCase(c, mode, { rootDir: root }))
        expect(`${reading} ${mode}: ${a.outcome} ${a.outcome === 'pass' ? '' : a.reason}`).toBe(`${reading} ${mode}: pass `)
        const b = await withRegistration(reading, () => compareCfCase(c, mode, { rootDir: root }))
        expect(`${reading} registered ${mode}: ${b.outcome} ${b.outcome === 'pass' ? '' : b.reason}`).toBe(`${reading} registered ${mode}: pass `)
      }
      const other = S2_UPPER_SEGMENT_READINGS.find(x => x !== reading)!
      const wrong = await withRegistration(other, () => stage2Conformance(m, root, [ID]))
      expect(wrong.violations.length, `${other} on a ${reading} world`).toBe(48)
    }
  }, 900_000)

  it('Stage 5: every measured case agrees with Excel under each registration; the only records that change are S4 and S5 of stage5-cfvo-nonnumeric-v2 (the literal-num twin rule, f = 1/4 and 1/2), newly painted at Excel\'s fill; the formula-bounded E4 and E5 stay withheld', async () => {
    const { stage5CaseRecords } = await import('../scripts/cf-refs/audit')
    const base = await unregistered(() => stage5CaseRecords(manifest, REPO_ROOT))
    expect(base.filter(r => r.compared && !r.agrees)).toEqual([])
    const committed = await stage5CaseRecords(manifest, REPO_ROOT)
    expect(committed.filter(r => r.compared && !r.agrees)).toEqual([])
    const k = (r: (typeof base)[number]): string => `${r.caseId}:${r.ref}`
    const was = new Map(base.map(r => [k(r), r]))
    for (const reading of S2_UPPER_SEGMENT_READINGS) {
      const rows = await withRegistration(reading, () => stage5CaseRecords(manifest, REPO_ROOT, { published: id => VERIFIED_CAPABILITIES.has(id) || id === CAP }))
      expect(rows.length, reading).toBe(base.length)
      expect(rows.filter(r => r.compared && !r.agrees).map(r => `${r.caseId}:${r.ref} ${r.disagreement}`), reading).toEqual([])
      const changed = rows.filter(r => JSON.stringify(r.production) !== JSON.stringify(was.get(k(r))!.production))
      expect(changed.map(k), reading).toEqual(['stage5-cfvo-nonnumeric-v2:S4', 'stage5-cfvo-nonnumeric-v2:S5'])
      for (const r of changed) {
        expect(was.get(k(r))!.production.fill, k(r)).toBe('none')
        expect(r.production.fill, k(r)).toBe(r.excel?.fill)
      }
      for (const ref of ['E4', 'E5']) expect(rows.find(r => r.caseId === 'stage5-cfvo-nonnumeric-v2' && r.ref === ref)!.production.fill, ref).toBe('none')
    }
    // As committed: the same two records changed.
    expect(committed.filter(r => JSON.stringify(r.production) !== JSON.stringify(was.get(k(r))!.production)).map(k)).toEqual(['stage5-cfvo-nonnumeric-v2:S4', 'stage5-cfvo-nonnumeric-v2:S5'])
  }, 900_000)
})

describe('the capture (Excel 16.115, 2ee3735a): index-position matched both questions', () => {
  it('verifyCfCase: verified; scale:upper-segment and scale:upper-segment-near-top both matched index-position; nothing refused', () => {
    const res = verifyCfCase(caseOf(), { rootDir: REPO_ROOT })
    expect(res.kind).toBe('verified')
    if (res.kind !== 'verified') return
    expect(res.matchedHypotheses).toEqual({ [CAP]: 'index-position', [NEAR]: 'index-position' })
    expect(res.refusedCapabilities ?? []).toEqual([])
  })
  it('sound: every pixel equals its DisplayFormat fill (207/207, minimum contrast 130); controls white; positive controls and all 105 twins as predicted; every designated cell at index-position, the 48 decisive off truncation', () => {
    const c = caseOf()
    const sc = JSON.parse(readFileSync(resolve(REPO_ROOT, c.captures[0]!.sidecarPath), 'utf8')) as { observedCells: Array<{ ref: string; fillRgb: string; bgRgb: number[]; bgContrastDelta: number }> }
    const obs = new Map(sc.observedCells.map(o => [o.ref, o]))
    expect(sc.observedCells.length).toBe(207)
    for (const o of sc.observedCells) expect(o.bgRgb.map(x => x.toString(16).toUpperCase().padStart(2, '0')).join(''), o.ref).toBe(o.fillRgb.slice(2).toUpperCase())
    expect(Math.min(...sc.observedCells.filter(o => o.fillRgb !== 'FFFFFFFF').map(o => o.bgContrastDelta))).toBe(130)
    for (const e of c.expected.cells.filter(x => !x.matched)) expect(obs.get(e.ref)!.fillRgb, e.ref).toBe('FFFFFFFF')
    for (const p of layout.positiveControls) expect(obs.get(p.ref)!.fillRgb, p.ref).toBe(p.token)
    for (const t of layout.twins) expect(obs.get(t.ref)!.fillRgb, t.ref).toBe(t.fill)
    let off = 0
    for (const cell of designated) {
      const ex = obs.get(cell.ref)!.fillRgb.toUpperCase()
      expect(ex, cell.ref).toBe(s2UpperSegmentPrediction(cell, 'index-position'))
      if (ex !== s2UpperSegmentPrediction(cell, 'truncate')) off++
    }
    expect(off).toBe(48)
  })
})

describe('a committed input survives regeneration', () => {
  it('keepCommittedArchive keeps the committed bytes when the inflated parts match', async () => {
    const { keepCommittedArchive, generateCaseArchive } = await import('../scripts/cf-refs/generate')
    const committed = new Uint8Array(readFileSync(resolve(REPO_ROOT, caseOf().input)))
    expect(keepCommittedArchive(committed, await generateCaseArchive(ID))).toBe(committed)
  }, HEAVY)
})
