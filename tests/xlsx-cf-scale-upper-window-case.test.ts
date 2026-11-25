// @vitest-environment node
/**
 * The upper-segment window follow-up (brief cf-upperwindow), Phase 1 tasks 3-4: stage2-scale-upper-window, the
 * pre-registered case for the window class of 3-colour upper segments under statistical bounds, and the rehearsal of each
 * reading's registration. Captured since ed0fccb6 (Excel 16.115): no reading matched, and 14 lower-segment / 2-colour cells
 * just below a whole step were drawn one step smaller, now withheld as the guard-boundary shape (owner ruling; see
 * tests/xlsx-cf-scale-guard-boundary.test.ts). Checked here:
 * 1. the case design: ownership, registries, size, values, statistical thresholds bit for bit against their literal twins,
 *    every designated cell's class, >= 3 decisive cells per evidence configuration, the text scope's batch estimate;
 * 2. the independent model is production, today (window and integer classes withheld) and under each registration;
 * 3. the evaluator on synthetic observations: each planted reading comes back as the one match; worlds no reading explains
 *    (the literal capture's pattern, a moved control) are refused; it fails closed on broken controls;
 * 4. the audit gate end to end on a synthetic measured capture (real verifyCfCase and auditStage2Manifest);
 * 5. each registration against every committed capture, this one included: unregistered production agrees everywhere;
 *    guarded is refuted at 17 of this capture's statistical cells, raw-floor at E4, D19 and 5 of them, so neither can be
 *    published; Stage 5 records unchanged; the interactive and paged comparisons pass unregistered.
 * This checks the case and production, not Excel. Since the raw-floor registration (stage2-scale-raw-floor, ce3f701d,
 * docs/validation/xlsx-cf.md §14.z.18) production paints the 14 guard-boundary cells at Excel's fill; this rehearsal is pinned
 * as recorded at ed0fccb6, without that registration's evidence (`asRecorded`) and over the cases measured then.
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
  S2_UPPER_WINDOW_CAP,
  S2_UPPER_WINDOW_ID,
  S2_UPPER_WINDOW_LITERAL,
  S2_UPPER_WINDOW_READINGS,
  S2_UPPER_WINDOW_TWIN,
  s2UpperWindowBlocks,
  s2UpperWindowCells,
  s2UpperWindowClass,
  s2UpperWindowLayout,
  type CfManifest,
  type CfOracleCase,
  type S2UpperWindowReading,
} from '../scripts/cf-refs/cases'
import { CF_MANIFEST_PATH } from '../scripts/cf-refs/generate'
import { computeWorksheetRangeGeometry, evaluateFormulaCapabilities, formulaControlsOf, isCellInSqref, readZipPartsSync, stage2ReadingsObservationOf, verifyCfCase } from '../scripts/cf-refs/verify'
import { upperWindowPredictions } from '../scripts/cf-refs/upper-window-candidates'
import { stage2Conformance } from './helpers/cf-stage2-conformance'
import { OfficePackage } from '../src/core/zip'
import { parseXlsx } from '../src/xlsx/parse'
import { cfDocumentEvaluationOf, prepareCfPaint } from '../src/xlsx/cf/render'
import { cfRuntimeEvidence } from '../src/xlsx/cf/predicates'
import { CF_UPPER_WINDOW_CAP, CF_UPPER_WINDOW_READINGS, injectCfUpperWindowReadings, resolveScaleThresholds } from '../src/xlsx/cf/colorScale'
import { estimateCfPrepareBytes } from '../src/xlsx/cf/service'
import { VERIFIED_CAPABILITIES, VERIFIED_STAGE2_CAPABILITIES, withExactCfEvidence } from '../src/xlsx/cf/evidence'
import type { CfScalePoint } from '../src/xlsx/cf/model'
import { recordedBeforeUpperSegmentRegistration } from './helpers/cf-upper-segment-recorded'

// These tests record production as of their own captures, before the upper-segment registration (2ee3735a).
recordedBeforeUpperSegmentRegistration()

const REPO_ROOT = resolve(__dirname, '..')
const ID = S2_UPPER_WINDOW_ID
const CAP = S2_UPPER_WINDOW_CAP
const LIT = S2_UPPER_WINDOW_LITERAL
const TWIN = S2_UPPER_WINDOW_TWIN
const HEAVY = 300_000

afterEach(() => injectCfUpperWindowReadings(null))

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
const hex = (n: number): string => (n >>> 0).toString(16).toUpperCase().padStart(8, '0')
const lineFill = (k: number): string => hex(0xff14f028 + k * 0x10000 - k * 0x100 + k)
const twinK = new Map(s2UpperWindowLayout().twins.map(t => [t.ref, t.k]))
const predictions = new Map(upperWindowPredictions().map(p => [p.ref, p]))
/** The guard-boundary cells (owner ruling after ed0fccb6): withheld by production since, whatever a window reading says. */
const GUARD_BOUNDARY = ['A3', 'A4', 'D2', 'D4', 'E7', 'E8', 'E9', 'E10', 'E12', 'E13', 'E14', 'H8', 'H18', 'I5']

/** The evidence as recorded at ed0fccb6: every published capability but the later raw-floor (scale:guard-boundary-raw-floor) and literal-lower (scale:num-segment-literal-lower) registrations'. */
const RECORDED_EVIDENCE = [...VERIFIED_CAPABILITIES].filter(id => id !== 'scale:guard-boundary-raw-floor' && id !== 'scale:num-segment-literal-lower')
const asRecorded = <T>(fn: () => T): T => withExactCfEvidence(RECORDED_EVIDENCE, fn)

/** Rehearse a registration: the reading selected through the seam and the id's evidence held (as recorded), around `fn`. */
async function withRegistration<T>(reading: S2UpperWindowReading, fn: () => Promise<T>): Promise<T> {
  injectCfUpperWindowReadings({ [CF_UPPER_WINDOW_CAP]: reading })
  try { return await withExactCfEvidence([...RECORDED_EVIDENCE, CF_UPPER_WINDOW_CAP], fn) } finally { injectCfUpperWindowReadings(null) }
}

/** A synthetic world: each designated cell's fill by `pick` (default the named reading), twins on the line, positive controls as fixed, controls white. */
function plantedFills(c: CfOracleCase, world: S2UpperWindowReading | ((ref: string) => string)): Map<string, string> {
  const out = new Map<string, string>()
  for (const e of c.expected.cells) out.set(e.ref, e.matched ? e.fillRgb ?? '' : 'FFFFFFFF')
  for (const [ref, k] of twinK) out.set(ref, lineFill(k))
  for (const [ref, p] of predictions) out.set(ref, typeof world === 'string' ? p.readings[world]! : world(ref))
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

describe('stage2-scale-upper-window: the case', () => {
  it('is appended last; owns its three ids, each registered (the question until its capture); guarded and raw-floor are the readings', async () => {
    // Appended last at its time; the raw-floor follow-up stage2-scale-raw-floor is appended after it.
    expect(CF_STAGE2_CASES.map(c => c.id).slice(-5)).toEqual([ID, 'stage2-scale-raw-floor', 'stage2-scale-literal-lower', 'stage2-scale-upper-segment', 'stage2-scale-formula-bounds'])
    for (const id of [CAP, LIT, TWIN]) {
      expect(CF_STAGE2_CAPABILITIES.filter(e => e.capabilityId === id).map(e => e.caseId), id).toEqual([ID])
      expect(CF_STAGE2_NONDISCRIMINATING[id], id).toMatch(/./)
    }
    expect(CAP).toBe(CF_UPPER_WINDOW_CAP)
    expect([...S2_UPPER_WINDOW_READINGS]).toEqual([...CF_UPPER_WINDOW_READINGS])
    const c = caseOf()
    // Captured (ed0fccb6); no reading matched, so the three ids stay registered non-discriminating.
    expect(c.status).toBe('measured')
    expect(c.scaleReadings).toEqual(CF_STAGE2_CASES.find(t => t.id === ID)!.scaleReadings)
    expect(c.scaleReadings!.map(x => [x.id, x.capabilityId, Object.keys(x.hypotheses), x.role ?? (x.record ? 'record' : 'evidence')])).toEqual([
      ['pmp', CAP, ['guarded', 'raw-floor'], 'evidence'], ['percent', CAP, ['guarded', 'raw-floor'], 'evidence'], ['mixed', CAP, ['guarded', 'raw-floor'], 'evidence'],
      ['controls', CAP, ['guarded', 'raw-floor'], 'negative-control'], ['integer', CAP, ['guarded', 'raw-floor'], 'record'], ['literal', LIT, ['guarded', 'raw-floor'], 'evidence'],
    ])
    // Nothing is registered: no production reading names any of the three ids.
    const { CF_STAGE2_PRODUCTION_READINGS } = await import('../scripts/cf-refs/audit')
    for (const id of [CAP, LIT, TWIN]) expect(Object.prototype.hasOwnProperty.call(CF_STAGE2_PRODUCTION_READINGS, id), id).toBe(false)
  })

  it('147 cells: 93 designated (11 pmp, 8 percent, 8 mixed, 37 controls, 2 integer, 27 literal), 41 twins, 4 positive controls, 9 controls; A1:K21', () => {
    const c = caseOf()
    expect(c.range).toBe('A1:K21')
    expect(c.expected.cells.length).toBe(147)
    expect(c.scaleReadings!.map(x => x.refs.length)).toEqual([11, 8, 8, 37, 2, 27])
    expect(twinK.size).toBe(41)
    expect(c.expected.cells.filter(e => e.matched && !e.scaleApplied).length).toBe(4)
    expect(c.expected.negativeRefs.length).toBe(9)
    expect(c.markers.length).toBe(19)
    for (const e of c.expected.cells) expect(c.markers.filter(m => isCellInSqref(e.ref, m.sqref)).length, e.ref).toBeLessThanOrEqual(1)
    for (const ref of c.scaleReadings!.flatMap(x => x.refs)) expect(c.expected.cells.find(e => e.ref === ref)!.controlRef, ref).toBe(`${colOf(ref)}21`)
    for (const ref of c.expected.negativeRefs) expect(c.markers.some(m => isCellInSqref(ref, m.sqref)), ref).toBe(false)
  })

  it('the text scope admits the whole range in one batch, far under the 8,388,608-byte cap', async () => {
    const { sheet } = await sheetOf(caseOf())
    const bytes = estimateCfPrepareBytes(sheet.conditionalFormatting!, [{ firstRow: 0, lastRow: 20, firstCol: 0, lastCol: 10 }])
    expect(bytes).toBe(2472920)
    expect(bytes!).toBeLessThan(8388608 * 0.5)
  })

  it('the workbook: the catalogue values, twins of value k, positive controls on their dxf fills; no hidden rows, merges or errors', () => {
    const c = caseOf()
    const values = valuesOf(c)
    for (const cell of s2UpperWindowCells()) expect(values.get(cell.ref), cell.ref).toBe(Number(cell.text))
    for (const [ref, k] of twinK) expect(values.get(ref), ref).toBe(k)
    const xml = new TextDecoder().decode(partsOf(c, c.input).get('xl/worksheets/sheet1.xml')!)
    expect(xml).not.toMatch(/hidden="1"|<mergeCell|<autoFilter|t="e"/)
    const styles = new TextDecoder().decode(partsOf(c, c.input).get('xl/styles.xml')!)
    for (const p of s2UpperWindowLayout().positiveControls) {
      expect(xml, p.ref).toContain(`<conditionalFormatting sqref="${p.ref}"><cfRule type="cellIs" dxfId="`)
      expect(values.get(p.ref), p.ref).toBe(1)
      expect(styles, p.token).toContain(`<bgColor rgb="${p.token}"/>`)
    }
    for (const e of c.expected.cells) expect(typeof values.get(e.ref), e.ref).toBe('number')
  })

  it('every statistical block resolves (production\'s resolver) to its literal twin\'s thresholds bit for bit; percent 50 over a zero minimum is max/2 in every operation order', () => {
    const values = valuesOf(caseOf())
    for (const b of s2UpperWindowBlocks().filter(x => x.cfvo.some(v => v.type !== 'num'))) {
      const scope = b.values.map((_, i) => values.get(`${b.col}${b.firstRow + i}`)!)
      const sorted = Float64Array.from([...scope].sort((x, y) => x - y))
      const stats = { scopeKey: b.id, numericCount: scope.length, sum: 0, mean: undefined, min: sorted[0], max: sorted[sorted.length - 1], stddevPop: undefined, sortedAsc: sorted, frequencies: new Map(), declaredCells: scope.length, spanCells: scope.length, storedCells: scope.length }
      const points: CfScalePoint[] = b.cfvo.map(v => ({ bound: (v.type === 'min' || v.type === 'max' ? { type: v.type } : { type: v.type, value: Number(v.val) }) as CfScalePoint['bound'], argb: 0 }))
      const got = resolveScaleThresholds(points, stats, { hasEvidence: () => true })
      expect(got.kind, b.id).toBe('ready')
      const t = got.kind === 'ready' ? got.value : []
      expect(t.map(x => Object.is(x, x) && x), b.id).toEqual(b.thresholds.map(Number))
      for (let i = 0; i < t.length; i++) expect(Object.is(t[i], Number(b.thresholds[i])), `${b.id} ${i}`).toBe(true)
    }
    for (const max of [2.2, 1.4]) for (const alt of [max / 2, max * 0.5, (max * 50) / 100, max * (50 / 100), (0 + max) / 2, 0 * 0.5 + max * 0.5, 0 + (max - 0) * 0.5]) expect(Object.is(alt, max / 2), `${max}`).toBe(true)
    // Every statistical window cell has a literal twin with the same value and thresholds (so the same double f).
    const cells = s2UpperWindowCells()
    const key = (x: (typeof cells)[number]): string => `${x.text}@${x.lo}/${x.hi}/${x.from}/${x.to}`
    const literal = new Set(cells.filter(x => x.configuration === 'literal').map(key))
    const stat = cells.filter(x => ['pmp', 'percent', 'mixed'].includes(x.configuration ?? ''))
    expect(stat.length).toBe(27)
    for (const x of stat) expect(literal.has(key(x)), x.ref).toBe(true)
  })

  it('every window cell is in the window class, every control and integer cell as designed; each evidence configuration separates the readings at >= 3 cells, below and above an integer', () => {
    const cells = s2UpperWindowCells().filter(c => c.configuration !== undefined)
    for (const c of cells) {
      const cls = s2UpperWindowClass(c)
      if (c.role === 'window') expect(cls, c.ref).toBe('window')
      else if (c.role === 'integer') expect(cls, c.ref).toBe('integer')
      else expect(cls, c.ref).toBe('none')
    }
    const { configurations } = s2UpperWindowLayout()
    const decisive = (id: string): string[] => { const cfg = configurations.find(x => x.id === id)!; return cfg.refs.filter(r => cfg.hypotheses.guarded![r] !== cfg.hypotheses['raw-floor']![r]) }
    expect(Object.fromEntries(['pmp', 'percent', 'mixed', 'literal', 'controls', 'integer'].map(id => [id, decisive(id).length]))).toEqual({ pmp: 9, percent: 4, mixed: 5, literal: 18, controls: 0, integer: 0 })
    // Above an integer (the readings agree; a move there contradicts both): 2 in pmp, 4 in percent, 3 in mixed.
    const above = (id: string): number => configurations.find(x => x.id === id)!.refs.filter(r => !decisive(id).includes(r)).length
    expect([above('pmp'), above('percent'), above('mixed')]).toEqual([2, 4, 3])
    // E4's double (the E4 replica, value 3 in 2.2 .. 10) is in the pmp configuration and decisive.
    expect(decisive('pmp')).toContain('E4')
    expect(predictions.get('E4')!.readings).toEqual({ guarded: 'FFEFE784', 'raw-floor': 'FFF0E784' })
  })

  it('the independent model is production as recorded: painted cells are the guarded prediction; window and integer cells are withheld (literal ones as num-segment), and the 14 guard-boundary cells', async () => {
    const c = caseOf()
    const fills = await asRecorded(() => productionFills(c, c.expected.cells.map(e => e.ref)))
    const want = plantedFills(c, 'guarded')
    const withheld: string[] = []
    for (const e of c.expected.cells) {
      const got = fills.get(e.ref)
      if (got === undefined && e.matched) { withheld.push(e.ref); continue }
      expect(got, e.ref).toBe(e.matched ? want.get(e.ref) : undefined)
    }
    expect(withheld.sort()).toEqual([...[...predictions.values()].filter(p => p.production !== 'painted').map(p => p.ref), ...GUARD_BOUNDARY].sort())
    expect(withheld.length).toBe(27 + 2 + 27 + 14)
  }, HEAVY)

  for (const reading of S2_UPPER_WINDOW_READINGS) {
    it(`the independent model is production under the ${reading} registration: every statistical window cell paints the reading; integer and literal cells stay withheld`, async () => {
      const c = caseOf()
      const fills = await withRegistration(reading, () => productionFills(c, c.expected.cells.map(e => e.ref)))
      const want = plantedFills(c, reading)
      for (const e of c.expected.cells) {
        const p = predictions.get(e.ref)
        const withheldStill = (p !== undefined && p.registered !== 'painted' && p.registered !== 'reading') || GUARD_BOUNDARY.includes(e.ref)
        expect(fills.get(e.ref), e.ref).toBe(!e.matched || withheldStill ? undefined : want.get(e.ref))
      }
    }, HEAVY)
  }
})

describe('stage2-scale-upper-window: the evaluator on synthetic observations', () => {
  for (const reading of S2_UPPER_WINDOW_READINGS) {
    it(`a planted ${reading} world matches exactly ${reading}, for the question and for the literal twins`, () => {
      const r = judge(caseOf(), plantedFills(caseOf(), reading))
      expect([r[CAP]!.status, r[CAP]!.matched, r[LIT]!.status, r[LIT]!.matched]).toEqual(['discriminating', reading, 'discriminating', reading])
    })
  }

  it('no reading matches: the literal capture\'s pattern (below an integer raw floor except E4\'s double, F19\'s double one smaller) is refused, and so is a mixed world', () => {
    const c = caseOf()
    // 54fc3388's literal outcomes by double: raw floor below an integer except 4/39 (guarded); 1.36 in 1.3 .. 2.2 one smaller.
    const literalPattern = (ref: string): string => {
      const p = predictions.get(ref)!
      if (p.value === '3') return p.readings.guarded!
      if (p.value === '1.36' && p.block.endsWith('binade')) return lineFill(Number.parseInt(p.readings.guarded!.slice(2, 4), 16) - 20 + 1)
      return p.readings['raw-floor']!
    }
    for (const world of [literalPattern, (ref: string) => (ref === 'E4' ? predictions.get(ref)!.readings.guarded! : predictions.get(ref)!.readings['raw-floor']!)]) {
      const r = judge(c, plantedFills(c, world))
      expect([r[CAP]!.status, r[CAP]!.code]).toEqual(['refused', 'formula-hypothesis-none'])
    }
  })

  it('fails closed: a formatted control, a missing or wrong positive control, a twin rendering nothing, a moved negative control', () => {
    const c = caseOf()
    for (const reading of S2_UPPER_WINDOW_READINGS) {
      const fills = plantedFills(c, reading)
      const edit = (ref: string, fill: string | undefined): Map<string, string> => { const m = new Map(fills); if (fill === undefined) m.delete(ref); else m.set(ref, fill); return m }
      const pc = Object.keys(c.scaleReadings!.find(x => x.id === 'pmp')!.positiveControls!)[0]!
      expect(judge(c, edit('A21', 'FF14F028'))[CAP]!.status, reading).toBe('refused')
      expect(judge(c, edit(pc, undefined))[CAP]!.code, reading).toBe('fill-unobserved')
      expect(judge(c, edit(pc, 'FF000000'))[CAP]!.code, reading).toBe('formula-positive-control-failed')
      expect(judge(c, edit([...twinK.keys()][0]!, 'FFFFFFFF'))[CAP]!.status, reading).toBe('refused')
      // A lower-segment control (A3 at step 21, its product just below an integer) drawn one smaller (step 20): the model's scope is broken, nothing matches.
      const moved = judge(c, edit('A3', lineFill(20)))
      expect([moved[CAP]!.status, moved[CAP]!.code], reading).toEqual(['refused', 'formula-hypothesis-none'])
      // A 2-colour control (E7) at raw floor: refused too.
      const two = judge(c, edit('E7', predictions.get('A11')!.readings['raw-floor']!))
      expect([two[CAP]!.status, two[CAP]!.code], reading).toEqual(['refused', 'formula-hypothesis-none'])
    }
  })
})

// ---- a synthetic measured capture through the real verifier, audit, conformance pass and cf-compare (temp root) ----
const sha = (b: Uint8Array): string => createHash('sha256').update(b).digest('hex')

function syntheticRoot(fills: Map<string, string>): { root: string; manifest: CfManifest } {
  const root = mkdtempSync(join(tmpdir(), 'cf-s2-upperwindow-'))
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
  const at = '2026-11-22T00:00:00.000Z'
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
  const published = Object.fromEntries(Object.entries(CF_STAGE2_NONDISCRIMINATING).filter(([id]) => id !== CAP))

  it('the real (measured) manifest: Stage 2 audits ok and complete, the three ids unverified-nondiscriminating, the verified set unchanged', async () => {
    const { auditStage2Manifest } = await import('../scripts/cf-refs/audit')
    const r = auditStage2Manifest(manifest, { rootDir: REPO_ROOT })
    expect([r.ok, r.stageComplete]).toEqual([true, true])
    for (const id of [CAP, LIT, TWIN]) expect(disposition(r, id), id).toBe('unverified-nondiscriminating')
    expect(r.perCapability.filter(p => p.disposition === 'verified').map(p => p.capabilityId).sort()).toEqual(before)
  }, HEAVY)

  for (const reading of S2_UPPER_WINDOW_READINGS) {
    it(`planted ${reading}: matched itself; registry kept: unverified; entry removed, unregistered: withheld; registered: verified; the other reading registered: withheld`, async () => {
      const { auditStage2Manifest, CF_STAGE2_PRODUCTION_READINGS } = await import('../scripts/cf-refs/audit')
      const { root, manifest: m } = syntheticRoot(plantedFills(caseOf(), reading))
      const res = verifyCfCase(m.cases.find(x => x.id === ID)!, { rootDir: root })
      expect(res.kind, res.kind === 'refusal' ? `${res.code}: ${res.reason}` : '').toBe('verified')
      if (res.kind !== 'verified') return
      expect(res.matchedHypotheses).toEqual({ [CAP]: reading, [LIT]: reading })
      const kept = auditStage2Manifest(m, { rootDir: root })
      expect(kept.ok).toBe(true)
      for (const id of [CAP, LIT, TWIN]) expect(disposition(kept, id), id).toBe('unverified-nondiscriminating')
      expect(disposition(auditStage2Manifest(m, { rootDir: root, nonDiscriminating: published }), CAP)).toBe('supported-structural')
      const registered = auditStage2Manifest(m, { rootDir: root, nonDiscriminating: published, productionReadings: { ...CF_STAGE2_PRODUCTION_READINGS, [CAP]: reading } })
      expect(disposition(registered, CAP)).toBe('verified')
      expect(registered.perCapability.filter(p => p.disposition === 'verified').map(p => p.capabilityId).sort()).toEqual([...before, CAP].sort())
      const other = S2_UPPER_WINDOW_READINGS.find(x => x !== reading)!
      expect(disposition(auditStage2Manifest(m, { rootDir: root, nonDiscriminating: published, productionReadings: { ...CF_STAGE2_PRODUCTION_READINGS, [CAP]: other } }), CAP)).toBe('supported-structural')
    }, HEAVY)
  }

  it('planted "no reading matches" (the literal capture\'s pattern): formula-hypothesis-none, never verified, and production keeps the withhold', async () => {
    const { auditStage2Manifest, CF_STAGE2_PRODUCTION_READINGS } = await import('../scripts/cf-refs/audit')
    const world = (ref: string): string => { const p = predictions.get(ref)!; return p.value === '3' ? p.readings.guarded! : p.readings['raw-floor']! }
    const { root, manifest: m } = syntheticRoot(plantedFills(caseOf(), world))
    const res = verifyCfCase(m.cases.find(x => x.id === ID)!, { rootDir: root })
    expect(res.kind).toBe('verified')
    if (res.kind !== 'verified') return
    expect((res.refusedCapabilities ?? []).map(x => [x.capabilityId, x.code])).toEqual([[CAP, 'formula-hypothesis-none'], [LIT, 'formula-hypothesis-none']])
    for (const reading of S2_UPPER_WINDOW_READINGS) expect(disposition(auditStage2Manifest(m, { rootDir: root, nonDiscriminating: published, productionReadings: { ...CF_STAGE2_PRODUCTION_READINGS, [CAP]: reading } }), CAP), reading).toBe('supported-structural')
    // Unregistered production on that world: every window cell withheld and justified; nothing disagrees.
    const r = await asRecorded(() => stage2Conformance(m, root, [ID]))
    expect(r.violations).toEqual([])
    expect([r.total, r.withheld, r.upperWholeWithheld, r.numSegmentWithheld, r.guardBoundaryWithheld]).toEqual([147, 70, 29, 27, 14])
  }, HEAVY)
})

/** The Stage 2 cases measured as of ed0fccb6 (the later stage2-scale-raw-floor left out). */
const STAGE2_MEASURED = manifest.cases.filter(x => x.stage === 2 && x.status === 'measured' && x.id !== 'stage2-scale-raw-floor' && x.id !== 'stage2-scale-literal-lower' && x.id !== 'stage2-scale-upper-segment').map(x => x.id)

describe('the rehearsal: each registration against every committed capture, and this case on a synthetic capture', () => {
  it('as recorded (nothing registered): no violation over the measured Stage 2 cases, this capture included; E4 and D19 withheld', async () => {
    expect(STAGE2_MEASURED).toContain(ID)
    const r = await asRecorded(() => stage2Conformance(manifest, REPO_ROOT, STAGE2_MEASURED))
    expect(r.violations).toEqual([])
    // 962 cells before + this capture's 147: 77 exact, 70 withheld (27 num-segment, 29 upper whole product, 14 guard-boundary).
    expect([r.total, r.exact, r.withheld, r.numSegmentWithheld, r.upperWholeWithheld, r.guardBoundaryWithheld]).toEqual([1109, 761, 348, 254, 50, 14])
  }, HEAVY)

  it('guarded registered: the capture refutes it, at exactly the 17 statistical window cells Excel drew one step smaller', async () => {
    const r = await withRegistration('guarded', () => stage2Conformance(manifest, REPO_ROOT, STAGE2_MEASURED))
    expect(r.violations.map(v => v.split(':').slice(0, 2).join(':'))).toEqual(['A11', 'A12', 'A13', 'A14', 'B10', 'B11', 'B12', 'B13', 'B14', 'C5', 'C6', 'C7', 'C8', 'C15', 'D9', 'D10', 'D18'].map(x => `${ID}:${x}`))
  }, HEAVY)

  it('raw-floor registered: the earlier E4 and D19 disagree, and this capture\'s E4, D8, D11 (unmoved) and B14, C15 (above an integer, moved)', async () => {
    const r = await withRegistration('raw-floor', () => stage2Conformance(manifest, REPO_ROOT, STAGE2_MEASURED))
    expect(r.violations.map(v => v.split(':').slice(0, 2).join(':'))).toEqual(['stage2-scale-thresholds:E4', 'stage2-scale-upper-whole:D19', ...['B14', 'C15', 'D8', 'D11', 'E4'].map(x => `${ID}:${x}`)])
    expect(predictions.get('E4')!.readings['raw-floor']).toBe('FFF0E784')
  }, HEAVY)

  it('the conformance pass and cf-compare (interactive, paged) on this case: each planted world passes unregistered and under its own registration', async () => {
    const { compareCfCase } = await import('../scripts/cf-compare')
    for (const reading of S2_UPPER_WINDOW_READINGS) {
      const { root, manifest: m } = syntheticRoot(plantedFills(caseOf(), reading))
      const c = m.cases.find(x => x.id === ID)!
      const unregistered = await asRecorded(() => stage2Conformance(m, root, [ID]))
      expect(unregistered.violations, reading).toEqual([])
      expect([unregistered.total, unregistered.withheld], reading).toEqual([147, 70])
      const registered = await withRegistration(reading, () => stage2Conformance(m, root, [ID]))
      expect(registered.violations, reading).toEqual([])
      expect([registered.total, registered.withheld, registered.upperWholeWithheld], `${reading} registered`).toEqual([147, 43, 2])
      for (const mode of ['interactive', 'paged'] as const) {
        const a = await asRecorded(() => compareCfCase(c, mode, { rootDir: root }))
        expect(`${reading} ${mode}: ${a.outcome} ${a.outcome === 'pass' ? '' : a.reason}`).toBe(`${reading} ${mode}: pass `)
        const b = await withRegistration(reading, () => compareCfCase(c, mode, { rootDir: root }))
        expect(`${reading} registered ${mode}: ${b.outcome} ${b.outcome === 'pass' ? '' : b.reason}`).toBe(`${reading} registered ${mode}: pass `)
      }
      // The other reading registered on this world: its window cells disagree (caught).
      const other = S2_UPPER_WINDOW_READINGS.find(x => x !== reading)!
      const wrong = await withRegistration(other, () => stage2Conformance(m, root, [ID]))
      expect(wrong.violations.length, `${other} on a ${reading} world`).toBe(18)
    }
  }, 900_000)

  it('the interactive and paged comparisons of the measured Stage 2 cases pass unregistered; under the guarded registration this capture fails (caught)', async () => {
    const { compareCfCase } = await import('../scripts/cf-compare')
    for (const c of manifest.cases.filter(x => STAGE2_MEASURED.includes(x.id))) for (const mode of ['interactive', 'paged'] as const) {
      const res = await asRecorded(() => compareCfCase(c, mode, { rootDir: REPO_ROOT }))
      expect(res.outcome, `${c.id} ${mode}: ${res.reason}`).toBe('pass')
    }
    await withRegistration('guarded', async () => {
      for (const mode of ['interactive', 'paged'] as const) {
        const res = await compareCfCase(caseOf(), mode, { rootDir: REPO_ROOT })
        expect(`${res.outcome} ${res.reason}`, mode).toBe('fail Cell A11 observed Excel fill FFA163B5, got FFA064B4')
      }
    })
  }, 900_000)

  it('Stage 5: every measured case agrees with Excel, records unchanged, under each registration (no 3-colour Stage 5 cell is in the window class)', async () => {
    const { stage5CaseRecords } = await import('../scripts/cf-refs/audit')
    const base = await asRecorded(() => stage5CaseRecords(manifest, REPO_ROOT))
    expect(base.filter(r => r.compared && !r.agrees)).toEqual([])
    const key = (rows: typeof base) => rows.map(r => `${r.caseId}:${r.ref}:${r.agrees}:${JSON.stringify(r.production)}`).join('\n')
    for (const reading of S2_UPPER_WINDOW_READINGS) {
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
