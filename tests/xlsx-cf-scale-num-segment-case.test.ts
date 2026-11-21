// @vitest-environment node
/**
 * The num-segment follow-up (brief 2026-11-22), Phase 1 tasks 3-4 and fix round 1: stage2-scale-num-segment, the
 * pre-registered confirmation case for the finding of capture e448d831 (T4 and V3 of stage2-scale-scope-disjoint one unit
 * higher per channel than production's truncation), and the rehearsal of every outcome's publication. No capture yet.
 * Checked here:
 * 1. the case design: ownership, isolation, size, bit-identical thresholds for the statistical twin, the twin line, and
 *    >= 3 cells separating every pair of readings in each configuration where they differ;
 * 2. the independent model of the predictions equals production (class consensus) through the whole paint pipeline;
 * 3. each reading, planted on a synthetic observation, comes back as the one match; the evaluator fails closed;
 * 4. the audit gate end to end on a synthetic measured capture (real verifyCfCase and auditStage2Manifest);
 * 5. the rehearsal of each registration: literal-num readings paint by class consensus (withholding where members
 *    disagree, justified); the statistical percent-roundtrip outcome withholds (R2), through the real interactive and
 *    paged comparison and the shared Stage 2 conformance pass, against every committed capture and this case.
 * This checks the family and production, not Excel.
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
  S2_NUM_SEGMENT,
  S2_NUM_SEGMENT_CAP,
  S2_NUM_SEGMENT_ID,
  S2_NUM_SEGMENT_READINGS,
  S2_NUM_SEGMENT_STATISTICAL_CAP,
  S2_NUM_SEGMENT_STATISTICAL_READINGS,
  S2_NUM_SEGMENT_TWIN,
  S2_NUM_SEGMENT_TWO_COLOUR,
  s2NumSegmentLayout,
  s2NumSegmentRules,
  s2NumSegmentValueAt,
  type CfManifest,
  type CfOracleCase,
} from '../scripts/cf-refs/cases'
import { CF_MANIFEST_PATH } from '../scripts/cf-refs/generate'
import { computeWorksheetRangeGeometry, evaluateFormulaCapabilities, formulaControlsOf, isCellInSqref, readZipPartsSync, stage2ReadingsObservationOf, verifyCfCase } from '../scripts/cf-refs/verify'
import { collectSegmentCells } from '../scripts/cf-refs/num-segment-candidates'
import { stage2NumSegmentShape } from '../scripts/cf-refs/stage2-withhold'
import { stage2Conformance } from './helpers/cf-stage2-conformance'
import { OfficePackage } from '../src/core/zip'
import { parseXlsx } from '../src/xlsx/parse'
import { cfDocumentEvaluationOf, prepareCfPaint } from '../src/xlsx/cf/render'
import { cfRuntimeEvidence } from '../src/xlsx/cf/predicates'
import { CF_NUM_SEGMENT_READINGS, injectCfNumSegmentReadings, numSegmentConsensusColor, percentileOrderStatistic, resolveScaleThresholds } from '../src/xlsx/cf/colorScale'
import { CF_NUM_SEGMENT_CLASSES } from '../src/xlsx/cf/numSegmentClasses'
import { VERIFIED_CAPABILITIES, VERIFIED_STAGE2_CAPABILITIES, withExactCfEvidence } from '../src/xlsx/cf/evidence'
import type { CfScalePoint } from '../src/xlsx/cf/model'
import type { XlsxSheet } from '../src/xlsx/types'

const REPO_ROOT = resolve(__dirname, '..')
const ID = S2_NUM_SEGMENT_ID
const CAP = S2_NUM_SEGMENT_CAP
const STAT = S2_NUM_SEGMENT_STATISTICAL_CAP
const HEAVY = 300_000
const FX = S2_NUM_SEGMENT
const READINGS = S2_NUM_SEGMENT_READINGS
/** The readings rehearsed through the slow end-to-end paths: the control, the percent class, the largest class, two fixed-point classes and the last. */
const LARGEST = Object.entries(CF_NUM_SEGMENT_CLASSES).sort((a, b) => b[1].length - a[1].length)[0]![0]
const SUBSET = [...new Set(['truncate', 'percent-roundtrip', LARGEST, READINGS[2]!, READINGS[3]!, READINGS.at(-1)!])]

afterEach(() => injectCfNumSegmentReadings(null))

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
/** The twin line's fill at step k: P + k·(1, −1, 1). */
const lineFill = (k: number): string => hex(0xff14f028 + k * 0x10000 - k * 0x100 + k)
const twinK = new Map(s2NumSegmentLayout().twins.map(t => [t.ref, t.k]))
const fillOfToken = (token: string): string => (token.startsWith('twin:') ? lineFill(twinK.get(token.slice('twin:'.length))!) : token)

/** Planted DisplayFormat fills: literal-num cells under `reading`, the statistical twin under `stat`, twins on the line, positive controls as fixed, controls none. */
function plantedFills(c: CfOracleCase, reading: string, stat = 'truncate'): Map<string, string> {
  const out = new Map<string, string>()
  for (const e of c.expected.cells) out.set(e.ref, e.matched ? e.fillRgb ?? '' : 'FFFFFFFF')
  for (const [ref, k] of twinK) out.set(ref, lineFill(k))
  for (const cfg of c.scaleReadings!) for (const ref of cfg.refs) out.set(ref, fillOfToken(cfg.hypotheses[cfg.capabilityId === STAT ? stat : reading]![ref]!))
  return out
}
const plantedSidecar = (c: CfOracleCase, fills: Map<string, string>): Record<string, unknown> => ({
  observedCells: c.expected.cells.map(e => ({ ref: e.ref, controlRef: e.controlRef, fillSource: 'excel-display-format', fillRgb: fills.get(e.ref) })),
})
const judge = (c: CfOracleCase, sidecar: Record<string, unknown>) => evaluateFormulaCapabilities(c.scaleReadings!, formulaControlsOf(c), stage2ReadingsObservationOf(sidecar))

async function sheetOf(c: CfOracleCase, root = REPO_ROOT): Promise<{ sheet: XlsxSheet; doc: Awaited<ReturnType<typeof parseXlsx>> }> {
  const doc = await parseXlsx(await OfficePackage.load(new Uint8Array(readFileSync(resolve(root, c.saved)))))
  return { doc, sheet: doc.sheets.find(s => s.name === c.sheet) ?? doc.sheets[0]! }
}
async function productionFills(c: CfOracleCase, refs: readonly string[]): Promise<Map<string, string | undefined>> {
  const { sheet, doc } = await sheetOf(c)
  const paint = await prepareCfPaint(sheet, { evidence: cfRuntimeEvidence, evaluation: cfDocumentEvaluationOf(doc) })
  const out = new Map<string, string | undefined>()
  for (const ref of refs) { const { row, col } = rc(ref); out.set(ref, paint?.layer.styleAt(row, col, undefined)?.fillColor?.toUpperCase()) }
  paint?.dispose()
  return out
}
const withRegistration = <T>(readings: Record<string, string>, fn: () => Promise<T>): Promise<T> => {
  injectCfNumSegmentReadings(readings)
  return withExactCfEvidence([...VERIFIED_CAPABILITIES, ...Object.keys(readings)], fn).finally(() => injectCfNumSegmentReadings(null))
}

describe('stage2-scale-num-segment: the case', () => {
  it('is appended last; owns its four ids, registered pre-capture (the two questions) and as comparisons (the 2-colour twins, the twin line)', () => {
    expect(CF_STAGE2_CASES.at(-1)!.id).toBe(ID)
    for (const id of [CAP, STAT, S2_NUM_SEGMENT_TWO_COLOUR, S2_NUM_SEGMENT_TWIN]) {
      expect(CF_STAGE2_CAPABILITIES.filter(e => e.capabilityId === id).map(e => e.caseId), id).toEqual([ID])
      expect(CF_STAGE2_NONDISCRIMINATING[id], id).toBeDefined()
    }
    expect(caseOf().status).toBe('recipe-pending')
    expect(caseOf().scaleReadings).toEqual(CF_STAGE2_CASES.find(t => t.id === ID)!.scaleReadings)
    // The readings are the committed classes (fix round 1 R1), the statistical ones truncate and percent-roundtrip.
    expect([...READINGS]).toEqual([...CF_NUM_SEGMENT_READINGS])
    expect(READINGS.length).toBe(126)
    expect([...S2_NUM_SEGMENT_STATISTICAL_READINGS]).toEqual(['truncate', 'percent-roundtrip'])
  })

  it('the production readings the audit gates on: none for the literal-num id, truncate for the statistical id', async () => {
    const { CF_STAGE2_PRODUCTION_READINGS } = await import('../scripts/cf-refs/audit')
    expect(CF_STAGE2_PRODUCTION_READINGS).toEqual({ 'scale:scope-disjoint': 'union', [STAT]: 'truncate' })
  })

  it('the rules: nine colour scales, three twin-line rules and literal cellIs positive controls for the replica\'s fills; no cell under two colour rules', () => {
    const c = caseOf()
    expect(c.markers.map(m => [m.capabilityId, m.sqref, m.scale!.cfvo.map(x => x.type + (x.val ?? '')).join(','), m.scale!.colors.join('/')])).toEqual([
      [CAP, 'B5:B28', 'num7,num217,num637', 'FF14F028/FFE61EFA/FF14F028'],
      [CAP, 'C5:C28', 'num7,num217,num637', 'FFE61EFA/FF14F028/FFE61EFA'],
      [CAP, 'D5:D8', 'num7,num637,num217', 'FF14F028/FFE61EFA/FF14F028'],
      [CAP, 'D10:D13', 'num0,num30,num60', 'FFF8696B/FFFFEB84/FF63BE7B'],
      [STAT, 'E2:E28', 'min,percentile50,max', 'FF14F028/FFE61EFA/FF14F028'],
      [S2_NUM_SEGMENT_TWO_COLOUR, 'F5:F16', 'num7,num217', 'FF14F028/FFE61EFA'],
      [S2_NUM_SEGMENT_TWO_COLOUR, 'F17:F28', 'num217,num637', 'FFE61EFA/FF14F028'],
      [CAP, 'H5:H30', 'num7,num217,num637', 'FF14F028/FFE61EFA/FF14F028'],
      [CAP, 'I5:I30', 'num7,num217,num637', 'FFE61EFA/FF14F028/FFE61EFA'],
      [S2_NUM_SEGMENT_TWIN, 'K2:K30', 'num0,num210', 'FF14F028/FFE61EFA'],
      [S2_NUM_SEGMENT_TWIN, 'L2:L30', 'num0,num210', 'FF14F028/FFE61EFA'],
      [S2_NUM_SEGMENT_TWIN, 'M2:M4', 'num0,num210', 'FF14F028/FFE61EFA'],
    ])
    for (const e of c.expected.cells) expect(c.markers.filter(m => isCellInSqref(e.ref, m.sqref)).length, e.ref).toBeLessThanOrEqual(1)
    const xml = new TextDecoder().decode(partsOf(c, c.input).get('xl/worksheets/sheet1.xml')!)
    expect(xml).not.toMatch(/hidden="1"|<mergeCell|<autoFilter|t="e"/)
    const values = valuesOf(c)
    for (const e of c.expected.cells) expect(typeof values.get(e.ref), e.ref).toBe('number')
    const pcs = c.expected.cells.filter(e => e.matched && !e.scaleApplied)
    expect(pcs.length).toBe(8)
    for (const p of pcs) {
      expect(xml, p.ref).toContain(`<conditionalFormatting sqref="${p.ref}"><cfRule type="cellIs" dxfId="`)
      expect([values.get(p.ref), p.controlRef], p.ref).toEqual([1, FX.positiveControlShared])
    }
    const styles = new TextDecoder().decode(partsOf(c, c.input).get('xl/styles.xml')!)
    for (const p of pcs) expect(styles, p.fillRgb).toContain(`<bgColor rgb="${p.fillRgb}"/>`)
    // The twin line: one cell per predicted step k, value k under num 0 / 210 P to Q.
    for (const [ref, k] of twinK) expect(values.get(ref), ref).toBe(k)
  })

  it('281 cells, under 300: 156 designated, 61 twins, 8 positive controls, 56 controls; each designated cell\'s control holds its value', () => {
    const c = caseOf()
    const values = valuesOf(c)
    expect(c.expected.cells.length).toBe(281)
    expect(c.expected.cells.filter(e => e.scaleApplied && !twinK.has(e.ref)).length).toBe(156)
    expect(twinK.size).toBe(61)
    expect(c.expected.negativeRefs.length).toBe(56)
    expect(c.range).toBe('A1:N30')
    for (const e of c.expected.cells.filter(x => x.scaleApplied && !twinK.has(x.ref))) expect(values.get(e.controlRef), `${e.ref} -> ${e.controlRef}`).toBe(values.get(e.ref))
  })

  it('values: bounds 7 / 217 / 637, twelve fractions per segment, thirteen separators (1/8 first), the replica of T4/V3', () => {
    const values = valuesOf(caseOf())
    expect([2, 3, 4].map(r => values.get(`E${r}`))).toEqual([7, 217, 637])
    expect(Array.from({ length: 12 }, (_, i) => values.get(`B${5 + i}`))).toEqual([77, 147, 42, 182, 49, 28, 70, 37, 187, 112, 59.5, 21])
    expect(Array.from({ length: 12 }, (_, i) => values.get(`B${17 + i}`))).toEqual([357, 497, 287, 567, 301, 259, 343, 277, 577, 427, 322, 245])
    expect(Array.from({ length: 13 }, (_, i) => values.get(`H${5 + i}`))).toEqual([...FX.separators])
    expect(values.get('H5')).toBe(7 + 210 / 8)
    expect(Array.from({ length: 13 }, (_, i) => values.get(`H${18 + i}`))).toEqual(FX.separators.map(v => Number((217 + 2 * (v - 7)).toFixed(6))))
    expect(['D10', 'D11', 'D12', 'D13'].map(r => values.get(r))).toEqual([10, 20, 40, 50])
  })

  it('the statistical twin resolves to the literal rules\' thresholds bit for bit (production\'s resolver, PERCENTILE.INC)', () => {
    const values = valuesOf(caseOf())
    const g = Array.from({ length: 27 }, (_, i) => values.get(`E${2 + i}`)!)
    const sorted = Float64Array.from([...g].sort((a, b) => a - b))
    const stats = { scopeKey: 'E', numericCount: g.length, sum: 0, mean: undefined, min: sorted[0], max: sorted[sorted.length - 1], stddevPop: undefined, sortedAsc: sorted, frequencies: new Map(), declaredCells: 27, spanCells: 27, storedCells: 27 }
    const points: CfScalePoint[] = [{ bound: { type: 'min' }, argb: 0 }, { bound: { type: 'percentile', value: 50 }, argb: 0 }, { bound: { type: 'max' }, argb: 0 }]
    expect(resolveScaleThresholds(points, stats, { hasEvidence: () => true })).toEqual({ kind: 'ready', value: [7, 217, 637] })
    expect(Object.is(percentileOrderStatistic(sorted, 50, 'inc'), 217)).toBe(true)
  })

  it('every pair of readings differs at 0 or >= 4 cells of each configuration, and somewhere; the 2-colour control agrees', () => {
    const cfgs = caseOf().scaleReadings!
    expect(cfgs.map(x => [x.id, x.capabilityId, x.role ?? '', x.refs.length])).toEqual([
      ['num-pqp', CAP, '', 32], ['num-qpq', CAP, '', 24], ['num-separators', CAP, '', 52], ['two-colour-control', CAP, 'negative-control', 24], ['statistical', STAT, '', 24],
    ])
    const evidence = cfgs.filter(x => !x.role)
    for (const cfg of cfgs) {
      const hs = Object.keys(cfg.hypotheses)
      for (const [i, a] of hs.entries()) for (const b of hs.slice(i + 1)) {
        const n = cfg.refs.filter(r => cfg.hypotheses[a]![r] !== cfg.hypotheses[b]![r]).length
        if (cfg.role === 'negative-control') expect(n, `${cfg.id} ${a}/${b}`).toBe(0)
        else if (n > 0) expect(n, `${cfg.id} ${a}/${b}`).toBeGreaterThanOrEqual(4)
      }
    }
    for (const [i, a] of READINGS.entries()) for (const b of READINGS.slice(i + 1)) {
      expect(evidence.some(cfg => cfg.capabilityId === CAP && cfg.refs.some(r => cfg.hypotheses[a]![r] !== cfg.hypotheses[b]![r])), `${a}/${b}`).toBe(true)
    }
    // The replica re-observes the finding: truncate predicts production's T4/V3 colours, every other reading Excel's.
    const pqp = cfgs[0]!
    expect([pqp.hypotheses.truncate!.D12, pqp.hypotheses.truncate!.D13]).toEqual(['FFCBDC81', 'FF97CD7E'])
    for (const r of READINGS.slice(1)) expect([pqp.hypotheses[r]!.D12, pqp.hypotheses[r]!.D13], r).toEqual(['FFCCDD82', 'FF98CE7F'])
  })

  it('every member of each class predicts the reading\'s fill at every literal-num cell (the class is observationally one on the case)', () => {
    const cfgs = caseOf().scaleReadings!.filter(x => x.capabilityId === CAP && !x.role)
    const rules = s2NumSegmentRules()
    for (const cfg of cfgs) for (const ref of cfg.refs) {
      const rule = rules.find(r => r.col === colOf(ref) && r.rows.includes(rowOf(ref)))!
      const v = s2NumSegmentValueAt(rule, rowOf(ref)), t = rule.thresholds, seg = v <= t[1]! ? 0 : 1
      const s = { value: v, lo: t[seg]!, hi: t[seg + 1]!, fraction: Math.max(0, Math.min(1, (v - t[seg]!) / (t[seg + 1]! - t[seg]!))), c0: Number.parseInt(rule.colors[seg]!, 16) >>> 0, c1: Number.parseInt(rule.colors[seg + 1]!, 16) >>> 0, colours: 3 as const, segment: seg as 0 | 1, loKind: 'num' as const, hiKind: 'num' as const, thresholds: [...t] }
      for (const r of READINGS) expect(hex(numSegmentConsensusColor(r, s) ?? 0), `${r} ${ref}`).toBe(fillOfToken(cfg.hypotheses[r]![ref]!))
    }
  })

  it('the independent model is production: with each reading registered (rehearsed), the input paints every prediction exactly', async () => {
    const c = caseOf()
    const refs = c.expected.cells.map(e => e.ref)
    for (const [r, s] of [...READINGS.map(x => [x, 'truncate'] as const), ['percent-roundtrip', 'truncate'] as const]) {
      const fills = await withRegistration({ [CAP]: r, [STAT]: s }, () => productionFills(c, refs))
      const want = plantedFills(c, r, s)
      for (const e of c.expected.cells) expect(fills.get(e.ref), `${r}/${s} ${e.ref}`).toBe(e.matched ? want.get(e.ref) : undefined)
    }
    // The statistical percent-roundtrip outcome withholds the statistical twin's non-dyadic interior cells (R2).
    const fills = await withRegistration({ [CAP]: READINGS[1]!, [STAT]: 'percent-roundtrip' }, () => productionFills(c, refs))
    const stat = c.scaleReadings!.find(x => x.capabilityId === STAT)!
    const withheld = stat.refs.filter(ref => fills.get(ref) === undefined)
    expect(withheld).toEqual(stat.refs.filter(ref => { const v = valuesOf(c).get(ref)!; const [lo, hi] = v <= 217 ? [7, 217] : [217, 637]; return !Number.isInteger(((v - lo) / (hi - lo)) * 1048576) }))
    expect(withheld.length).toBe(20)
  }, 600_000)

  it('production today: the literal-num shape is withheld where f·2^20 is not an integer; twins, controls and positive controls paint as predicted', async () => {
    const c = caseOf()
    const fills = await productionFills(c, c.expected.cells.map(e => e.ref))
    const want = plantedFills(c, 'truncate')
    const literal = new Set(c.scaleReadings!.filter(x => x.capabilityId === CAP && x.role !== 'negative-control').flatMap(x => x.refs))
    const withheld: string[] = []
    for (const e of c.expected.cells) {
      const got = fills.get(e.ref)
      if (got === undefined && literal.has(e.ref)) { withheld.push(e.ref); continue }
      expect(got, e.ref).toBe(e.matched ? want.get(e.ref) : undefined)
    }
    // Every literal-num cell but the halves, quarters and the separators' eighth (f·2^20 an integer) is withheld.
    const dyadic = [...literal].filter(ref => !withheld.includes(ref))
    expect(dyadic.sort()).toEqual(['B14', 'B15', 'B26', 'B27', 'C14', 'C15', 'C26', 'C27', 'H18', 'H5', 'I18', 'I5'].sort())
  }, HEAVY)
})

describe('stage2-scale-num-segment: each reading discriminates on a synthetic observation', () => {
  it('planted reading r (statistical truncate): r and truncate come back as the one matches, for every one of the readings', () => {
    const c = caseOf()
    for (const r of READINGS) {
      const ev = judge(c, plantedSidecar(c, plantedFills(c, r)))
      expect([ev[CAP]!.status, ev[CAP]!.matched], `${r}: ${ev[CAP]!.reason}`).toEqual(['discriminating', r])
      expect([ev[STAT]!.status, ev[STAT]!.matched], ev[STAT]!.reason).toEqual(['discriminating', 'truncate'])
    }
  }, HEAVY)

  it('planted statistical percent-roundtrip: the statistical id matches it, the literal-num id its own reading', () => {
    const c = caseOf()
    const ev = judge(c, plantedSidecar(c, plantedFills(c, READINGS[5]!, 'percent-roundtrip')))
    expect([ev[CAP]!.matched, ev[STAT]!.matched]).toEqual([READINGS[5], 'percent-roundtrip'])
  })

  it('fails closed: a formatted control, a missing or wrong positive control, a twin rendering nothing, a 2-colour control off truncation, or a mixed world refuses', () => {
    const c = caseOf()
    const fills = plantedFills(c, READINGS[2]!)
    const edit = (ref: string, fill: string | undefined) => {
      const m = new Map(fills)
      if (fill === undefined) m.delete(ref); else m.set(ref, fill)
      const sc = plantedSidecar(c, m) as { observedCells: Array<Record<string, unknown>> }
      return { observedCells: sc.observedCells.filter(o => o.fillRgb !== undefined) }
    }
    const pc = Object.keys(c.scaleReadings!.find(x => x.id === 'num-pqp')!.positiveControls!)[0]!
    expect(judge(c, edit('G5', 'FF14F028'))[CAP]!.status).toBe('refused')
    expect([judge(c, edit(pc, undefined))[CAP]!.status, judge(c, edit(pc, undefined))[CAP]!.code]).toEqual(['refused', 'fill-unobserved'])
    expect([judge(c, edit(pc, 'FF000000'))[CAP]!.status, judge(c, edit(pc, 'FF000000'))[CAP]!.code]).toEqual(['refused', 'formula-positive-control-failed'])
    expect(judge(c, edit([...twinK.keys()][0]!, 'FFFFFFFF'))[CAP]!.status).toBe('refused')
    const pqp = c.scaleReadings!.find(x => x.id === 'num-pqp')!
    expect(judge(c, edit('F5', fillOfToken(pqp.hypotheses[READINGS[2]!]!.B5!)))[CAP]!.status).toBe('refused')
    const mixed = edit('B9', fillOfToken(pqp.hypotheses.truncate!.B9!))
    expect([judge(c, mixed)[CAP]!.status, judge(c, mixed)[CAP]!.code]).toEqual(['refused', 'formula-hypothesis-none'])
  })
})

// ---- a synthetic measured capture through the real verifier and audit (temp root) ----
const sha = (b: Uint8Array): string => createHash('sha256').update(b).digest('hex')

/** A temp root holding every Stage 2 case's files, with this case measured on a synthetic capture of `fills`. */
function syntheticRoot(fills: Map<string, string>): { root: string; manifest: CfManifest } {
  const root = mkdtempSync(join(tmpdir(), 'cf-s2-numseg-'))
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
  const published = Object.fromEntries(Object.entries(CF_STAGE2_NONDISCRIMINATING).filter(([id]) => id !== CAP && id !== STAT))
  const disposition = (r: { perCapability: ReadonlyArray<{ capabilityId: string; disposition: string }> }, id: string) => r.perCapability.find(p => p.capabilityId === id)!.disposition
  const before = [...VERIFIED_STAGE2_CAPABILITIES].sort()

  it('the real (recipe-pending) manifest: Stage 2 audits ok and complete, the four ids unverified-nondiscriminating', async () => {
    const { auditStage2Manifest } = await import('../scripts/cf-refs/audit')
    const r = auditStage2Manifest(manifest, { rootDir: REPO_ROOT })
    expect([r.ok, r.stageComplete]).toEqual([true, true])
    for (const id of [CAP, STAT, S2_NUM_SEGMENT_TWO_COLOUR, S2_NUM_SEGMENT_TWIN]) expect(disposition(r, id), id).toBe('unverified-nondiscriminating')
    expect(r.perCapability.filter(p => p.disposition === 'verified').map(p => p.capabilityId).sort()).toEqual(before)
  }, HEAVY)

  for (const reading of SUBSET) {
    it(`planted ${reading}: verified as the match; published only with the registry entries gone and ${reading} as production's reading`, async () => {
      const { auditStage2Manifest, CF_STAGE2_PRODUCTION_READINGS } = await import('../scripts/cf-refs/audit')
      const { root, manifest: m } = syntheticRoot(plantedFills(caseOf(), reading))
      const res = verifyCfCase(m.cases.find(x => x.id === ID)!, { rootDir: root })
      expect(res.kind, res.kind === 'refusal' ? `${res.code}: ${res.reason}` : '').toBe('verified')
      if (res.kind !== 'verified') return
      expect(res.matchedHypotheses).toEqual({ [CAP]: reading, [STAT]: 'truncate' })
      expect([...res.verifiedCapabilities].sort()).toEqual([CAP, STAT].sort())
      const kept = auditStage2Manifest(m, { rootDir: root })
      expect(kept.ok).toBe(true)
      expect([disposition(kept, CAP), disposition(kept, STAT)]).toEqual(['unverified-nondiscriminating', 'unverified-nondiscriminating'])
      const unregistered = auditStage2Manifest(m, { rootDir: root, nonDiscriminating: published })
      expect([disposition(unregistered, CAP), disposition(unregistered, STAT)]).toEqual(['supported-structural', 'verified'])
      const registered = auditStage2Manifest(m, { rootDir: root, nonDiscriminating: published, productionReadings: { ...CF_STAGE2_PRODUCTION_READINGS, [CAP]: reading } })
      expect(registered.ok).toBe(true)
      expect(registered.perCapability.filter(p => p.disposition === 'verified').map(p => p.capabilityId).sort()).toEqual([...before, CAP, STAT].sort())
      for (const id of [S2_NUM_SEGMENT_TWO_COLOUR, S2_NUM_SEGMENT_TWIN]) expect(disposition(registered, id)).toBe('unverified-nondiscriminating')
      const other = READINGS.find(x => x !== reading)!
      expect(disposition(auditStage2Manifest(m, { rootDir: root, nonDiscriminating: published, productionReadings: { ...CF_STAGE2_PRODUCTION_READINGS, [CAP]: other } }), CAP)).toBe('supported-structural')
    }, HEAVY)
  }

  it('planted statistical percent-roundtrip: matched; withheld unless percent-roundtrip (the withhold) is registered as production\'s reading', async () => {
    const { auditStage2Manifest, CF_STAGE2_PRODUCTION_READINGS } = await import('../scripts/cf-refs/audit')
    const { root, manifest: m } = syntheticRoot(plantedFills(caseOf(), READINGS[2]!, 'percent-roundtrip'))
    const res = verifyCfCase(m.cases.find(x => x.id === ID)!, { rootDir: root })
    expect(res.kind === 'verified' ? res.matchedHypotheses : res).toEqual({ [CAP]: READINGS[2], [STAT]: 'percent-roundtrip' })
    expect(disposition(auditStage2Manifest(m, { rootDir: root, nonDiscriminating: published }), STAT)).toBe('supported-structural')
    expect(disposition(auditStage2Manifest(m, { rootDir: root, nonDiscriminating: published, productionReadings: { ...CF_STAGE2_PRODUCTION_READINGS, [STAT]: 'percent-roundtrip' } }), STAT)).toBe('verified')
  }, HEAVY)
})

describe('the interactive and paged comparisons of this case on a synthetic measured capture (scripts/cf-compare.ts)', () => {
  const compare = async (fills: Map<string, string>, readings: Record<string, string> | undefined): Promise<string[]> => {
    const { compareCfCase } = await import('../scripts/cf-compare')
    const { root, manifest: m } = syntheticRoot(fills)
    const c = m.cases.find(x => x.id === ID)!
    const run = async (): Promise<string[]> => {
      const out: string[] = []
      for (const mode of ['interactive', 'paged'] as const) { const res = await compareCfCase(c, mode, { rootDir: root }); out.push(`${mode}:${res.outcome}${res.outcome === 'pass' ? '' : ` ${res.reason}`}`) }
      return out
    }
    return readings ? withRegistration(readings, run) : run()
  }

  it('the literal-num outcomes: unregistered, the withheld shape is justified; under each rehearsed registration every cell matches', async () => {
    for (const reading of SUBSET) {
      expect(await compare(plantedFills(caseOf(), reading), undefined), `${reading} unregistered`).toEqual(['interactive:pass', 'paged:pass'])
      expect(await compare(plantedFills(caseOf(), reading), { [CAP]: reading }), `${reading} registered`).toEqual(['interactive:pass', 'paged:pass'])
    }
  }, 900_000)

  it('the statistical outcomes (R2): a truncate world passes unchanged; a percent-roundtrip world fails at E5 until the withhold is registered, then passes', async () => {
    const lit = READINGS[2]!
    expect(await compare(plantedFills(caseOf(), lit, 'truncate'), { [CAP]: lit })).toEqual(['interactive:pass', 'paged:pass'])
    const pr = plantedFills(caseOf(), lit, 'percent-roundtrip')
    // The review's case: without the withhold production paints truncation at E5, which the world contradicts.
    expect(await compare(pr, { [CAP]: lit })).toEqual([`interactive:fail Cell E5 observed Excel fill ${pr.get('E5')}, got ${plantedFills(caseOf(), lit).get('E5')}`, `paged:fail Cell E5 observed Excel fill ${pr.get('E5')}, got ${plantedFills(caseOf(), lit).get('E5')}`])
    expect(pr.get('E5')).not.toBe(plantedFills(caseOf(), lit).get('E5'))
    // With the statistical percent-roundtrip outcome registered, production withholds E5 (and every non-dyadic interior statistical cell), justified.
    expect(await compare(pr, { [CAP]: lit, [STAT]: 'percent-roundtrip' })).toEqual(['interactive:pass', 'paged:pass'])
  }, 900_000)
})

// ---- the rehearsal of each publication against every committed capture ----
const STAGE2_MEASURED = manifest.cases.filter(x => x.stage === 2 && x.status === 'measured').map(x => x.id)

describe('the rehearsal: each registration against every committed capture', () => {
  it('every reading, on every captured 3-colour literal-num cell: the class consensus is Excel\'s fill, or the class disagrees there and the withhold is justified', async () => {
    const { cells } = await collectSegmentCells(manifest, REPO_ROOT)
    const literal = cells.filter(c => c.segment.colours === 3 && c.segment.loKind === 'num' && c.segment.hiKind === 'num')
    expect(literal.length).toBe(24)
    const sd = manifest.cases.find(x => x.id === 'stage2-scale-scope-disjoint')!
    const withheldBy: Record<string, number> = {}
    for (const r of READINGS) {
      injectCfNumSegmentReadings({ [CAP]: r })
      await withExactCfEvidence([...VERIFIED_CAPABILITIES, CAP], async () => {
        for (const c of literal) {
          const got = numSegmentConsensusColor(r, c.segment)
          if (r === 'truncate') continue
          if (got === undefined) { withheldBy[r] = (withheldBy[r] ?? 0) + 1; expect(stage2NumSegmentShape(sd.markers, c.ref, c.segment.value), `${r} ${c.ref}: justified`).toBe(true) }
          else expect(hex(got), `${r} ${c.caseId}:${c.ref}`).toBe(c.excel)
        }
      })
      injectCfNumSegmentReadings(null)
    }
    // No class disagrees on any captured literal-num cell (each survives them all, so its members agree there).
    expect(withheldBy).toEqual({})
  }, HEAVY)

  it('today (nothing registered): the shared conformance pass has no violation; 11 num-segment cells withheld', async () => {
    const r = await stage2Conformance(manifest, REPO_ROOT, STAGE2_MEASURED)
    expect(r.violations).toEqual([])
    expect([r.total, r.numSegmentWithheld, r.statisticalWithheld]).toEqual([486, 11, 0])
  }, HEAVY)

  for (const reading of SUBSET) {
    it(`${reading} registered: the shared conformance pass ${reading === 'truncate' ? 'fails at exactly T4 and V3 (e448d831 contradicts it)' : 'paints every cell Excel\'s fill, T4 and V3 included'}; the interactive and paged comparisons agree`, async () => {
      const r = await withRegistration({ [CAP]: reading }, () => stage2Conformance(manifest, REPO_ROOT, STAGE2_MEASURED))
      if (reading === 'truncate') {
        expect(r.violations).toEqual(["stage2-scale-scope-disjoint:T4: production FFCBDC81 is not Excel's FFCCDD82", "stage2-scale-scope-disjoint:V3: production FF97CD7E is not Excel's FF98CE7F"])
        return
      }
      expect(r.violations).toEqual([])
      expect([r.numSegmentWithheld, r.statisticalWithheld]).toEqual([0, 0])
      const { compareCfCase } = await import('../scripts/cf-compare')
      await withRegistration({ [CAP]: reading }, async () => {
        for (const id of STAGE2_MEASURED) for (const mode of ['interactive', 'paged'] as const) {
          const res = await compareCfCase(manifest.cases.find(x => x.id === id)!, mode, { rootDir: REPO_ROOT })
          expect(res.outcome, `${reading} ${id} ${mode}: ${res.reason}`).toBe('pass')
        }
      })
    }, 900_000)
  }

  it('the statistical withhold registered (R2): no violation; every withheld cell is a confirmed-exact published cell, counted', async () => {
    const r = await withRegistration({ [STAT]: 'percent-roundtrip' }, () => stage2Conformance(manifest, REPO_ROOT, STAGE2_MEASURED))
    expect(r.violations).toEqual([])
    // Today these cells paint Excel's exact fill; the withhold path would stop painting them.
    expect(r.statisticalWithheld).toBe(STAT_WITHHELD_CONFIRMED)
    const { compareCfCase } = await import('../scripts/cf-compare')
    await withRegistration({ [STAT]: 'percent-roundtrip' }, async () => {
      for (const id of STAGE2_MEASURED) for (const mode of ['interactive', 'paged'] as const) {
        const res = await compareCfCase(manifest.cases.find(x => x.id === id)!, mode, { rootDir: REPO_ROOT })
        expect(res.outcome, `${id} ${mode}: ${res.reason}`).toBe('pass')
      }
    })
  }, 900_000)

  it('Stage 5: every measured case agrees with Excel with no unjustified refusal under the literal and statistical registrations', async () => {
    const { stage5CaseRecords } = await import('../scripts/cf-refs/audit')
    const base = await stage5CaseRecords(manifest, REPO_ROOT)
    expect(base.filter(r => r.compared && !r.agrees)).toEqual([])
    const key = (rows: typeof base) => rows.map(r => `${r.caseId}:${r.ref}:${r.agrees}:${JSON.stringify(r.production)}`).join('\n')
    for (const [cap, reading] of [[CAP, READINGS[2]!], [CAP, LARGEST], [STAT, 'percent-roundtrip']] as const) {
      const rows = await withRegistration({ [cap]: reading }, () => stage5CaseRecords(manifest, REPO_ROOT, { published: id => VERIFIED_CAPABILITIES.has(id) || id === cap }))
      expect(rows.filter(r => r.compared && !r.agrees).map(r => `${r.caseId}:${r.ref} ${r.disagreement}`), `${cap} ${reading}`).toEqual([])
      expect(key(rows) === key(base), `${cap} ${reading}: Stage 5 records unchanged`).toBe(true)
    }
  }, 900_000)
})

/** The confirmed-exact published Stage 2 cells the statistical withhold would stop painting (pinned; reported in the recipe). */
const STAT_WITHHELD_CONFIRMED = 15

describe('a committed input survives regeneration (codex pre-capture P2)', () => {
  it('keepCommittedArchive keeps the committed bytes when the inflated parts match', async () => {
    const { keepCommittedArchive, generateCaseArchive } = await import('../scripts/cf-refs/generate')
    const committed = new Uint8Array(readFileSync(resolve(REPO_ROOT, caseOf().input)))
    expect(keepCommittedArchive(committed, await generateCaseArchive(ID))).toBe(committed)
  }, HEAVY)
})
