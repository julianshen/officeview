// @vitest-environment node
/**
 * The num-segment follow-up (brief 2026-11-22), Phase 1 tasks 3-4: stage2-scale-num-segment, the pre-registered confirmation
 * case for the finding of capture e448d831 (T4 and V3 of stage2-scale-scope-disjoint one unit higher per channel than
 * production's truncation), and the rehearsal of every reading's publication. It has no capture yet. Checked here:
 * 1. the case design: ownership, isolation, size, bit-identical thresholds for the statistical twin, the survivors' classes
 *    on its cells (every survivor of the candidate search coincides with exactly one reading), and >= 3 cells separating
 *    every pair of readings in each configuration where they differ;
 * 2. the independent model of the predictions equals production's implementation through the whole paint pipeline;
 * 3. each reading, planted on a synthetic observation, comes back as the one match; the evaluator fails closed;
 * 4. the audit gate end to end on a synthetic measured capture (real verifyCfCase and auditStage2Manifest);
 * 5. the rehearsal of each proposed registration (reading selected, evidence held, withhold and its justification gone):
 *    every measured Stage 2 cell paints Excel's fill (T4 and V3 included) except for `truncate`, which e448d831
 *    contradicts at exactly T4 and V3; the interactive and paged comparisons pass; Stage 5 agrees with Excel.
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
  S2_NUM_SEGMENT_TWO_COLOUR,
  s2NumSegmentRules,
  s2NumSegmentValueAt,
  type CfManifest,
  type CfOracleCase,
  type S2NumSegmentReading,
} from '../scripts/cf-refs/cases'
import { CF_MANIFEST_PATH } from '../scripts/cf-refs/generate'
import { computeWorksheetRangeGeometry, evaluateFormulaCapabilities, formulaControlsOf, isCellInSqref, readZipPartsSync, stage2ReadingsObservationOf, verifyCfCase } from '../scripts/cf-refs/verify'
import { collectSegmentCells, survivorClassesOnCase } from '../scripts/cf-refs/num-segment-candidates'
import { STAGE2_WITHHELD_FINDINGS, stage2NumSegmentShape } from '../scripts/cf-refs/stage2-withhold'
import { OfficePackage } from '../src/core/zip'
import { parseXlsx } from '../src/xlsx/parse'
import { cfDocumentEvaluationOf, prepareCfPaint } from '../src/xlsx/cf/render'
import { cfRuntimeEvidence } from '../src/xlsx/cf/predicates'
import { CF_NUM_SEGMENT_READINGS, injectCfNumSegmentReadings, percentileOrderStatistic, resolveScaleThresholds } from '../src/xlsx/cf/colorScale'
import { VERIFIED_CAPABILITIES, VERIFIED_STAGE2_CAPABILITIES, withExactCfEvidence } from '../src/xlsx/cf/evidence'
import type { CfScalePoint } from '../src/xlsx/cf/model'
import type { XlsxSheet } from '../src/xlsx/types'

const REPO_ROOT = resolve(__dirname, '..')
const ID = S2_NUM_SEGMENT_ID
const CAP = S2_NUM_SEGMENT_CAP
const STAT = S2_NUM_SEGMENT_STATISTICAL_CAP
const HEAVY = 300_000
const FX = S2_NUM_SEGMENT

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

/** Planted DisplayFormat fills: literal-num cells under `reading`, the statistical twin under `stat`, the 2-colour control and positive controls as fixed, controls none. */
function plantedFills(c: CfOracleCase, reading: S2NumSegmentReading, stat: S2NumSegmentReading = 'truncate'): Map<string, string> {
  const out = new Map<string, string>()
  for (const e of c.expected.cells) out.set(e.ref, e.matched ? e.fillRgb ?? '' : 'FFFFFFFF')
  for (const cfg of c.scaleReadings!) for (const ref of cfg.refs) out.set(ref, cfg.hypotheses[cfg.capabilityId === STAT ? stat : reading]![ref]!)
  return out
}
const plantedSidecar = (c: CfOracleCase, fills: Map<string, string>): Record<string, unknown> => ({
  observedCells: c.expected.cells.map(e => ({ ref: e.ref, controlRef: e.controlRef, fillSource: 'excel-display-format', fillRgb: fills.get(e.ref) })),
})
const judge = (c: CfOracleCase, sidecar: Record<string, unknown>) => evaluateFormulaCapabilities(c.scaleReadings!, formulaControlsOf(c), stage2ReadingsObservationOf(sidecar))

async function sheetOf(c: CfOracleCase, file = c.saved): Promise<{ sheet: XlsxSheet; doc: Awaited<ReturnType<typeof parseXlsx>> }> {
  const doc = await parseXlsx(await OfficePackage.load(new Uint8Array(readFileSync(resolve(REPO_ROOT, file)))))
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
const withRegistration = <T>(readings: Partial<Record<string, S2NumSegmentReading>>, fn: () => Promise<T>): Promise<T> => {
  injectCfNumSegmentReadings(readings)
  return withExactCfEvidence([...VERIFIED_CAPABILITIES, ...Object.keys(readings)], fn).finally(() => injectCfNumSegmentReadings(null))
}

describe('stage2-scale-num-segment: the case', () => {
  it('is appended last; owns its three ids, registered pre-capture (the two questions) and as a comparison (the 2-colour twins)', () => {
    expect(CF_STAGE2_CASES.at(-1)!.id).toBe(ID)
    for (const id of [CAP, STAT, S2_NUM_SEGMENT_TWO_COLOUR]) {
      expect(CF_STAGE2_CAPABILITIES.filter(e => e.capabilityId === id).map(e => e.caseId), id).toEqual([ID])
      expect(CF_STAGE2_NONDISCRIMINATING[id], id).toBeDefined()
    }
    expect(caseOf().status).toBe('recipe-pending')
    // The catalogue's readings are the manifest's (Lesson 14: fixed before capture), and the readings production implements.
    expect(caseOf().scaleReadings).toEqual(CF_STAGE2_CASES.find(t => t.id === ID)!.scaleReadings)
    expect([...S2_NUM_SEGMENT_READINGS]).toEqual([...CF_NUM_SEGMENT_READINGS])
    expect([...S2_NUM_SEGMENT_STATISTICAL_READINGS]).toEqual(['truncate', 'percent-roundtrip'])
  })

  it('the production readings the audit gates on: none for the literal-num id, truncate for the statistical id', async () => {
    const { CF_STAGE2_PRODUCTION_READINGS } = await import('../scripts/cf-refs/audit')
    expect(CF_STAGE2_PRODUCTION_READINGS).toEqual({ 'scale:scope-disjoint': 'union', [STAT]: 'truncate' })
  })

  it('the rules: seven colour scales and one literal cellIs positive control per predicted fill; no cell under two colour rules', () => {
    const c = caseOf()
    expect(c.markers.map(m => [m.capabilityId, m.sqref, m.scale!.cfvo.map(x => x.type + (x.val ?? '')).join(','), m.scale!.colors.join('/')])).toEqual([
      [CAP, 'B5:B28', 'num7,num217,num637', 'FF14F028/FFE61EFA/FF14F028'],
      [CAP, 'C5:C28', 'num7,num217,num637', 'FFE61EFA/FF14F028/FFE61EFA'],
      [CAP, 'D5:D8', 'num7,num637,num217', 'FF14F028/FFE61EFA/FF14F028'],
      [CAP, 'D10:D13', 'num0,num30,num60', 'FFF8696B/FFFFEB84/FF63BE7B'],
      [STAT, 'E2:E28', 'min,percentile50,max', 'FF14F028/FFE61EFA/FF14F028'],
      [S2_NUM_SEGMENT_TWO_COLOUR, 'F5:F16', 'num7,num217', 'FF14F028/FFE61EFA'],
      [S2_NUM_SEGMENT_TWO_COLOUR, 'F17:F28', 'num217,num637', 'FFE61EFA/FF14F028'],
    ])
    for (const e of c.expected.cells) expect(c.markers.filter(m => isCellInSqref(e.ref, m.sqref)).length, e.ref).toBeLessThanOrEqual(1)
    const xml = new TextDecoder().decode(partsOf(c, c.input).get('xl/worksheets/sheet1.xml')!)
    expect(xml).not.toMatch(/hidden="1"|<mergeCell|<autoFilter|t="e"/)
    const values = valuesOf(c)
    for (const e of c.expected.cells) expect(typeof values.get(e.ref), e.ref).toBe('number')
    // Every positive control is a cellIs rule on its own cell, value 1 = 1, with the dxf of its token; its control M2 holds 1 too.
    const pcs = c.expected.cells.filter(e => e.matched && !e.scaleApplied)
    expect(pcs.length).toBe(43)
    for (const p of pcs) {
      expect(xml, p.ref).toContain(`<conditionalFormatting sqref="${p.ref}"><cfRule type="cellIs" dxfId="`)
      expect([values.get(p.ref), p.controlRef], p.ref).toEqual([1, FX.positiveControlShared])
    }
    const styles = new TextDecoder().decode(partsOf(c, c.input).get('xl/styles.xml')!)
    for (const p of pcs) expect(styles, p.fillRgb).toContain(`<bgColor rgb="${p.fillRgb}"/>`)
    expect(values.get(FX.positiveControlShared)).toBe(1)
  })

  it('small: 176 cells (104 designated, 43 positive controls, 29 controls), each control holding its cell\'s value', () => {
    const c = caseOf()
    const values = valuesOf(c)
    expect(c.expected.cells.length).toBe(176)
    expect(c.expected.cells.filter(e => e.scaleApplied).length).toBe(104)
    expect(c.expected.negativeRefs.length).toBe(29)
    expect(c.range).toBe('A1:I28')
    for (const e of c.expected.cells.filter(x => x.scaleApplied)) expect(values.get(e.controlRef), `${e.ref} -> ${e.controlRef}`).toBe(values.get(e.ref))
  })

  it('values: bounds 7 / 217 / 637; twelve fractions per segment, the quarter a decimal; the replica is T4/V3\'s rule and values', () => {
    const values = valuesOf(caseOf())
    expect([2, 3, 4].map(r => values.get(`E${r}`))).toEqual([7, 217, 637])
    expect(Array.from({ length: 12 }, (_, i) => values.get(`B${5 + i}`))).toEqual([77, 147, 42, 182, 49, 28, 70, 37, 187, 112, 59.5, 21])
    expect(Array.from({ length: 12 }, (_, i) => values.get(`B${17 + i}`))).toEqual([357, 497, 287, 567, 301, 259, 343, 277, 577, 427, 322, 245])
    expect(['D10', 'D11', 'D12', 'D13'].map(r => values.get(r))).toEqual([10, 20, 40, 50])
    expect(['D5', 'D6', 'D7', 'D8'].map(r => values.get(r))).toEqual([77, 147, 357, 497])
  })

  it('the statistical twin resolves to the literal rules\' thresholds bit for bit (production\'s resolver, PERCENTILE.INC)', () => {
    const values = valuesOf(caseOf())
    const g = Array.from({ length: 27 }, (_, i) => values.get(`E${2 + i}`)!)
    const sorted = Float64Array.from([...g].sort((a, b) => a - b))
    const stats = { scopeKey: 'E', numericCount: g.length, sum: 0, mean: undefined, min: sorted[0], max: sorted[sorted.length - 1], stddevPop: undefined, sortedAsc: sorted, frequencies: new Map(), declaredCells: 27, spanCells: 27, storedCells: 27 }
    const points: CfScalePoint[] = [{ bound: { type: 'min' }, argb: 0 }, { bound: { type: 'percentile', value: 50 }, argb: 0 }, { bound: { type: 'max' }, argb: 0 }]
    const t = resolveScaleThresholds(points, stats, { hasEvidence: () => true })
    expect(t).toEqual({ kind: 'ready', value: [7, 217, 637] })
    expect(Object.is(percentileOrderStatistic(sorted, 50, 'inc'), 217)).toBe(true)
  })

  it('every survivor of the candidate search coincides on the case\'s cells with exactly one reading; truncate is no survivor', async () => {
    const { cells } = await collectSegmentCells(manifest, REPO_ROOT)
    const lit = survivorClassesOnCase(cells, '3-colour-literal-num', CAP, S2_NUM_SEGMENT_READINGS)
    expect(lit.unmatched).toEqual([])
    expect(Object.fromEntries(Object.entries(lit.classes).map(([r, ids]) => [r, ids.length]))).toEqual({ truncate: 0, 'percent-roundtrip': 5, 'fixed-256': 596, 'fixed-1000': 147, 'fixed-308': 88, 'fixed-700': 21 })
    const stat = survivorClassesOnCase(cells, '3-colour', STAT, S2_NUM_SEGMENT_STATISTICAL_READINGS)
    expect(stat.unmatched).toEqual([])
    expect(stat.classes).toEqual({ truncate: [], 'percent-roundtrip': ['percent-roundtrip/step-truncate-unguarded', 'sweep-roundtrip-100'] })
  }, HEAVY)

  it('every pair of readings differs at 0 or >= 4 cells of each configuration, and somewhere in an evidence configuration; the 2-colour control agrees', () => {
    const cfgs = caseOf().scaleReadings!
    expect(cfgs.map(x => [x.id, x.capabilityId, x.role ?? '', x.refs.length])).toEqual([
      ['num-pqp', CAP, '', 32], ['num-qpq', CAP, '', 24], ['two-colour-control', CAP, 'negative-control', 24], ['statistical', STAT, '', 24],
    ])
    for (const cfg of cfgs) {
      const hs = Object.keys(cfg.hypotheses)
      for (const [i, a] of hs.entries()) for (const b of hs.slice(i + 1)) {
        const n = cfg.refs.filter(r => cfg.hypotheses[a]![r] !== cfg.hypotheses[b]![r]).length
        if (cfg.role === 'negative-control') expect(n, `${cfg.id} ${a}/${b}`).toBe(0)
        else expect(n, `${cfg.id} ${a}/${b}`).toBeGreaterThanOrEqual(4)
      }
    }
    // The replica re-observes the finding: truncate predicts production's T4/V3 colours, every other reading Excel's.
    const pqp = cfgs[0]!
    expect([pqp.hypotheses.truncate!.D12, pqp.hypotheses.truncate!.D13]).toEqual(['FFCBDC81', 'FF97CD7E'])
    for (const r of S2_NUM_SEGMENT_READINGS.filter(x => x !== 'truncate')) expect([pqp.hypotheses[r]!.D12, pqp.hypotheses[r]!.D13], r).toEqual(['FFCCDD82', 'FF98CE7F'])
  })

  it('the independent model is production: with each reading registered (rehearsed), the input paints every prediction exactly', async () => {
    const c = caseOf()
    const refs = c.expected.cells.map(e => e.ref)
    for (const [r, s] of [...S2_NUM_SEGMENT_READINGS.map(x => [x, 'truncate'] as const), ['fixed-256', 'percent-roundtrip'] as const]) {
      const fills = await withRegistration({ [CAP]: r, [STAT]: s }, () => productionFills(c, refs))
      const want = plantedFills(c, r, s)
      for (const e of c.expected.cells) expect(fills.get(e.ref), `${r}/${s} ${e.ref}`).toBe(e.matched ? want.get(e.ref) : undefined)
    }
  }, HEAVY)

  it('production today: the literal-num shape is withheld where f·2^20 is not an integer; the twins, controls and positive controls paint as predicted', async () => {
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
    // Every literal-num cell but the halves and quarters (f·2^20 an integer) is withheld: 56 - 8 (B and C, both segments).
    expect(withheld.length).toBe(48)
    for (const ref of ['B14', 'B15', 'B26', 'B27', 'C14', 'C15', 'C26', 'C27']) expect(withheld, ref).not.toContain(ref)
  }, HEAVY)
})

describe('stage2-scale-num-segment: each reading discriminates on a synthetic observation', () => {
  for (const r of S2_NUM_SEGMENT_READINGS) {
    it(`planted ${r} (statistical truncate): ${r} and truncate come back as the one matches`, () => {
      const c = caseOf()
      const ev = judge(c, plantedSidecar(c, plantedFills(c, r)))
      expect([ev[CAP]!.status, ev[CAP]!.matched], ev[CAP]!.reason).toEqual(['discriminating', r])
      expect([ev[STAT]!.status, ev[STAT]!.matched], ev[STAT]!.reason).toEqual(['discriminating', 'truncate'])
    })
  }

  it('planted statistical percent-roundtrip: the statistical id matches it, the literal-num id its own reading', () => {
    const c = caseOf()
    const ev = judge(c, plantedSidecar(c, plantedFills(c, 'fixed-700', 'percent-roundtrip')))
    expect([ev[CAP]!.matched, ev[STAT]!.matched]).toEqual(['fixed-700', 'percent-roundtrip'])
  })

  it('fails closed: a formatted control, a missing or wrong positive control, a 2-colour control off truncation, or a world no reading explains refuses', () => {
    const c = caseOf()
    const fills = plantedFills(c, 'fixed-256')
    const edit = (ref: string, fill: string | undefined) => {
      const m = new Map(fills)
      if (fill === undefined) m.delete(ref); else m.set(ref, fill)
      const sc = plantedSidecar(c, m) as { observedCells: Array<Record<string, unknown>> }
      return { observedCells: sc.observedCells.filter(o => o.fillRgb !== undefined) }
    }
    // A positive control of the literal-num configurations (the bounds' own colours P and Q are the statistical twin's only).
    const pc = Object.keys(c.scaleReadings!.find(x => x.id === 'num-pqp')!.positiveControls!)[0]!
    expect(judge(c, edit('G5', 'FF14F028'))[CAP]!.status).toBe('refused')
    expect([judge(c, edit(pc, undefined))[CAP]!.status, judge(c, edit(pc, undefined))[CAP]!.code]).toEqual(['refused', 'fill-unobserved'])
    expect([judge(c, edit(pc, 'FF000000'))[CAP]!.status, judge(c, edit(pc, 'FF000000'))[CAP]!.code]).toEqual(['refused', 'formula-positive-control-failed'])
    // The 2-colour negative control lowered (a planted fixed-256 colour there): no reading holds.
    expect(judge(c, edit('F5', c.scaleReadings!.find(x => x.id === 'num-pqp')!.hypotheses['fixed-256']!.B5!))[CAP]!.status).toBe('refused')
    // B5 under fixed-256 but B9 under truncate (a mixed world): no reading holds, a finding, never fitted.
    const mixed = edit('B9', c.scaleReadings!.find(x => x.id === 'num-pqp')!.hypotheses.truncate!.B9!)
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

  it('the real (recipe-pending) manifest: Stage 2 audits ok and complete, the three ids unverified-nondiscriminating', async () => {
    const { auditStage2Manifest } = await import('../scripts/cf-refs/audit')
    const r = auditStage2Manifest(manifest, { rootDir: REPO_ROOT })
    expect([r.ok, r.stageComplete]).toEqual([true, true])
    for (const id of [CAP, STAT, S2_NUM_SEGMENT_TWO_COLOUR]) expect(disposition(r, id), id).toBe('unverified-nondiscriminating')
    expect(r.perCapability.filter(p => p.disposition === 'verified').map(p => p.capabilityId).sort()).toEqual(before)
  }, HEAVY)

  for (const reading of S2_NUM_SEGMENT_READINGS) {
    it(`planted ${reading}: verified as the match; published only with the registry entries gone and ${reading} as production's reading`, async () => {
      const { auditStage2Manifest, CF_STAGE2_PRODUCTION_READINGS } = await import('../scripts/cf-refs/audit')
      const { root, manifest: m } = syntheticRoot(plantedFills(caseOf(), reading))
      const res = verifyCfCase(m.cases.find(x => x.id === ID)!, { rootDir: root })
      expect(res.kind, res.kind === 'refusal' ? `${res.code}: ${res.reason}` : '').toBe('verified')
      if (res.kind !== 'verified') return
      expect(res.matchedHypotheses).toEqual({ [CAP]: reading, [STAT]: 'truncate' })
      expect([...res.verifiedCapabilities].sort()).toEqual([CAP, STAT].sort())
      // Registry entries kept: measured, never verified.
      const kept = auditStage2Manifest(m, { rootDir: root })
      expect(kept.ok).toBe(true)
      expect([disposition(kept, CAP), disposition(kept, STAT)]).toEqual(['unverified-nondiscriminating', 'unverified-nondiscriminating'])
      // Entries gone but no reading registered: the literal-num id is withheld by the production-reading gate.
      const unregistered = auditStage2Manifest(m, { rootDir: root, nonDiscriminating: published })
      expect([disposition(unregistered, CAP), disposition(unregistered, STAT)]).toEqual(['supported-structural', 'verified'])
      // The proposed registration: exactly the two ids join the verified Stage 2 set.
      const registered = auditStage2Manifest(m, { rootDir: root, nonDiscriminating: published, productionReadings: { ...CF_STAGE2_PRODUCTION_READINGS, [CAP]: reading } })
      expect(registered.ok).toBe(true)
      expect(registered.perCapability.filter(p => p.disposition === 'verified').map(p => p.capabilityId).sort()).toEqual([...before, CAP, STAT].sort())
      expect(disposition(registered, S2_NUM_SEGMENT_TWO_COLOUR)).toBe('unverified-nondiscriminating')
      // Another reading registered than the one matched: withheld.
      const other = S2_NUM_SEGMENT_READINGS.find(x => x !== reading)!
      expect(disposition(auditStage2Manifest(m, { rootDir: root, nonDiscriminating: published, productionReadings: { ...CF_STAGE2_PRODUCTION_READINGS, [CAP]: other } }), CAP)).toBe('supported-structural')
    }, HEAVY)
  }

  it('planted statistical percent-roundtrip: matched, but production reads truncate, so the statistical id is withheld', async () => {
    const { auditStage2Manifest } = await import('../scripts/cf-refs/audit')
    const { root, manifest: m } = syntheticRoot(plantedFills(caseOf(), 'fixed-256', 'percent-roundtrip'))
    const res = verifyCfCase(m.cases.find(x => x.id === ID)!, { rootDir: root })
    expect(res.kind === 'verified' ? res.matchedHypotheses : res).toEqual({ [CAP]: 'fixed-256', [STAT]: 'percent-roundtrip' })
    expect(disposition(auditStage2Manifest(m, { rootDir: root, nonDiscriminating: published }), STAT)).toBe('supported-structural')
  }, HEAVY)
})

// ---- the rehearsal of each publication against every committed capture ----
const STAGE2_MEASURED = manifest.cases.filter(x => x.stage === 2 && x.status === 'measured')

/** Every measured Stage 2 observed cell (non-hidden DisplayFormat): production's fill against Excel's, under whatever is injected. */
async function stage2Disagreements(): Promise<{ disagreements: string[]; exact: number; withheld: number }> {
  const disagreements: string[] = []
  let exact = 0, withheld = 0
  for (const c of STAGE2_MEASURED) {
    const sc = JSON.parse(readFileSync(resolve(REPO_ROOT, c.captures[0]!.sidecarPath), 'utf8')) as { observedCells: Array<{ ref: string; fillSource: string; fillRgb: string; hidden?: boolean }> }
    const observed = sc.observedCells.filter(o => o.fillSource === 'excel-display-format' && !o.hidden)
    const { sheet, doc } = await sheetOf(c)
    const paint = await prepareCfPaint(sheet, { evidence: cfRuntimeEvidence, evaluation: cfDocumentEvaluationOf(doc) })
    for (const o of observed) {
      const { row, col } = rc(o.ref)
      const cell = sheet.rows.find(r => r.index === row)?.cells.find(x => x.col === col)
      const st = paint!.layer.styleAt(row, col, { numFmtId: 0, ...(cell?.style ?? {}), fillColor: cell?.style?.fillColor ?? 'FFFFFFFF' })
      const got = (st?.fillColor ?? 'FFFFFFFF').toUpperCase()
      const key = `${c.id}:${o.ref}`
      if (got === o.fillRgb.toUpperCase()) { exact++; continue }
      // A withheld cell is accepted only where a finding or the num-segment shape (while unregistered) justifies it.
      if (got === 'FFFFFFFF' && (STAGE2_WITHHELD_FINDINGS[key] !== undefined || stage2NumSegmentShape(c.markers, o.ref, cell?.value))) { withheld++; continue }
      disagreements.push(`${key} ${got}≠${o.fillRgb}`)
    }
    paint!.dispose()
  }
  return { disagreements, exact, withheld }
}

describe('the interactive and paged comparisons of this case on a synthetic measured capture', () => {
  it('today: the withheld shape is justified (num-segment), everything else matches; under each registration every cell matches', async () => {
    const { compareCfCase } = await import('../scripts/cf-compare')
    for (const reading of [undefined, ...S2_NUM_SEGMENT_READINGS]) {
      const { root, manifest: m } = syntheticRoot(plantedFills(caseOf(), reading ?? 'fixed-256'))
      const c = m.cases.find(x => x.id === ID)!
      const run = async (): Promise<void> => {
        for (const mode of ['interactive', 'paged'] as const) {
          const res = await compareCfCase(c, mode, { rootDir: root })
          expect(res.outcome, `${reading ?? 'unregistered'} ${mode}: ${res.reason}`).toBe('pass')
        }
      }
      if (reading === undefined) await run()
      else await withRegistration({ [CAP]: reading }, run)
    }
  }, 600_000)
})

describe('the rehearsal: each reading registered, against every committed capture', () => {
  it('today (nothing registered): no disagreement; the 11 num-segment cells and the 34 findings are the withheld ones', async () => {
    const r = await stage2Disagreements()
    expect(r.disagreements).toEqual([])
    // 486 observed cells: 445 exact (Excel white at 4 of the findings' cells), 41 withheld (30 findings, 11 num-segment).
    expect([r.exact, r.withheld]).toEqual([445, 41])
  }, HEAVY)

  for (const reading of S2_NUM_SEGMENT_READINGS) {
    it(`${reading} registered: the num-segment justification is gone; ${reading === 'truncate' ? 'exactly T4 and V3 disagree (e448d831 contradicts it)' : 'every cell paints Excel\'s fill, T4 and V3 included'}`, async () => {
      const r = await withRegistration({ [CAP]: reading }, async () => {
        expect(stage2NumSegmentShape(manifest.cases.find(x => x.id === 'stage2-scale-scope-disjoint')!.markers, 'T4', 40)).toBe(false)
        return stage2Disagreements()
      })
      if (reading === 'truncate') {
        expect(r.disagreements).toEqual(['stage2-scale-scope-disjoint:T4 FFCBDC81≠FFCCDD82', 'stage2-scale-scope-disjoint:V3 FF97CD7E≠FF98CE7F'])
      } else {
        expect(r.disagreements).toEqual([])
        expect([r.exact, r.withheld]).toEqual([456, 30])
      }
    }, HEAVY)
  }

  it('statistical percent-roundtrip registered: every Stage 2 cell still paints Excel\'s fill (it survives every 3-colour capture)', async () => {
    const r = await withRegistration({ [STAT]: 'percent-roundtrip' }, () => stage2Disagreements())
    expect(r.disagreements).toEqual([])
    expect([r.exact, r.withheld]).toEqual([445, 41])
  }, HEAVY)

  it('the interactive and paged comparisons of the measured Stage 2 cases pass under each non-truncate registration', async () => {
    const { compareCfCase } = await import('../scripts/cf-compare')
    for (const reading of S2_NUM_SEGMENT_READINGS.filter(x => x !== 'truncate')) {
      await withRegistration({ [CAP]: reading }, async () => {
        for (const c of STAGE2_MEASURED) for (const mode of ['interactive', 'paged'] as const) {
          const res = await compareCfCase(c, mode, { rootDir: REPO_ROOT })
          expect(res.outcome, `${reading} ${c.id} ${mode}: ${res.reason}`).toBe('pass')
        }
      })
    }
  }, 600_000)

  it('Stage 5: every measured case agrees with Excel with no unjustified refusal under each registration (no Stage 5 cell is in the literal-num shape)', async () => {
    const { stage5CaseRecords } = await import('../scripts/cf-refs/audit')
    const base = await stage5CaseRecords(manifest, REPO_ROOT)
    expect(base.filter(r => r.compared && !r.agrees)).toEqual([])
    const key = (rows: typeof base) => rows.map(r => `${r.caseId}:${r.ref}:${r.agrees}:${JSON.stringify(r.production)}`).join('\n')
    for (const [cap, reading] of [[CAP, 'fixed-256'], [CAP, 'percent-roundtrip'], [STAT, 'percent-roundtrip']] as const) {
      const rows = await withRegistration({ [cap]: reading }, () => stage5CaseRecords(manifest, REPO_ROOT, { published: id => VERIFIED_CAPABILITIES.has(id) || id === cap }))
      expect(rows.filter(r => r.compared && !r.agrees).map(r => `${r.caseId}:${r.ref} ${r.disagreement}`), `${cap} ${reading}`).toEqual([])
      expect(key(rows) === key(base), `${cap} ${reading}: Stage 5 records unchanged`).toBe(true)
    }
  }, 900_000)
})

describe('a committed input survives regeneration (codex pre-capture P2)', () => {
  it('keepCommittedArchive keeps the committed bytes when the inflated parts match', async () => {
    const { keepCommittedArchive, generateCaseArchive } = await import('../scripts/cf-refs/generate')
    const committed = new Uint8Array(readFileSync(resolve(REPO_ROOT, caseOf().input)))
    expect(keepCommittedArchive(committed, await generateCaseArchive(ID))).toBe(committed)
    // The rules and values are the catalogue's (the generator reads the same layout).
    const rule = s2NumSegmentRules().find(r => r.id === 'num-replica')!
    expect(rule.rows.map(row => s2NumSegmentValueAt(rule, row))).toEqual([10, 20, 40, 50])
  }, HEAVY)
})
