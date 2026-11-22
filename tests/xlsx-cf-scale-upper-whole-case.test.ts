// @vitest-environment node
/**
 * The upper-segment whole-product follow-up: stage2-scale-upper-whole, an EXPLORATORY measurement case (the extended
 * search of scripts/cf-refs/upper-whole-candidates.ts left no survivor, so only the control reading `truncate` is
 * pre-registered and nothing is registrable from the capture). Checked here, before any capture:
 * 1. the case design (ownership, registries, size, values, statistical thresholds bit for bit, twins and positive controls);
 * 2. the independent model is production: every cell production paints today is the truncate prediction, and every cell it
 *    withholds is withheld by one of the two existing withholds;
 * 3. the evaluator on synthetic observations: a truncation world matches `truncate`; the 377564e0 departure (minus one at
 *    every exactly-whole upper-segment channel) and the float32 near miss are recorded as formula-hypothesis-none; it fails
 *    closed on broken controls;
 * 4. the rehearsal on a synthetic measured capture through the real verifier, audit, conformance pass and cf-compare:
 *    nothing is ever verified or registered, whatever the world, and production is unchanged; a move at D19 (the one painted
 *    upper-segment cell, E4's replica) is caught as a disagreement, as it must be.
 * This checks the pipeline, not Excel.
 */
import { copyFileSync, existsSync, mkdirSync, mkdtempSync, readFileSync, writeFileSync } from 'node:fs'
import { tmpdir } from 'node:os'
import { dirname, join, resolve } from 'node:path'
import { createHash } from 'node:crypto'
import { createCanvas } from 'canvas'
import { describe, expect, it } from 'vitest'
import {
  CF_STAGE2_CAPABILITIES,
  CF_STAGE2_CASES,
  CF_STAGE2_NONDISCRIMINATING,
  S2_UPPER_WHOLE_CAP,
  S2_UPPER_WHOLE_CONTROL_CAP,
  S2_UPPER_WHOLE_ID,
  S2_UPPER_WHOLE_READINGS,
  S2_UPPER_WHOLE_TWIN,
  s2UpperWholeCells,
  s2UpperWholeLayout,
  s2UpperWholeTruncation,
  type CfManifest,
  type CfOracleCase,
} from '../scripts/cf-refs/cases'
import { CF_MANIFEST_PATH } from '../scripts/cf-refs/generate'
import { computeWorksheetRangeGeometry, evaluateFormulaCapabilities, formulaControlsOf, isCellInSqref, readZipPartsSync, stage2ReadingsObservationOf, verifyCfCase } from '../scripts/cf-refs/verify'
import { upperWholePredictions } from '../scripts/cf-refs/upper-whole-candidates'
import { stage2Conformance } from './helpers/cf-stage2-conformance'
import { OfficePackage } from '../src/core/zip'
import { parseXlsx } from '../src/xlsx/parse'
import { cfDocumentEvaluationOf, prepareCfPaint } from '../src/xlsx/cf/render'
import { cfRuntimeEvidence } from '../src/xlsx/cf/predicates'
import { percentileOrderStatistic, resolveScaleThresholds } from '../src/xlsx/cf/colorScale'
import { VERIFIED_STAGE2_CAPABILITIES } from '../src/xlsx/cf/evidence'
import type { CfScalePoint } from '../src/xlsx/cf/model'

const REPO_ROOT = resolve(__dirname, '..')
const ID = S2_UPPER_WHOLE_ID
const CAP = S2_UPPER_WHOLE_CAP
const CTRL = S2_UPPER_WHOLE_CONTROL_CAP
const TWIN = S2_UPPER_WHOLE_TWIN
const HEAVY = 300_000

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
const twinK = new Map(s2UpperWholeLayout().twins.map(t => [t.ref, t.k]))
const predictions = new Map(upperWholePredictions().map(p => [p.ref, p]))

type World = 'truncate' | 'minus-one' | 'minus-one-except-D19' | 'near-miss'
/** Planted DisplayFormat fills: designated cells by `world` (controls always at truncation), twins on the line, positive controls as fixed, controls none. */
function plantedFills(c: CfOracleCase, world: World): Map<string, string> {
  const out = new Map<string, string>()
  for (const e of c.expected.cells) out.set(e.ref, e.matched ? e.fillRgb ?? '' : 'FFFFFFFF')
  for (const [ref, k] of twinK) out.set(ref, lineFill(k))
  for (const [ref, p] of predictions) {
    const moved = world === 'near-miss' ? p.nearMiss : world === 'truncate' || p.configuration === 'control' || (world === 'minus-one-except-D19' && ref === 'D19') ? p.truncate : p.minusOne
    out.set(ref, moved)
  }
  return out
}
const plantedSidecar = (c: CfOracleCase, fills: Map<string, string>): { observedCells: Array<Record<string, unknown>> } => ({
  observedCells: c.expected.cells.filter(e => fills.get(e.ref) !== undefined).map(e => ({ ref: e.ref, controlRef: e.controlRef, fillSource: 'excel-display-format', fillRgb: fills.get(e.ref) })),
})
const judge = (c: CfOracleCase, fills: Map<string, string>) => evaluateFormulaCapabilities(c.scaleReadings!, formulaControlsOf(c), stage2ReadingsObservationOf(plantedSidecar(c, fills)))

async function productionFills(c: CfOracleCase, refs: readonly string[]): Promise<Map<string, string | undefined>> {
  const doc = await parseXlsx(await OfficePackage.load(new Uint8Array(readFileSync(resolve(REPO_ROOT, c.saved)))))
  const sheet = doc.sheets.find(s => s.name === c.sheet) ?? doc.sheets[0]!
  const paint = await prepareCfPaint(sheet, { evidence: cfRuntimeEvidence, evaluation: cfDocumentEvaluationOf(doc) })
  const out = new Map<string, string | undefined>()
  for (const ref of refs) { const { row, col } = rc(ref); out.set(ref, paint?.layer.styleAt(row, col, undefined)?.fillColor?.toUpperCase()) }
  paint?.dispose()
  return out
}

describe('stage2-scale-upper-whole: the case (exploratory)', () => {
  it('is appended last; owns its three ids, every one registered non-discriminating; truncate is the only reading', async () => {
    expect(CF_STAGE2_CASES.at(-1)!.id).toBe(ID)
    for (const id of [CAP, CTRL, TWIN]) {
      expect(CF_STAGE2_CAPABILITIES.filter(e => e.capabilityId === id).map(e => e.caseId), id).toEqual([ID])
      expect(CF_STAGE2_NONDISCRIMINATING[id], id).toMatch(/./)
    }
    expect([...S2_UPPER_WHOLE_READINGS]).toEqual(['truncate'])
    const c = caseOf()
    expect(c.status).toBe('recipe-pending')
    expect(c.scaleReadings).toEqual(CF_STAGE2_CASES.find(t => t.id === ID)!.scaleReadings)
    expect(c.scaleReadings!.map(x => [x.id, x.capabilityId, Object.keys(x.hypotheses)])).toEqual([
      ['h2', CAP, ['truncate']], ['h1', CAP, ['truncate']], ['h3', CAP, ['truncate']], ['e4', CAP, ['truncate']], ['control', CTRL, ['truncate']],
    ])
    // Nothing is registrable: no production reading names any of the three ids.
    const { CF_STAGE2_PRODUCTION_READINGS } = await import('../scripts/cf-refs/audit')
    for (const id of [CAP, CTRL, TWIN]) expect(Object.prototype.hasOwnProperty.call(CF_STAGE2_PRODUCTION_READINGS, id), id).toBe(false)
  })

  it('195 cells: 127 designated (41 h2, 35 h1, 26 h3, 14 e4, 11 control), 32 twins, 28 positive controls, 8 controls; A1:J27', () => {
    const c = caseOf()
    expect(c.range).toBe('A1:J27')
    expect(c.expected.cells.length).toBe(195)
    expect(c.scaleReadings!.map(x => x.refs.length)).toEqual([41, 35, 26, 14, 11])
    expect(twinK.size).toBe(32)
    expect(c.expected.cells.filter(e => e.matched && !e.scaleApplied).length).toBe(28)
    expect(c.expected.negativeRefs.length).toBe(8)
    expect(c.markers.length).toBe(27)
    for (const e of c.expected.cells) expect(c.markers.filter(m => isCellInSqref(e.ref, m.sqref)).length, e.ref).toBeLessThanOrEqual(1)
    // Every designated cell's control is its column's row-27 cell (unformatted, under no rule).
    for (const ref of c.scaleReadings!.flatMap(x => x.refs)) expect(c.expected.cells.find(e => e.ref === ref)!.controlRef, ref).toBe(`${colOf(ref)}27`)
    for (const ref of c.expected.negativeRefs) expect(c.markers.some(m => isCellInSqref(ref, m.sqref)), ref).toBe(false)
  })

  it('the workbook: the catalogue values, twins of value k, positive controls on their dxf fills; no hidden rows, merges or errors', () => {
    const c = caseOf()
    const values = valuesOf(c)
    for (const cell of s2UpperWholeCells()) expect(values.get(cell.ref), cell.ref).toBe(Number(cell.text))
    for (const [ref, k] of twinK) expect(values.get(ref), ref).toBe(k)
    const xml = new TextDecoder().decode(partsOf(c, c.input).get('xl/worksheets/sheet1.xml')!)
    expect(xml).not.toMatch(/hidden="1"|<mergeCell|<autoFilter|t="e"/)
    const styles = new TextDecoder().decode(partsOf(c, c.input).get('xl/styles.xml')!)
    for (const p of s2UpperWholeLayout().positiveControls) {
      expect(xml, p.ref).toContain(`<conditionalFormatting sqref="${p.ref}"><cfRule type="cellIs" dxfId="`)
      expect(values.get(p.ref), p.ref).toBe(1)
      expect(styles, p.token).toContain(`<bgColor rgb="${p.token}"/>`)
    }
    for (const e of c.expected.cells) expect(typeof values.get(e.ref), e.ref).toBe('number')
  })

  it('every designated upper-segment cell of h1/h2 is a whole product on every channel; h3 and e4 as designed', () => {
    for (const p of predictions.values()) {
      if (p.configuration === 'h1' || p.configuration === 'h2') expect(p.channels.every(x => x.whole), p.ref).toBe(true)
      if (p.configuration !== 'control') expect(p.segment, p.ref).toBe('U')
    }
    // E4's shape: exact product whole, double product not (D19 = E4 of stage2-scale-thresholds).
    expect(predictions.get('D19')!.fraction).toBe('4/39')
    expect(((3 - 2.2) / (10 - 2.2)) * 156).not.toBe(16)
    // H3: one channel whole and another not, recorded per channel.
    expect(predictions.get('E3')!.channels.map(x => x.whole)).toEqual([true, true, false])
    expect(predictions.get('F6')!.channels.map(x => x.whole)).toEqual([true, false, false])
    expect(predictions.get('E19')!.channels.map(x => x.whole)).toEqual([false, false, false])
  })

  it('the E4 replica resolves to 1 / 2.2 / 10 bit for bit (production\'s resolver, PERCENTILE.INC)', () => {
    const values = valuesOf(caseOf())
    const g = [17, 18, 19, 20, 21].map(r => values.get(`D${r}`)!)
    expect(g).toEqual([1, 2, 3, 4, 10])
    const sorted = Float64Array.from([...g].sort((a, b) => a - b))
    const stats = { scopeKey: 'E', numericCount: g.length, sum: 0, mean: undefined, min: sorted[0], max: sorted[sorted.length - 1], stddevPop: undefined, sortedAsc: sorted, frequencies: new Map(), declaredCells: 5, spanCells: 5, storedCells: 5 }
    const points: CfScalePoint[] = [{ bound: { type: 'min' }, argb: 0 }, { bound: { type: 'percentile', value: 30 }, argb: 0 }, { bound: { type: 'max' }, argb: 0 }]
    expect(resolveScaleThresholds(points, stats, { hasEvidence: () => true })).toEqual({ kind: 'ready', value: [1, 2.2, 10] })
    expect(Object.is(percentileOrderStatistic(sorted, 30, 'inc'), 2.2)).toBe(true)
  })

  it('the independent model is production: painted cells are the truncate prediction; withheld ones are the two existing withholds', async () => {
    const c = caseOf()
    const fills = await productionFills(c, c.expected.cells.map(e => e.ref))
    const want = plantedFills(c, 'truncate')
    const withheld: string[] = []
    for (const e of c.expected.cells) {
      const got = fills.get(e.ref)
      if (got === undefined && e.matched) { withheld.push(e.ref); continue }
      expect(got, e.ref).toBe(e.matched ? want.get(e.ref) : undefined)
    }
    const expectWithheld = [...predictions.values()].filter(p => p.production !== 'painted').map(p => p.ref)
    expect(withheld.sort()).toEqual(expectWithheld.sort())
    expect([...predictions.values()].filter(p => p.production === 'painted').map(p => p.ref)).toEqual(['D12', 'D13', 'D14', 'D15', 'D16', 'D19'])
    for (const cell of s2UpperWholeCells().filter(x => x.designated)) expect(predictions.get(cell.ref)!.truncate).toBe(s2UpperWholeTruncation(cell))
  }, HEAVY)
})

describe('stage2-scale-upper-whole: the evaluator on synthetic observations', () => {
  it('a truncation world matches truncate for the finding and the controls', () => {
    const r = judge(caseOf(), plantedFills(caseOf(), 'truncate'))
    expect([r[CAP]!.status, r[CAP]!.matched, r[CTRL]!.status, r[CTRL]!.matched]).toEqual(['discriminating', 'truncate', 'discriminating', 'truncate'])
  })

  it('the 377564e0 departure (minus one) and the float32 near miss: formula-hypothesis-none for the finding (recorded), the controls still truncate', () => {
    for (const world of ['minus-one', 'near-miss'] as const) {
      const r = judge(caseOf(), plantedFills(caseOf(), world))
      expect([r[CAP]!.status, r[CAP]!.code], world).toEqual(['refused', 'formula-hypothesis-none'])
      expect(r[CTRL]!.matched, world).toBe('truncate')
    }
  })

  it('fails closed: a formatted control, a missing or wrong positive control, a twin rendering nothing, a moved control', () => {
    const c = caseOf()
    const fills = plantedFills(c, 'truncate')
    const edit = (ref: string, fill: string | undefined): Map<string, string> => { const m = new Map(fills); if (fill === undefined) m.delete(ref); else m.set(ref, fill); return m }
    const pc = Object.keys(c.scaleReadings!.find(x => x.id === 'h3')!.positiveControls!)[0]!
    expect(judge(c, edit('E27', 'FF14F028'))[CAP]!.status).toBe('refused')
    expect([judge(c, edit(pc, undefined))[CAP]!.status, judge(c, edit(pc, undefined))[CAP]!.code]).toEqual(['refused', 'fill-unobserved'])
    expect([judge(c, edit(pc, 'FF000000'))[CAP]!.status, judge(c, edit(pc, 'FF000000'))[CAP]!.code]).toEqual(['refused', 'formula-positive-control-failed'])
    expect(judge(c, edit([...twinK.keys()][0]!, 'FFFFFFFF'))[CAP]!.status).toBe('refused')
    const moved = judge(c, edit('A20', predictions.get('A20')!.minusOne))
    expect([moved[CTRL]!.status, moved[CTRL]!.code, moved[CAP]!.matched]).toEqual(['refused', 'formula-hypothesis-none', 'truncate'])
  })
})

// ---- a synthetic measured capture through the real verifier, audit, conformance pass and cf-compare (temp root) ----
const sha = (b: Uint8Array): string => createHash('sha256').update(b).digest('hex')

/** A temp root holding every Stage 2 case's files, with this case measured on a synthetic capture of `fills`. */
function syntheticRoot(fills: Map<string, string>): { root: string; manifest: CfManifest } {
  const root = mkdtempSync(join(tmpdir(), 'cf-s2-upperwhole-'))
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

describe('the rehearsal on a synthetic measured capture: nothing is ever registered, production is unchanged', () => {
  const disposition = (r: { perCapability: ReadonlyArray<{ capabilityId: string; disposition: string }> }, id: string) => r.perCapability.find(p => p.capabilityId === id)!.disposition
  const before = [...VERIFIED_STAGE2_CAPABILITIES].sort()
  const withoutOurs = Object.fromEntries(Object.entries(CF_STAGE2_NONDISCRIMINATING).filter(([id]) => ![CAP, CTRL].includes(id)))

  it('the real (recipe-pending) manifest: Stage 2 audits ok and complete, the three ids unverified-nondiscriminating, the verified set unchanged', async () => {
    const { auditStage2Manifest } = await import('../scripts/cf-refs/audit')
    const r = auditStage2Manifest(manifest, { rootDir: REPO_ROOT })
    expect([r.ok, r.stageComplete]).toEqual([true, true])
    for (const id of [CAP, CTRL, TWIN]) expect(disposition(r, id), id).toBe('unverified-nondiscriminating')
    expect(r.perCapability.filter(p => p.disposition === 'verified').map(p => p.capabilityId).sort()).toEqual(before)
  }, HEAVY)

  for (const world of ['truncate', 'minus-one-except-D19', 'near-miss'] as const) {
    it(`planted ${world}: the verifier records it; the audit verifies none of the three ids, even with their registry entries removed`, async () => {
      const { auditStage2Manifest } = await import('../scripts/cf-refs/audit')
      const { root, manifest: m } = syntheticRoot(plantedFills(caseOf(), world))
      const res = verifyCfCase(m.cases.find(x => x.id === ID)!, { rootDir: root })
      expect(res.kind, res.kind === 'refusal' ? `${res.code}: ${res.reason}` : '').toBe('verified')
      if (res.kind !== 'verified') return
      if (world === 'truncate') expect(res.matchedHypotheses).toEqual({ [CAP]: 'truncate', [CTRL]: 'truncate' })
      else {
        expect(res.matchedHypotheses).toEqual({ [CTRL]: 'truncate' })
        expect((res.refusedCapabilities ?? []).map(x => [x.capabilityId, x.code])).toEqual([[CAP, 'formula-hypothesis-none']])
      }
      const kept = auditStage2Manifest(m, { rootDir: root })
      expect(kept.ok).toBe(true)
      for (const id of [CAP, CTRL, TWIN]) expect(disposition(kept, id), id).toBe('unverified-nondiscriminating')
      // Even with the two questions' registry entries removed (the twin line stays a comparison), no production reading
      // exists for them: never verified.
      const bare = auditStage2Manifest(m, { rootDir: root, nonDiscriminating: withoutOurs })
      expect([disposition(bare, CAP), disposition(bare, CTRL)]).toEqual(['supported-structural', 'supported-structural'])
      for (const r of [kept, bare]) expect(r.perCapability.filter(p => p.disposition === 'verified').map(p => p.capabilityId).sort()).toEqual(before)
    }, HEAVY)
  }

  it('the conformance pass and cf-compare (interactive, paged): truncation and the departure away from D19 pass, withheld cells justified (121 of 195)', async () => {
    const { compareCfCase } = await import('../scripts/cf-compare')
    for (const world of ['truncate', 'minus-one-except-D19'] as const) {
      const { root, manifest: m } = syntheticRoot(plantedFills(caseOf(), world))
      const r = await stage2Conformance(m, root, [ID])
      expect(r.violations, world).toEqual([])
      expect([r.total, r.withheld], world).toEqual([195, 121])
      for (const mode of ['interactive', 'paged'] as const) {
        const res = await compareCfCase(m.cases.find(x => x.id === ID)!, mode, { rootDir: root })
        expect(`${world} ${mode}:${res.outcome} ${res.outcome === 'pass' ? '' : res.reason}`).toBe(`${world} ${mode}:pass `)
      }
    }
  }, 900_000)

  it('a move at D19 (E4\'s replica, painted today; the minus-one world and the near miss both move it) is caught as a disagreement: production does not follow the world', async () => {
    const { compareCfCase } = await import('../scripts/cf-compare')
    for (const world of ['minus-one', 'near-miss'] as const) {
      const fills = plantedFills(caseOf(), world)
      expect(fills.get('D19'), world).toBe(predictions.get('D19')!.minusOne)
      const { root, manifest: m } = syntheticRoot(fills)
      const r = await stage2Conformance(m, root, [ID])
      expect(r.violations, world).toEqual([`${ID}:D19: production ${predictions.get('D19')!.truncate} is not Excel's ${predictions.get('D19')!.minusOne}`])
      for (const mode of ['interactive', 'paged'] as const) expect((await compareCfCase(m.cases.find(x => x.id === ID)!, mode, { rootDir: root })).outcome, `${world} ${mode}`).toBe('fail')
    }
  }, 900_000)
})
