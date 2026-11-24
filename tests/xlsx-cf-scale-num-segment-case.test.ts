// @vitest-environment node
/**
 * The num-segment follow-up: stage2-scale-num-segment, the pre-registered confirmation case for the finding of capture
 * e448d831, now captured (Excel 16.115). Checked here:
 * 1. the case design (ownership, size, values, bit-identical statistical thresholds, separation, class membership);
 * 2. the evaluator: each reading planted on a synthetic observation comes back as the one match; it fails closed;
 * 3. the capture: no pre-registered reading matches either question (formula-hypothesis-none, recorded findings, never
 *    verified); the 2-colour twins and the twin line are non-discriminating comparisons; the capture rendered as designed
 *    (controls white, positive controls and the twin line exact); default production paints every cell Excel's fill or
 *    withholds it with a recorded justification, and the interactive and paged comparisons pass.
 * The arithmetic finding itself is pinned in tests/xlsx-cf-scale-upper-whole-product.test.ts.
 */
import { readFileSync } from 'node:fs'
import { resolve } from 'node:path'
import { describe, expect, it } from 'vitest'
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
import { evaluateFormulaCapabilities, formulaControlsOf, isCellInSqref, readZipPartsSync, stage2ReadingsObservationOf, verifyCfCase } from '../scripts/cf-refs/verify'
import { stage2Conformance } from './helpers/cf-stage2-conformance'
import { percentileOrderStatistic, resolveScaleThresholds } from '../src/xlsx/cf/colorScale'
import { CF_NUM_SEGMENT_CLASSES } from '../scripts/cf-refs/num-segment-classes'
import { numSegmentCandidate } from '../scripts/cf-refs/num-segment-formulas'
import { VERIFIED_STAGE2_CAPABILITIES } from '../src/xlsx/cf/evidence'
import type { CfScalePoint } from '../src/xlsx/cf/model'

const REPO_ROOT = resolve(__dirname, '..')
const ID = S2_NUM_SEGMENT_ID
const CAP = S2_NUM_SEGMENT_CAP
const STAT = S2_NUM_SEGMENT_STATISTICAL_CAP
const HEAVY = 300_000
const FX = S2_NUM_SEGMENT
const READINGS = S2_NUM_SEGMENT_READINGS
/** The readings rehearsed through the slow end-to-end paths: the control, the percent class, the largest class, two fixed-point classes and the last. */


const manifest = JSON.parse(readFileSync(CF_MANIFEST_PATH, 'utf8')) as CfManifest
const caseOf = (): CfOracleCase => {
  const c = manifest.cases.find(x => x.id === ID)
  if (!c) throw new Error(`${ID} is not in the manifest (generate it with generate.ts --case ${ID})`)
  return c
}
const colOf = (ref: string): string => /^[A-Z]+/.exec(ref)![0]
const rowOf = (ref: string): number => Number(/\d+$/.exec(ref)![0])
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

describe('stage2-scale-num-segment: the case', () => {
  it('is appended last; owns its four ids, registered as recorded findings (the two questions) and as comparisons (the 2-colour twins, the twin line)', () => {
    // Appended last at its time; the exploratory follow-up stage2-scale-upper-whole and the window follow-up are appended after it.
    expect(CF_STAGE2_CASES.map(c => c.id).slice(-5)).toEqual([ID, 'stage2-scale-upper-whole', 'stage2-scale-upper-window', 'stage2-scale-raw-floor', 'stage2-scale-literal-lower'])
    for (const id of [CAP, STAT, S2_NUM_SEGMENT_TWO_COLOUR, S2_NUM_SEGMENT_TWIN]) {
      expect(CF_STAGE2_CAPABILITIES.filter(e => e.capabilityId === id).map(e => e.caseId), id).toEqual([ID])
      expect(CF_STAGE2_NONDISCRIMINATING[id], id).toBeDefined()
    }
    expect(caseOf().status).toBe('measured')
    expect(caseOf().scaleReadings).toEqual(CF_STAGE2_CASES.find(t => t.id === ID)!.scaleReadings)
    // The readings are the committed classes (fix round 1 R1), the statistical ones truncate and percent-roundtrip.
    expect(READINGS.length).toBe(126)
    expect([...S2_NUM_SEGMENT_STATISTICAL_READINGS]).toEqual(['truncate', 'percent-roundtrip'])
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

  it('every member of each class predicts the reading\'s fill at every literal-num cell (the class is observationally one on the case)', () => {
    const cfgs = caseOf().scaleReadings!.filter(x => x.capabilityId === CAP && !x.role)
    const rules = s2NumSegmentRules()
    expect(Object.keys(CF_NUM_SEGMENT_CLASSES)).toEqual([...READINGS])
    for (const cfg of cfgs) for (const ref of cfg.refs) {
      const rule = rules.find(r => r.col === colOf(ref) && r.rows.includes(rowOf(ref)))!
      const v = s2NumSegmentValueAt(rule, rowOf(ref)), t = rule.thresholds, seg = v <= t[1]! ? 0 : 1
      const s = { value: v, lo: t[seg]!, hi: t[seg + 1]!, fraction: Math.max(0, Math.min(1, (v - t[seg]!) / (t[seg + 1]! - t[seg]!))), c0: Number.parseInt(rule.colors[seg]!, 16) >>> 0, c1: Number.parseInt(rule.colors[seg + 1]!, 16) >>> 0, colours: 3 as const, segment: seg as 0 | 1, loKind: 'num' as const, hiKind: 'num' as const, thresholds: [...t] }
      for (const r of READINGS) for (const id of CF_NUM_SEGMENT_CLASSES[r]!) expect(hex(numSegmentCandidate(id)!.colour(s)), `${r} ${id} ${ref}`).toBe(fillOfToken(cfg.hypotheses[r]![ref]!))
    }
  })
})

const excelFills = (): Map<string, string> => {
  const sc = JSON.parse(readFileSync(resolve(REPO_ROOT, caseOf().captures[0]!.sidecarPath), 'utf8')) as { observedCells: Array<{ ref: string; fillRgb: string; fillSource: string }> }
  return new Map(sc.observedCells.map(o => [o.ref, o.fillRgb]))
}

describe('the capture (Excel 16.115): no reading matches; production withholds or paints Excel\'s fill', () => {
  it('measured: verifyCfCase refuses both questions with formula-hypothesis-none; the comparison ids are non-discriminating; nothing verifies', () => {
    const c = caseOf()
    expect(c.status).toBe('measured')
    const res = verifyCfCase(c, { rootDir: REPO_ROOT })
    expect(res.kind, res.kind === 'refusal' ? `${res.code}: ${res.reason}` : '').toBe('verified')
    if (res.kind !== 'verified') return
    expect(res.verifiedCapabilities).toEqual([])
    expect(res.matchedHypotheses ?? {}).toEqual({})
    expect((res.refusedCapabilities ?? []).map(r => [r.capabilityId, r.code]).sort()).toEqual([[CAP, 'formula-hypothesis-none'], [STAT, 'formula-hypothesis-none']])
    expect([...(res.nonDiscriminatingCapabilities ?? [])].sort()).toEqual([S2_NUM_SEGMENT_TWO_COLOUR, S2_NUM_SEGMENT_TWIN].sort())
  })

  it('the capture rendered as designed: controls white, positive controls their dxf fill, every twin its model, every 2-colour twin truncation', () => {
    const c = caseOf()
    const excel = excelFills()
    for (const e of c.expected.cells.filter(x => !x.matched)) expect(excel.get(e.ref), e.ref).toBe('FFFFFFFF')
    for (const e of c.expected.cells.filter(x => x.matched && !x.scaleApplied)) expect(excel.get(e.ref), e.ref).toBe(e.fillRgb)
    for (const [ref, k] of twinK) expect(excel.get(ref), ref).toBe(lineFill(k))
    const two = c.scaleReadings!.find(x => x.id === 'two-colour-control')!
    for (const ref of two.refs) expect(excel.get(ref), ref).toBe(fillOfToken(two.hypotheses.truncate![ref]!))
  })

  it('the audit: Stage 2 ok and complete; the four ids stay recorded findings (unverified-nondiscriminating); the verified Stage 2 set is unchanged', async () => {
    const { auditStage2Manifest, CF_STAGE2_PRODUCTION_READINGS } = await import('../scripts/cf-refs/audit')
    // scale:scope-disjoint (union) and, since stage2-scale-raw-floor (ce3f701d), the guard-boundary question (raw-floor).
    expect(CF_STAGE2_PRODUCTION_READINGS).toEqual({ 'scale:scope-disjoint': 'union', 'scale:guard-boundary-raw-floor': 'raw-floor', 'scale:num-segment-literal-lower': 'lower-arithmetic' })
    const r = auditStage2Manifest(manifest, { rootDir: REPO_ROOT })
    expect([r.ok, r.stageComplete]).toEqual([true, true])
    for (const id of [CAP, STAT, S2_NUM_SEGMENT_TWO_COLOUR, S2_NUM_SEGMENT_TWIN]) expect(r.perCapability.find(p => p.capabilityId === id)!.disposition, id).toBe('unverified-nondiscriminating')
    expect(r.perCapability.filter(p => p.disposition === 'verified').map(p => p.capabilityId).sort()).toEqual([...VERIFIED_STAGE2_CAPABILITIES].sort())
  }, HEAVY)

  it('the shared Stage 2 conformance pass over this case: no violation; every withheld cell is justified (48 num-segment upper-segment cells since the literal-lower registration, a3f85309; 13 upper-segment whole product)', async () => {
    const r = await stage2Conformance(manifest, REPO_ROOT, [ID])
    expect(r.violations).toEqual([])
    expect([r.total, r.numSegmentWithheld, r.upperWholeWithheld, r.withheld, r.exact]).toEqual([281, 48, 13, 61, 220])
  }, HEAVY)

  it('the interactive and paged comparisons (scripts/cf-compare.ts) pass', async () => {
    const { compareCfCase } = await import('../scripts/cf-compare')
    for (const mode of ['interactive', 'paged'] as const) {
      const res = await compareCfCase(caseOf(), mode, { rootDir: REPO_ROOT })
      expect(res.outcome, `${mode}: ${res.reason}`).toBe('pass')
    }
  }, HEAVY)
})

describe('a committed input survives regeneration (codex pre-capture P2)', () => {
  it('keepCommittedArchive keeps the committed bytes when the inflated parts match', async () => {
    const { keepCommittedArchive, generateCaseArchive } = await import('../scripts/cf-refs/generate')
    const committed = new Uint8Array(readFileSync(resolve(REPO_ROOT, caseOf().input)))
    expect(keepCommittedArchive(committed, await generateCaseArchive(ID))).toBe(committed)
  }, HEAVY)
})
