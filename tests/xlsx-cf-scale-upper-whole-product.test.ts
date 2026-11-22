// @vitest-environment node
/**
 * The num-segment follow-up, phase 2 (owner ruling A after capture stage2-scale-num-segment): production withholds a
 * 3-colour colour-scale cell strictly inside the UPPER segment (between the second and third thresholds after the
 * midpoint-above-max re-sort), for ANY bound kind, whenever the step product is a whole number: for some RGB channel
 * whose endpoints differ by d != 0, `fraction * d` (production's `safeFraction` double times d) is an integer
 * (`cf-unmeasured-upper-segment-whole-product`). The literal-num withhold (`numSegmentUnmeasured`, f·2^20) is unchanged and
 * takes precedence. Nothing about the pre-registered readings is selectable: no reading was matched.
 */
import { readFileSync } from 'node:fs'
import { resolve } from 'node:path'
import { describe, expect, it } from 'vitest'
import * as colorScale from '../src/xlsx/cf/colorScale'
import { interpolateScaleColor, scaleColorAt, upperSegmentWholeProduct } from '../src/xlsx/cf/colorScale'
import { OfficePackage } from '../src/core/zip'
import { parseXlsx } from '../src/xlsx/parse'
import { cfDocumentEvaluationOf, prepareCfPaint } from '../src/xlsx/cf/render'
import { cfRuntimeEvidence } from '../src/xlsx/cf/predicates'
import type { CfScalePoint } from '../src/xlsx/cf/model'
import { s2NumSegmentColour, s2NumSegmentRules, s2NumSegmentValueAt, type CfManifest } from '../scripts/cf-refs/cases'
import { xlsxBytes } from './helpers/xlsx-package'
import { collectSegmentCells } from '../scripts/cf-refs/num-segment-candidates'
import { STAGE5_UPPER_WHOLE_PRODUCT_REFUSALS } from '../scripts/cf-refs/stage5-withhold'
import { VERIFIED_CAPABILITIES } from '../src/xlsx/cf/evidence'
import { cfvoNonNumericTreatmentOf } from '../src/xlsx/cf/formulaReadings'

const REPO_ROOT = resolve(__dirname, '..')
const RED = 0xfff8696b, YELLOW = 0xffffeb84, GREEN = 0xff63be7b
const P = 0xff14f028, Q = 0xffe61efa
const num = (value: number, argb: number): CfScalePoint => ({ bound: { type: 'num', value }, argb })
const all = { hasEvidence: () => true }
const UPPER = { kind: 'unresolved', reason: 'cf-unmeasured-upper-segment-whole-product' }
const LITERAL = { kind: 'unresolved', reason: 'cf-unmeasured-num-segment-fraction' }
const manifest = JSON.parse(readFileSync(resolve(REPO_ROOT, 'tests/fixtures/xlsx-cf/manifest.json'), 'utf8')) as CfManifest
const sidecar = JSON.parse(readFileSync(resolve(REPO_ROOT, 'tests/fixtures/xlsx-cf/refs/stage2-scale-num-segment.json'), 'utf8')) as { observedCells: Array<{ ref: string; fillRgb: string }> }
const excel = new Map(sidecar.observedCells.map(o => [o.ref, o.fillRgb]))

describe('the shape, exactly', () => {
  it('upper segment only, strictly inside, some channel with d != 0 and Number.isInteger(f·d)', () => {
    expect(upperSegmentWholeProduct(YELLOW, GREEN, 40, 30, 60, 1 / 3)).toBe(true) // T4: 156/3 = 52
    expect(upperSegmentWholeProduct(YELLOW, GREEN, 50, 30, 60, 2 / 3)).toBe(true) // V3
    expect(upperSegmentWholeProduct(YELLOW, GREEN, 70, 60, 100, 0.25)).toBe(true) // 156/4 = 39 (one channel suffices)
    expect(upperSegmentWholeProduct(YELLOW, GREEN, 30, 30, 60, 0)).toBe(false) // an endpoint
    expect(upperSegmentWholeProduct(YELLOW, GREEN, 60, 30, 60, 1)).toBe(false)
    expect(upperSegmentWholeProduct(YELLOW, GREEN, 41, 30, 60, 11 / 30)).toBe(false) // 57.2, 16.5, 3.3
    expect(upperSegmentWholeProduct(YELLOW, YELLOW, 40, 30, 60, 1 / 3)).toBe(false) // d = 0 on every channel
    // E4 of stage2-scale-thresholds: 15.999999999999996 is not an integer (Excel drew 16, production paints 16).
    expect(upperSegmentWholeProduct(YELLOW, GREEN, 3, 2.2, 10, (3 - 2.2) / (10 - 2.2))).toBe(false)
  })

  it('production: the upper-segment shape is withheld for every bound kind; the lower segment, 2-colour scales and endpoints paint as before', () => {
    const lit = [num(0, RED), num(30, YELLOW), num(60, GREEN)]
    expect(scaleColorAt(40, lit, [0, 30, 60], all)).toEqual(LITERAL) // the literal withhold keeps its reason
    expect(scaleColorAt(45, lit, [0, 30, 60], all)).toEqual(UPPER) // dyadic, so only the new withhold applies
    expect(scaleColorAt(15, lit, [0, 30, 60], all)).toEqual({ kind: 'ready', value: interpolateScaleColor(RED, YELLOW, 0.5) }) // lower
    expect(scaleColorAt(30, lit, [0, 30, 60], all)).toEqual({ kind: 'ready', value: YELLOW })
    expect(scaleColorAt(60, lit, [0, 30, 60], all)).toEqual({ kind: 'ready', value: GREEN })
    const stat: CfScalePoint[] = [{ bound: { type: 'min' }, argb: RED }, { bound: { type: 'percentile', value: 50 }, argb: YELLOW }, { bound: { type: 'max' }, argb: GREEN }]
    expect(scaleColorAt(40, stat, [0, 30, 60], all)).toEqual(UPPER)
    expect(scaleColorAt(41, stat, [0, 30, 60], all)).toEqual({ kind: 'ready', value: interpolateScaleColor(YELLOW, GREEN, 11 / 30) })
    expect(scaleColorAt(10, stat, [0, 30, 60], all)).toEqual({ kind: 'ready', value: interpolateScaleColor(RED, YELLOW, 1 / 3) })
    const mixed: CfScalePoint[] = [num(0, RED), num(30, YELLOW), { bound: { type: 'max' }, argb: GREEN }]
    expect(scaleColorAt(40, mixed, [0, 30, 60], all)).toEqual(UPPER)
    const formula: CfScalePoint[] = [num(0, RED), { bound: { type: 'num', value: 30, resolved: 'formula' }, argb: YELLOW }, num(60, GREEN)]
    expect(scaleColorAt(40, formula, [0, 30, 60], all)).toEqual(UPPER)
    // The midpoint above the maximum re-sorts first: num 0 / 60 / 30 has the upper segment 30..60.
    expect(scaleColorAt(45, [num(0, RED), num(60, YELLOW), num(30, GREEN)], [0, 60, 30], all)).toEqual(UPPER)
    expect(scaleColorAt(15, [num(0, RED), num(60, YELLOW), num(30, GREEN)], [0, 60, 30], all)).toEqual({ kind: 'ready', value: interpolateScaleColor(RED, YELLOW, 0.5) })
    expect(scaleColorAt(20, [num(0, RED), num(40, GREEN)], [0, 40], all)).toEqual({ kind: 'ready', value: interpolateScaleColor(RED, GREEN, 0.5) })
  })

  it('nothing from the pre-registered readings is selectable: no reading, consensus or registration seam is exported', () => {
    for (const name of ['injectCfNumSegmentReadings', 'cfNumSegmentReading', 'numSegmentConsensus', 'numSegmentConsensusColor', 'numSegmentReadingColor', 'CF_NUM_SEGMENT_READINGS', 'CF_NUM_SEGMENT_PRODUCTION_READINGS']) expect(name in colorScale, name).toBe(false)
  })
})

describe('the capture (stage2-scale-num-segment): every moved cell is in the shape', () => {
  const rules = s2NumSegmentRules()
  const designated = rules.flatMap(r => (r.designatedRows ?? r.rows).map(row => ({ r, row, ref: `${r.col}${row}` })))
  const truncation = (r: (typeof rules)[number], row: number): string => {
    const v = s2NumSegmentValueAt(r, row), t = r.thresholds, seg = t.length === 2 || v <= t[1]! ? 0 : 1
    return s2NumSegmentColour('truncate', r.colors[seg]!, r.colors[seg + 1]!, v, t[seg]!, t[seg + 1]!)
  }

  it('Excel departs from truncation at exactly 19 cells, all one unit smaller per step, all in the shape (plus T4 and V3 of e448d831)', () => {
    const moved = designated.filter(d => excel.get(d.ref) !== truncation(d.r, d.row))
    expect(moved.map(d => d.ref)).toEqual(['B17', 'B18', 'B21', 'B24', 'B28', 'C17', 'C18', 'C21', 'C24', 'C28', 'D7', 'D8', 'D12', 'D13', 'E17', 'E18', 'E21', 'E24', 'E28'])
    for (const d of moved) {
      const v = s2NumSegmentValueAt(d.r, d.row), t = d.r.thresholds
      const [lo, hi] = [t[1]!, t[2]!]
      const f = (v - lo) / (hi - lo)
      expect(v > t[1]!, `${d.ref}: upper segment`).toBe(true)
      expect(upperSegmentWholeProduct(Number.parseInt(d.r.colors[1]!, 16) >>> 0, Number.parseInt(d.r.colors[2]!, 16) >>> 0, v, lo, hi, f), d.ref).toBe(true)
    }
    expect(upperSegmentWholeProduct(YELLOW, GREEN, 40, 30, 60, (40 - 30) / 30)).toBe(true)
    expect(upperSegmentWholeProduct(YELLOW, GREEN, 50, 30, 60, (50 - 30) / 30)).toBe(true)
  })

  it('default production on the saved workbook withholds every moved cell (E17, E18, E21, E24, E28 included) and paints every other painted cell at Excel\'s fill', async () => {
    const c = manifest.cases.find(x => x.id === 'stage2-scale-num-segment')!
    const doc = await parseXlsx(await OfficePackage.load(new Uint8Array(readFileSync(resolve(REPO_ROOT, c.saved)))))
    const paint = await prepareCfPaint(doc.sheets[0]!, { evidence: cfRuntimeEvidence, evaluation: cfDocumentEvaluationOf(doc) })
    const fillAt = (ref: string): string | undefined => { const m = /^([A-Z]+)(\d+)$/.exec(ref)!; return paint!.layer.styleAt(Number(m[2]) - 1, [...m[1]!].reduce((n, ch) => n * 26 + ch.charCodeAt(0) - 64, 0) - 1, undefined)?.fillColor?.toUpperCase() }
    const withheld: string[] = []
    for (const e of c.expected.cells) {
      const got = fillAt(e.ref)
      if (!e.matched) { expect(got, e.ref).toBeUndefined(); continue }
      if (got === undefined) withheld.push(e.ref)
      else expect(got, e.ref).toBe(excel.get(e.ref))
    }
    paint!.dispose()
    for (const ref of ['B17', 'B18', 'B21', 'B24', 'B28', 'C17', 'C18', 'C21', 'C24', 'C28', 'D7', 'D8', 'D12', 'D13', 'E17', 'E18', 'E21', 'E24', 'E28']) expect(withheld, ref).toContain(ref)
    // 96 literal-num cells (the num-segment withhold) and 13 by the new rule: B26, C26 (literal halves) and E17..E26, E28.
    expect(withheld.length).toBe(109)
  }, 300_000)
})

describe('the reviewer\'s overflow sheet (pre-capture fix round 2): default production is unchanged', () => {
  it('num -9e307 / 9e307 / 9.5e307, P/Q/P: B2 = 0 paints the safe half colour (lower segment), B3 and B4 their configured colours', async () => {
    const bytes = xlsxBytes({
      sheetData: '<row r="2"><c r="B2"><v>0</v></c></row><row r="3"><c r="B3"><v>9e307</v></c></row><row r="4"><c r="B4"><v>9.5e307</v></c></row>',
      afterSheetData: '<conditionalFormatting sqref="B2:B4"><cfRule type="colorScale" priority="1"><colorScale><cfvo type="num" val="-9e307"/><cfvo type="num" val="9e307"/><cfvo type="num" val="9.5e307"/><color rgb="FF14F028"/><color rgb="FFE61EFA"/><color rgb="FF14F028"/></colorScale></cfRule></conditionalFormatting>',
    })
    const doc = await parseXlsx(await OfficePackage.load(bytes))
    const paint = await prepareCfPaint(doc.sheets[0]!, { evidence: cfRuntimeEvidence })
    expect(['B2', 'B3', 'B4'].map((_, i) => paint?.layer.styleAt(i + 1, 1, undefined)?.fillColor?.toUpperCase())).toEqual(['FF7D8791', 'FFE61EFA', 'FF14F028'])
    paint?.dispose()
    expect(P).not.toBe(Q)
  })
})

describe('the reach: every newly withheld cell of every measured case is the shape, and Stage 5\'s list is derived', () => {
  it('per case (through production\'s own segment evaluations): the cells in the shape are exactly the recorded reach', async () => {
    const { cells } = await collectSegmentCells(manifest, REPO_ROOT, { include: 'all' })
    // The literal-num withhold takes precedence (num / num bounds, f·2^20 not an integer): those cells keep its reason.
    const literalWithheld = (s: (typeof cells)[number]['segment']): boolean => s.loKind === 'num' && s.hiKind === 'num' && !Number.isInteger(s.fraction * 1048576)
    const inShape = cells.filter(c => c.segment.colours === 3 && c.segment.segment === 1 && !literalWithheld(c.segment) && upperSegmentWholeProduct(c.segment.c0, c.segment.c1, c.segment.value, c.segment.lo, c.segment.hi, c.segment.fraction))
    const byCase: Record<string, string[]> = {}
    for (const c of inShape) (byCase[c.caseId] ??= []).push(c.ref)
    expect(Object.fromEntries(Object.entries(byCase).map(([k, v]) => [k, [...v].sort()]))).toEqual({
      'stage2-scale-interpolation': ['E10', 'E8', 'E9'],
      'stage2-scale-thresholds': ['E5'],
      'stage2-scale-scope-disjoint': ['U4'],
      'stage2-scale-num-segment': ['B26', 'C26', 'E17', 'E18', 'E19', 'E20', 'E21', 'E22', 'E23', 'E24', 'E25', 'E26', 'E28'].sort(),
      'stage5-cfvo-nonnumeric-v2': ['E4', 'E5', 'S4', 'S5'],
    })
    // Stage 5: the enumerated justification is exactly the derived set under the published readings.
    const listed = STAGE5_UPPER_WHOLE_PRODUCT_REFUSALS.filter(e => e.midpointKind === undefined || cfvoNonNumericTreatmentOf(e.midpointKind, id => VERIFIED_CAPABILITIES.has(id)) === 'zero').flatMap(e => e.refs).sort()
    expect(listed).toEqual(byCase['stage5-cfvo-nonnumeric-v2']!.sort())
  }, 300_000)
})

