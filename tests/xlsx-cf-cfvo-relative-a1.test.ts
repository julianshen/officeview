// @vitest-environment node
/**
 * The single-range A1-anchored relative cfvo bound (the cfvo-relative follow-up, brief 2026-11-21, Phase 1).
 *
 * stage5-cfvo-scope-v2 (5dc2fe9e) showed that Excel evaluates a relative cfvo bound as if its formula sat in A1, on one range
 * as on several. Production implements that as the reading `anchor-a1` of `formula:cfvo-relative`: `cfvoBoundAnchors` puts
 * a single range's bound at A1. Written before its registration (stage5-cfvo-relative-a1 confirmed it, f9d5be94), these
 * tests stand in for the registry: `cfFormulaReading` is wrapped to return a chosen reading for `formula:cfvo-relative`.
 * Every expectation is a literal twin rule production itself resolves; nothing here records or predicts an Excel capture.
 */
import { describe, expect, it, vi } from 'vitest'
import { parseXml } from '../src/core/xml'
import { cfEvaluationContextFor } from '../src/xlsx/cf/formula'
import { residentSource } from '../src/xlsx/cf/render'
import { cfvoBoundAnchors, createCfService } from '../src/xlsx/cf/service'
import { CF_LIMITS, createCfMemory } from '../src/xlsx/cf/limits'
import { VERIFIED_CAPABILITIES, withExactCfEvidence } from '../src/xlsx/cf/evidence'
import { CF_FORMULA_PRODUCTION_READINGS } from '../src/xlsx/cf/formulaReadings'
import { MEASURED_CF_OFFICE_THEME, parseCfSheet, type CfStylesSource } from '../src/xlsx/cf/parse'
import { cfRulesVersion } from '../src/xlsx/view/pagedSource'
import type { CfCellState, CfOutcome, CfSheetRules, CfStamp } from '../src/xlsx/cf/model'
import type { XlsxSheet } from '../src/xlsx/types'
import { UTC, rc, sheetOf } from './helpers/cf-formula-rules'

/** The reading the wrapped registry returns for `formula:cfvo-relative` (undefined: production's own registry). */
let relativeReading: string | undefined
vi.mock('../src/xlsx/cf/formulaReadings', async importOriginal => {
  const actual = await importOriginal<typeof import('../src/xlsx/cf/formulaReadings')>()
  return {
    ...actual,
    cfFormulaReading: (capability: string, check: (c: string) => boolean) =>
      capability === 'formula:cfvo-relative' && relativeReading !== undefined ? (check(capability) ? relativeReading : undefined) : actual.cfFormulaReading(capability, check),
  }
})

const NS = 'http://schemas.openxmlformats.org/spreadsheetml/2006/main'
const styles: CfStylesSource = { colors: { theme: MEASURED_CF_OFFICE_THEME }, dxfNodes: [] }
const CAPS: readonly string[] = [
  ...VERIFIED_CAPABILITIES, 'formula:dialect-base', 'formula:anchor-single', 'formula:fn-aggregate', 'formula:cfvo-reference', 'formula:cfvo-expression',
  'scale:thresholds-formula', 'bar:bounds-formula', 'icon:thresholds-formula', 'bar:bounds-num', 'formula:cfvo-relative',
]
const COLORS = '<color rgb="FFF8696B"/><color rgb="FF63BE7B"/>'
const scale = (sqref: string, priority: number, cfvos: string): string =>
  `<conditionalFormatting sqref="${sqref}"><cfRule type="colorScale" priority="${priority}"><colorScale>${cfvos}${COLORS}</colorScale></cfRule></conditionalFormatting>`
const icon = (sqref: string, priority: number, cfvos: string): string =>
  `<conditionalFormatting sqref="${sqref}"><cfRule type="iconSet" priority="${priority}"><iconSet iconSet="3Arrows">${cfvos}</iconSet></cfRule></conditionalFormatting>`
const parseRules = (inner: string, caps: Iterable<string>): CfSheetRules =>
  withExactCfEvidence(caps, () => parseCfSheet(parseXml(`<worksheet xmlns="${NS}">${inner}</worksheet>`), styles, createCfMemory(), { sheetRightToLeft: false }).value)
const col = (letter: string, values: readonly number[], from = 2): Record<string, number> => Object.fromEntries(values.map((v, i) => [`${letter}${from + i}`, v]))
const rows = (letter: string, from: number, to: number): string[] => Array.from({ length: to - from + 1 }, (_, i) => `${letter}${from + i}`)

async function prepare(sheet: XlsxSheet, refs: readonly string[], caps: Iterable<string>): Promise<Map<string, CfOutcome<CfCellState>>> {
  const set = new Set(caps)
  const memory = createCfMemory()
  const service = createCfService(sheet.conditionalFormatting!, memory, { evidence: { allows: c => set.has(c) }, evaluation: cfEvaluationContextFor(UTC, () => Date.UTC(2026, 10, 21)), sheetName: sheet.name })
  const stamp: CfStamp = { owner: 't', sheetId: sheet.sheetId ?? 'sheet1', workbookValueGen: 0, cfGen: cfRulesVersion(sheet.conditionalFormatting!), viewGen: 0, structureGen: 0, clockGen: 0 }
  const rects = refs.map(r => { const { row, col: c } = rc(r); return { firstRow: row, lastRow: row, firstCol: c, lastCol: c } })
  const lease = await service.prepareCf({ stamp, sourceRects: rects, signal: new AbortController().signal, grantBytes: CF_LIMITS.scratchBytes }, residentSource(sheet, stamp, memory, 'engine'))
  const out = new Map<string, CfOutcome<CfCellState>>()
  for (const r of refs) { const { row, col: c } = rc(r); out.set(r, lease.value.at(row, c)) }
  lease.release(); service.dispose()
  return out
}
function visualOf(o: CfOutcome<CfCellState> | undefined): string {
  if (!o) return 'missing'
  if (o.kind !== 'ready') return `unresolved:${o.kind === 'unresolved' ? o.reason : o.kind}`
  if ((o.value.unresolvedWrites?.size ?? 0) > 0) return 'withheld'
  const v = o.value.visual
  if (!v) return 'none'
  const strip = (x: object | undefined): string | undefined => { if (!x) return undefined; const { ruleId: _r, ...rest } = x as Record<string, unknown>; void _r; return JSON.stringify(rest) }
  return JSON.stringify({ scale: strip(v.colorScale), bar: strip(v.dataBar), icon: strip(v.iconSet) })
}
const twinsAgree = (out: Map<string, CfOutcome<CfCellState>>, pairs: ReadonlyArray<readonly [string, string]>): void => {
  for (const [a, b] of pairs) {
    expect(visualOf(out.get(a)), `${a} vs twin ${b}`).toBe(visualOf(out.get(b)))
    expect(visualOf(out.get(a)), a).toMatch(/^\{/)
  }
}
const WITHHELD = /^(unresolved:.*|withheld)$/

// A1:A6 (max 15) is the window from A1; B2:B7 (max 35) the window from the rule's own top-left.
const A = col('A', [-35, 15, -30, 5, -20, 10], 1)
const DATA = [-40, -25, -5, 5, 20, 35]
const twinRows = (twin: string): Array<readonly [string, string]> => rows('B', 2, 7).map(r => [r, `${twin}${r.slice(1)}`] as const)

describe('cfvoBoundAnchors on one range', () => {
  it('anchor-a1 puts the single scope at A1; anchor-top-left (and no reading) at the range\'s top-left', () => {
    const one = [{ firstRow: 1, lastRow: 6, firstCol: 1, lastCol: 1 }]
    expect(cfvoBoundAnchors(one, 'anchor-a1')).toEqual([{ row: 0, col: 0 }])
    expect(cfvoBoundAnchors(one, 'anchor-top-left')).toEqual([{ row: 1, col: 1 }])
    expect(cfvoBoundAnchors(one, undefined)).toEqual([{ row: 1, col: 1 }])
  })
})

describe('formula:cfvo-relative read as anchor-a1 (registered since stage5-cfvo-relative-a1, f9d5be94)', () => {
  it('the registry holds anchor-a1 and the capability is published (anchor-top-left stays reachable only through this test\'s wrapper)', () => {
    expect(CF_FORMULA_PRODUCTION_READINGS['formula:cfvo-relative']).toBe('anchor-a1')
    expect(VERIFIED_CAPABILITIES.has('formula:cfvo-relative')).toBe(true)
  })

  const xml = scale('B2:B7', 1, '<cfvo type="min"/><cfvo type="formula" val="MAX(B2:B7)"/>')
    + scale('D2:D7', 2, '<cfvo type="num" val="-40"/><cfvo type="num" val="15"/>')
    + scale('E2:E7', 3, '<cfvo type="num" val="-40"/><cfvo type="num" val="35"/>')
  const cells = { ...A, ...col('B', DATA), ...col('D', DATA), ...col('E', DATA) }
  const refs = [...rows('B', 2, 7), ...rows('D', 2, 7), ...rows('E', 2, 7)]

  it('a single-range scale reads its relative bound from A1 under anchor-a1 (=MAX(B2:B7) is MAX(A1:A6) = 15)', async () => {
    relativeReading = 'anchor-a1'
    try {
      const out = await prepare(sheetOf(parseRules(xml, CAPS), cells), refs, CAPS)
      twinsAgree(out, twinRows('D'))
    } finally { relativeReading = undefined }
  })

  it('under anchor-top-left (a historical hypothesis, reachable only through this test\'s wrapper) the same rule reads its own top-left (MAX(B2:B7) = 35)', async () => {
    relativeReading = 'anchor-top-left'
    try {
      const out = await prepare(sheetOf(parseRules(xml, CAPS), cells), refs, CAPS)
      twinsAgree(out, twinRows('E'))
    } finally { relativeReading = undefined }
  })

  it('an icon set\'s relative threshold reads from A1 too; an absolute one is the same value at either anchor', async () => {
    const ixml = icon('B2:B7', 1, '<cfvo type="percent" val="0"/><cfvo type="num" val="-20"/><cfvo type="formula" val="MAX(B2:B7)"/>')
      + icon('D2:D7', 2, '<cfvo type="percent" val="0"/><cfvo type="num" val="-20"/><cfvo type="num" val="15"/>')
      + icon('G2:G7', 3, '<cfvo type="percent" val="0"/><cfvo type="num" val="-20"/><cfvo type="formula" val="$H$1"/>')
      + icon('I2:I7', 4, '<cfvo type="percent" val="0"/><cfvo type="num" val="-20"/><cfvo type="num" val="0"/>')
    const icells = { ...A, H1: 0, ...col('B', DATA), ...col('D', DATA), ...col('G', DATA), ...col('I', DATA) }
    relativeReading = 'anchor-a1'
    try {
      const out = await prepare(sheetOf(parseRules(ixml, CAPS), icells), [...refs.slice(0, 12), ...rows('G', 2, 7), ...rows('I', 2, 7)], CAPS)
      twinsAgree(out, twinRows('D'))
      twinsAgree(out, rows('G', 2, 7).map(r => [r, `I${r.slice(1)}`] as const))
    } finally { relativeReading = undefined }
  })

  it('a window that would leave the sheet from A1 follows the edge reading: withheld without one', async () => {
    // =MAX(A2:A7) written on B2, read from A1, names the column left of A.
    const exml = scale('B2:B7', 1, '<cfvo type="min"/><cfvo type="formula" val="MAX(A2:A7)"/>')
    const caps = CAPS.filter(c => c !== 'formula:relative-wrap')
    relativeReading = 'anchor-a1'
    try {
      const out = await prepare(sheetOf(parseRules(exml, caps), { ...A, ...col('B', DATA) }), rows('B', 2, 7), caps)
      for (const r of rows('B', 2, 7)) expect(visualOf(out.get(r)), r).toMatch(WITHHELD)
    } finally { relativeReading = undefined }
  })
})
