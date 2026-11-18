/**
 * Shared harness for the Stage 5 formula-rule tests (Task 35): parse a `<conditionalFormatting>` block under an exact
 * evidence set, build a resident sheet from literal cells and prepare it through the real `createCfService` route with an
 * evaluation context. Nothing here decides an expectation: each test states its own.
 */
import { getChildren, parseXml, type XmlNode } from '../../src/core/xml'
import { withExactCfEvidence, injectCfFormulaSelections } from '../../src/xlsx/cf/evidence'
import { parseCfSheet, type CfStylesSource } from '../../src/xlsx/cf/parse'
import { cfEvaluationContextFor, type CfDocumentEvaluation } from '../../src/xlsx/cf/formula'
import { createCfService } from '../../src/xlsx/cf/service'
import { residentSource } from '../../src/xlsx/cf/render'
import { CF_LIMITS, createCfMemory } from '../../src/xlsx/cf/limits'
import { resolveCfCellStyle } from '../../src/xlsx/cf/dxf'
import type { CfCellState, CfOutcome, CfSheetRules, CfStamp } from '../../src/xlsx/cf/model'
import type { XlsxSheet } from '../../src/xlsx/types'

export const NS = 'http://schemas.openxmlformats.org/spreadsheetml/2006/main'
/** dxfId 0 red fill, 1 green fill, 2 strikethrough only (the Stage 5 oracle's dxfs). */
export const RED = 'FFFFC7CE', GREEN = 'FFC6EFCE'
const DXFS = `<dxfs count="3"><dxf><fill><patternFill patternType="solid"><bgColor rgb="${RED}"/></patternFill></fill></dxf>`
  + `<dxf><fill><patternFill patternType="solid"><bgColor rgb="${GREEN}"/></patternFill></fill></dxf><dxf><font><strike/></font></dxf></dxfs>`
export const stylesFor = (): CfStylesSource => {
  const root = parseXml(`<styleSheet xmlns="${NS}">${DXFS}</styleSheet>`)
  return { colors: {}, dxfNodes: getChildren(getChildren(root, 'dxfs')[0] as XmlNode, 'dxf') }
}
export const UTC: CfDocumentEvaluation = Object.freeze({ calcMode: 'auto', valuePolicy: 'engine', dateSystem: '1900', timeZone: 'UTC', unicode: { version: 1 as const, source: 'workbook-default' as const } })

/** Every capability a same-sheet DXF formula rule of these tests needs, with production's readings. */
export const FORMULA_CAPS: readonly string[] = [
  'formula:dialect-base', 'formula:dialect-x14', 'formula:anchor-single', 'formula:anchor-mixed-abs',
  'formula:truth-number', 'formula:truth-boolean', 'formula:truth-blank', 'formula:truth-error', 'formula:truth-text',
  'formula:fn-logical', 'formula:fn-information', 'formula:fn-text', 'formula:fn-math', 'formula:fn-aggregate',
  'formula:fn-conditional-aggregate', 'formula:fn-date', 'formula:fn-row-column',
  'formula:predicate-divergent', 'formula:cellis-operand-reference', 'formula:cellis-operand-relative', 'formula:cellis-operand-expression',
]
export const SELECTIONS: Readonly<Record<string, string>> = { 'formula:truth-text': 'text-no-match' }

export type Cell = number | string | boolean | { err: string }
const refRC = (ref: string): { row: number; col: number } => {
  const m = /^([A-Z]+)(\d+)$/.exec(ref)!
  let col = 0
  for (const ch of m[1]!) col = col * 26 + ch.charCodeAt(0) - 64
  return { row: Number(m[2]) - 1, col: col - 1 }
}
export const rc = refRC

export function parseRules(cfXml: string, caps: Iterable<string>): CfSheetRules {
  return withExactCfEvidence(caps, () => parseCfSheet(parseXml(`<worksheet xmlns="${NS}">${cfXml}</worksheet>`), stylesFor(), createCfMemory()).value)
}

export function sheetOf(rules: CfSheetRules, cells: Readonly<Record<string, Cell>>): XlsxSheet {
  const rows = new Map<number, unknown[]>()
  for (const [ref, v] of Object.entries(cells)) {
    const { row, col } = refRC(ref)
    const error = typeof v === 'object'
    rows.set(row, [...(rows.get(row) ?? []), { ref, row, col, styleIndex: 0, style: { numFmtId: 0 }, value: error ? v.err : v, ...(error ? { valueIsError: true } : {}), hasCachedValue: true }])
  }
  return {
    name: 'Sheet1', sheetId: 'sheet1', cols: [], merges: [], mergeRanges: [], conditionalFormatting: rules,
    rows: [...rows].sort((a, b) => a[0] - b[0]).map(([index, list]) => ({ index, cells: (list as Array<{ col: number }>).sort((a, b) => a.col - b.col) })),
  } as unknown as XlsxSheet
}

/** Prepare `refs` through the real service under exactly `caps` (service gate and compose) and `selections`. */
export async function prepare(sheet: XlsxSheet, refs: readonly string[], caps: Iterable<string>, options: { selections?: Readonly<Record<string, string>>; evaluation?: CfDocumentEvaluation | null } = {}): Promise<Map<string, CfOutcome<CfCellState>>> {
  const set = new Set(caps)
  const memory = createCfMemory()
  const evaluation = options.evaluation === null ? undefined : cfEvaluationContextFor(options.evaluation ?? UTC, () => Date.UTC(2026, 10, 16))
  const service = createCfService(sheet.conditionalFormatting!, memory, { evidence: { allows: cap => set.has(cap) }, ...(evaluation ? { evaluation } : {}), sheetName: sheet.name })
  const stamp: CfStamp = { owner: 't', sheetId: sheet.sheetId!, workbookValueGen: 0, cfGen: 1, viewGen: 0, structureGen: 0, clockGen: 0 }
  const out = new Map<string, CfOutcome<CfCellState>>()
  injectCfFormulaSelections(options.selections ?? SELECTIONS)
  try {
    const rects = refs.map(r => { const { row, col } = refRC(r); return { firstRow: row, lastRow: row, firstCol: col, lastCol: col } })
    const lease = await service.prepareCf({ stamp, sourceRects: rects, signal: new AbortController().signal, grantBytes: CF_LIMITS.scratchBytes }, residentSource(sheet, stamp, memory, 'engine'))
    for (const r of refs) { const { row, col } = refRC(r); out.set(r, lease.value.at(row, col)) }
    lease.release()
  } finally {
    injectCfFormulaSelections(null)
    service.dispose()
  }
  return out
}

/** What a cell shows: its CF fill (or `none`), strikethrough, and whether any write is withheld. */
export function shown(outcome: CfOutcome<CfCellState> | undefined): { fill: string; strike: boolean; withheld: boolean } {
  if (!outcome || outcome.kind !== 'ready') return { fill: 'none', strike: false, withheld: outcome?.kind === 'unresolved' }
  const style = resolveCfCellStyle(undefined, outcome.value)
  return { fill: style.fillColor ?? 'none', strike: style.strike === true, withheld: (outcome.value.unresolvedWrites?.size ?? 0) > 0 }
}
