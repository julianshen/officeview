/**
 * Data-bar fixtures for the Task 22 painting tests. Nothing Stage 3 is published (Task 22 addendum §1), so production
 * withholds every bar; these helpers paint bars under INJECTED evidence and an injected layout only. The numbers they
 * produce are the measured model run under stubs, never new Excel facts.
 */
import { injectCfBarLayout, withInjectedCfEvidence } from '../../src/xlsx/cf/evidence'
import type { CfDefinition, CfRule } from '../../src/xlsx/cf/model'
import type { XlsxCell, XlsxSheet } from '../../src/xlsx/types'

/** Every evaluation gate a solid min/max bar can need, plus the Task 22 paint gates (stubs only). */
export const BAR_EVAL_EVIDENCE = [
  'bar:bounds-min-max', 'bar:bounds-num', 'bar:lengths-explicit-0-100', 'bar:render-solid', 'bar:render-border',
  'bar:axis-automatic', 'bar:axis-middle', 'bar:axis-none', 'bar:direction-ltr', 'bar:direction-rtl', 'bar:direction-context-ltr-sheet',
  'bar:negative-fill-color', 'bar:negative-fill-same-as-positive', 'bar:negative-border-color', 'bar:negative-border-same-as-positive',
  'bar:participation-types', 'bar:scope-whole-column',
]
/** Parse gates of a workbook bar in Excel's paired x14 form (stubs only). */
export const BAR_PARSE_EVIDENCE = ['bar:settings-precondition-probe', 'bar:settings-schema-defaults', 'bar:settings-paired-precedence', 'bar:pairing-paired']
export const BAR_PAINT_EVIDENCE = ['bar:geometry-insets', 'bar:geometry-snapping', 'bar:geometry-narrow-clip']
export const BAR_SHOWVALUE_EVIDENCE = 'bar:showvalue-canvas-suppression'

/** The layout Task 23's export would publish (S3-B), as the harness currently selects it. */
export const BAR_LAYOUT = {
  allPositive: { origin: 'span', lengthBasis: 'cell', settledBy: ['bar:lengths-explicit-0-100'] },
  mixedSign: { origin: 'axis', lengthBasis: 'side', settledBy: ['bar:axis-automatic', 'bar:axis-middle'] },
}

/** Runs `fn` with injected bar evidence (joined to production evidence) and the injected layout; both are cleared after. */
export async function withBars<T>(fn: () => Promise<T> | T, options: { paint?: boolean; extra?: readonly string[]; omit?: readonly string[] } = {}): Promise<T> {
  injectCfBarLayout(BAR_LAYOUT)
  const caps = [...BAR_EVAL_EVIDENCE, ...BAR_PARSE_EVIDENCE, ...(options.paint === false ? [] : BAR_PAINT_EVIDENCE), ...(options.extra ?? [])].filter(c => !options.omit?.includes(c))
  try { return await withInjectedCfEvidence(caps, async () => fn()) } finally { injectCfBarLayout(null) }
}

export const BLUE = 0xff0000ff, RED = 0xffff0000, BLACK = 0xff000000, GREEN = 0xff00b050

type BarDef = Extract<CfDefinition, { kind: 'dataBar' }>
export function barDef(overrides: Partial<BarDef> = {}): BarDef {
  return {
    kind: 'dataBar', dialect: 'base', formulas: [],
    bounds: [{ type: 'min' }, { type: 'max' }],
    fillArgb: BLUE, negativeFillArgb: RED, axisArgb: BLACK,
    minLength: 0, maxLength: 100, gradient: false, border: false,
    negativeBarColorSameAsPositive: false, negativeBarBorderColorSameAsPositive: true,
    axisPosition: 'automatic', direction: 'leftToRight', showValue: true, sheetRightToLeft: false,
    ...overrides,
  }
}

export function barRule(range: { firstRow: number; lastRow: number; firstCol: number; lastCol: number }, overrides: Partial<BarDef> = {}, id = 1): CfRule {
  return { id, priority: id, documentOrder: id - 1, ranges: [range], stopIfTrue: false, writes: new Set(['dataBar']), definition: barDef(overrides) }
}

/**
 * A one-column bar sheet: `values` down column `col` from row 0, the column `widthChars` wide (200/7 characters is a
 * 200 px column), one bar rule over exactly those cells.
 */
export function barSheet(values: readonly number[], options: { col?: number; widthChars?: number; def?: Partial<BarDef>; name?: string; firstRow?: number; rows?: number } = {}): XlsxSheet {
  const col = options.col ?? 0, first = options.firstRow ?? 0
  const cells = values.map((value, i): XlsxCell => ({ ref: `${String.fromCharCode(65 + col)}${first + i + 1}`, row: first + i, col, value, styleIndex: 0, hasCachedValue: true }))
  const rows = Math.max(options.rows ?? 0, first + values.length)
  return {
    name: options.name ?? 'Bars', sheetId: options.name ?? 'bars',
    rows: Array.from({ length: rows }, (_, r) => ({ index: r, cells: cells.filter(c => c.row === r) })),
    cols: [{ min: col, max: col, widthChars: options.widthChars ?? 200 / 7 }],
    merges: [], mergeRanges: [],
    conditionalFormatting: { diagnostics: [], rules: [barRule({ firstRow: first, lastRow: first + values.length - 1, firstCol: col, lastCol: col }, options.def)] },
  } as XlsxSheet
}
