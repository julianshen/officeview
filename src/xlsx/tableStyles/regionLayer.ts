/** Lazy resident worksheet table-style composition. */
import { parseStylesExtras } from './parseStylesExtras'
import { builtinTableStyle } from './builtins'
import { orderRoles, resolveRoleStyle } from './resolve'
import type { DxfStyle, RoleRef, SheetRect, StylesExtras, TableStyleDef, TableStyleDiagnostic, TableStyleElementType } from './model'
import type { TableMetadata } from '../formula/types'
import type { ThemeContext } from '../../drawing/style'
import type { XlsxCellStyle, XlsxSheet } from '../types'
import { parseThemeContext } from '../../drawing/style'
import { registerTableStyleEngine } from '../tableStyleLoader'
import { pivotEngine } from '../pivotLoader'
import { MAX_COLS, MAX_ROWS } from '../view/coordinates'
import {
  MAX_REGION_STYLE_MEMO_ENTRIES, REGION_STYLE_LAYER_BASE_BYTES,
  REGION_STYLE_MEMO_ENTRY_ESTIMATE_BYTES, REGION_STYLE_REGION_ESTIMATE_BYTES,
} from '../view/contracts'

/** Page mode and export draw what Excel prints; the interactive viewer what it shows on screen (§7.10). */
export type RegionStyleMode = 'screen' | 'print'
export interface RegionStyleLayer {
  readonly rects: readonly SheetRect[]
  /** Conservative retained-size estimate; independent of the table's cell count. */
  readonly retainedBytes: number
  /** `mode` (default print) matters only beside pivot ± buttons, which move compact labels. */
  styleAt(row: number, col: number, own: XlsxCellStyle | undefined, mode?: RegionStyleMode): XlsxCellStyle | undefined
}
/** What one cell of a styled region layers: roles in ascending precedence, then
 * dxf overrides (a pivot's `formats`, named by `overrideKey` for the memo), and
 * a layout indent that replaces the cell's own. */
export interface RegionCell { roles: readonly RoleRef[]; overrides?: readonly DxfStyle[]; overrideKey?: string; indentPx?: number }
/** A styled table or pivot: its bounds, its style and its per-cell layering. */
export interface StyleRegion {
  readonly rects: readonly SheetRect[]
  readonly def: TableStyleDef
  readonly retainedBytes: number
  cell(row: number, col: number, mode: RegionStyleMode): RegionCell
}
export type RegionStyleTable = Pick<TableMetadata, 'id' | 'displayName' | 'sheetId' | 'extent' | 'headerRowCount' | 'totalsRowCount' | 'styleInfo'>

/** A bounded cache stores at most 4,096 role/own-style pairs at the 1,024 B
 * estimate, plus 4 KiB fixed overhead and 512 B per region. */
export { MAX_REGION_STYLE_MEMO_ENTRIES, MAX_REGION_STYLE_TABLES, REGION_STYLE_LAYER_BASE_BYTES, REGION_STYLE_LAYER_MAX_BYTES, REGION_STYLE_MEMO_ENTRY_ESTIMATE_BYTES, REGION_STYLE_REGION_ESTIMATE_BYTES } from '../view/contracts'

function band(element: TableStyleElementType, rect: SheetRect, row: number, col: number): RoleRef {
  return { element, edges: { top: row === rect.minRow, bottom: row === rect.maxRow, left: col === rect.minCol, right: col === rect.maxCol } }
}

/** The `[first, last]` rows (or columns) of the stripe holding `index`, for stripes of `first` and `second` lines from `start`. */
export function stripeBand(start: number, index: number, first: number, second: number): [number, number] {
  const period = first + second
  const periods = Math.floor(index / period)
  const remainder = index % period
  const inFirst = remainder < first
  const offset = periods * period + (inFirst ? 0 : first)
  const size = inFirst ? first : second
  return [start + offset, start + offset + size - 1]
}

/** Roles use source coordinates and carry the actual outer edges of each band. */
export function tableCellRoles(table: RegionStyleTable, row: number, col: number, def?: TableStyleDef): RoleRef[] {
  const { firstRow, firstCol, rows, cols } = table.extent
  if (![firstRow, firstCol, rows, cols, row, col].every(Number.isSafeInteger) || firstRow < 0 || firstCol < 0 || rows < 1 || cols < 1) return []
  const lastRow = firstRow + rows - 1, lastCol = firstCol + cols - 1
  if (row < firstRow || row > lastRow || col < firstCol || col > lastCol) return []
  const rect: SheetRect = { minRow: firstRow, minCol: firstCol, maxRow: lastRow, maxCol: lastCol }
  const out: RoleRef[] = [band('wholeTable', rect, row, col)]
  const headerCount = table.headerRowCount
  const totalCount = table.totalsRowCount
  const dataFirstRow = firstRow + headerCount, dataLastRow = lastRow - totalCount
  const dataFirstCol = firstCol, dataLastCol = lastCol
  const dataCell = row >= dataFirstRow && row <= dataLastRow
  const body: SheetRect = { minRow: dataFirstRow, minCol: dataFirstCol, maxRow: dataLastRow, maxCol: dataLastCol }

  // SpreadsheetML stripes alternate sized bands. Each size is read from its
  // matching style element; absent sizes default to one row/column.
  const rowFirst = Math.max(1, Math.min(dataLastRow - dataFirstRow + 1, def?.elements.firstRowStripe?.size ?? 1))
  const rowSecond = Math.max(1, Math.min(dataLastRow - dataFirstRow + 1, def?.elements.secondRowStripe?.size ?? 1))
  const colFirst = Math.max(1, Math.min(dataLastCol - dataFirstCol + 1, def?.elements.firstColumnStripe?.size ?? 1))
  const colSecond = Math.max(1, Math.min(dataLastCol - dataFirstCol + 1, def?.elements.secondColumnStripe?.size ?? 1))
  const info = table.styleInfo
  if (info?.showRowStripes && dataCell) {
    const [lo, hi] = stripeBand(dataFirstRow, row - dataFirstRow, rowFirst, rowSecond)
    const stripe: SheetRect = { minRow: lo, minCol: firstCol, maxRow: Math.min(hi, dataLastRow), maxCol: lastCol }
    out.push(band(((row - dataFirstRow) % (rowFirst + rowSecond)) < rowFirst ? 'firstRowStripe' : 'secondRowStripe', stripe, row, col))
  }
  if (info?.showColumnStripes && dataCell) {
    const [lo, hi] = stripeBand(dataFirstCol, col - dataFirstCol, colFirst, colSecond)
    const stripe: SheetRect = { minRow: dataFirstRow, minCol: lo, maxRow: dataLastRow, maxCol: Math.min(hi, dataLastCol) }
    out.push(band(((col - dataFirstCol) % (colFirst + colSecond)) < colFirst ? 'firstColumnStripe' : 'secondColumnStripe', stripe, row, col))
  }
  if (info?.showFirstColumn && dataCell && col === firstCol) out.push(band('firstColumn', { ...body, minCol: firstCol, maxCol: firstCol }, row, col))
  if (info?.showLastColumn && dataCell && col === lastCol) out.push(band('lastColumn', { ...body, minCol: lastCol, maxCol: lastCol }, row, col))
  if (headerCount && row < dataFirstRow) out.push(band('headerRow', { minRow: firstRow, minCol: firstCol, maxRow: dataFirstRow - 1, maxCol: lastCol }, row, col))
  if (totalCount && row > dataLastRow) out.push(band('totalRow', { minRow: dataLastRow + 1, minCol: firstCol, maxRow: lastRow, maxCol: lastCol }, row, col))
  if (headerCount && row < dataFirstRow && info?.showFirstColumn && col === firstCol) out.push(band('firstHeaderCell', { minRow: row, minCol: col, maxRow: row, maxCol: col }, row, col))
  if (headerCount && row < dataFirstRow && info?.showLastColumn && col === lastCol) out.push(band('lastHeaderCell', { minRow: row, minCol: col, maxRow: row, maxCol: col }, row, col))
  if (totalCount && row > dataLastRow && info?.showFirstColumn && col === firstCol) out.push(band('firstTotalCell', { minRow: row, minCol: col, maxRow: row, maxCol: col }, row, col))
  if (totalCount && row > dataLastRow && info?.showLastColumn && col === lastCol) out.push(band('lastTotalCell', { minRow: row, minCol: col, maxRow: row, maxCol: col }, row, col))
  return out
}

/** Table regions for `tables` (unknown and unplaceable styles reported and skipped). */
export function tableStyleRegions(tables: readonly RegionStyleTable[], extras: StylesExtras, report?: (diagnostic: TableStyleDiagnostic) => void): StyleRegion[] {
  return tables.flatMap(table => {
    const name = table.styleInfo?.name
    if (!name || table.extent.rows < 1 || table.extent.cols < 1) return []
    // A named custom style intentionally replaces a built-in with the same name.
    const def = extras.tableStyles.get(name) ?? builtinTableStyle(name)
    if (!def || !def.table) {
      const message = `Table "${table.displayName}" uses unknown table style "${name}"; stored cell formatting is used.`
      report?.({ kind: 'unsupported', feature: 'table-style-unknown', message })
      return []
    }
    const rect: SheetRect = { minRow: table.extent.firstRow, minCol: table.extent.firstCol, maxRow: table.extent.firstRow + table.extent.rows - 1, maxCol: table.extent.firstCol + table.extent.cols - 1 }
    if (![rect.minRow, rect.minCol, rect.maxRow, rect.maxCol].every(Number.isSafeInteger) || rect.minRow < 0 || rect.minCol < 0 || rect.maxRow >= MAX_ROWS || rect.maxCol >= MAX_COLS) return []
    return [{ rects: [rect], def, retainedBytes: REGION_STYLE_REGION_ESTIMATE_BYTES, cell: (row: number, col: number) => ({ roles: orderRoles(tableCellRoles(table, row, col, def), 'table') }) }]
  })
}

export function buildTableRegions(tables: readonly RegionStyleTable[], extras: StylesExtras, theme: ThemeContext, report?: (diagnostic: TableStyleDiagnostic) => void): RegionStyleLayer | undefined {
  return composeRegionLayer(tableStyleRegions(tables, extras, report), theme, report)
}

/** Build the same composed layer from transferred descriptors on a paged sheet. */
export function buildSheetRegionLayer(sheet: XlsxSheet, extras: StylesExtras, theme: ThemeContext, report?: (diagnostic: TableStyleDiagnostic) => void): RegionStyleLayer | undefined {
  const pivots = sheet.pivotTables?.length ? pivotEngine()?.regions(sheet, extras, theme, report ?? (() => {})) ?? [] : []
  return composeRegionLayer([...tableStyleRegions(sheet.tables ?? [], extras, report), ...pivots], theme, report)
}

const inRect = (rect: SheetRect, row: number, col: number) => row >= rect.minRow && row <= rect.maxRow && col >= rect.minCol && col <= rect.maxCol

/** One paint layer over styled tables and pivots (they never overlap in Excel; the first match wins). */
export function composeRegionLayer(regions: readonly StyleRegion[], theme: ThemeContext, report?: (diagnostic: TableStyleDiagnostic) => void): RegionStyleLayer | undefined {
  if (!regions.length) return undefined
  const rects = regions.flatMap(region => region.rects)
  const memo = new Map<StyleRegion, Map<string, Map<XlsxCellStyle | undefined, XlsxCellStyle | undefined>>>()
  let memoEntries = 0
  return {
    rects,
    get retainedBytes() { return REGION_STYLE_LAYER_BASE_BYTES + regions.reduce((sum, region) => sum + region.retainedBytes, 0) + memoEntries * REGION_STYLE_MEMO_ENTRY_ESTIMATE_BYTES },
    styleAt(row, col, own, mode = 'print') {
      const region = regions.find(candidate => candidate.rects.some(rect => inRect(rect, row, col)))
      if (!region) return own
      const { roles, overrides, overrideKey, indentPx } = region.cell(row, col, mode)
      if (!roles.length && !overrides?.length && indentPx === undefined) return own
      const signature = `${roles.map(role => typeof role === 'string' ? role : `${role.element}:${+role.edges.top}${+role.edges.bottom}${+role.edges.left}${+role.edges.right}`).join('|')}#${overrideKey ?? ''}@${indentPx ?? ''}`
      let bySignature = memo.get(region)
      if (!bySignature) memo.set(region, bySignature = new Map())
      const byOwn = bySignature.get(signature)
      if (byOwn?.has(own)) return byOwn.get(own)
      const scratch: TableStyleDiagnostic[] = []
      let style = roles.length || overrides?.length ? resolveRoleStyle(region.def, roles, own, theme, scratch, overrides) : own
      if (indentPx !== undefined) style = { ...(style ?? { numFmtId: 0 }), indentPx, ...(style?.horizontal === undefined || style.horizontal === 'general' ? { horizontal: 'left' } : {}) }
      for (const diagnostic of scratch) report?.(diagnostic)
      if (memoEntries >= MAX_REGION_STYLE_MEMO_ENTRIES) {
        memo.clear()
        memoEntries = 0
        bySignature = new Map()
        memo.set(region, bySignature)
      }
      let entries = bySignature.get(signature)
      if (!entries) bySignature.set(signature, entries = new Map())
      entries.set(own, style)
      memoEntries++
      return style
    },
  }
}

export function regionStylesFromSource(sheet: XlsxSheet): RegionStyleLayer | undefined {
  if (!sheet.tables?.length && !sheet.pivotTables?.length) return undefined
  const source = sheet.tableStyleSource
  const extras: StylesExtras = source ? parseStylesExtras(source.root, source.colors) : { dxfs: [], tableStyles: new Map(), slicerStyles: new Map() }
  const report = (diagnostic: TableStyleDiagnostic) => {
    const list = (sheet.diagnostics ??= []) as TableStyleDiagnostic[]
    if (!list.some(item => item.feature === diagnostic.feature && item.message === diagnostic.message)) list.push(diagnostic)
  }
  const theme = sheet.drawingTheme ?? parseThemeContext()
  // Pivot regions come from the lazy pivots chunk once it is loaded (tableStyleLoader waits for it).
  const pivots = sheet.pivotTables?.length ? pivotEngine()?.regions(sheet, extras, theme, report) ?? [] : []
  return composeRegionLayer([...tableStyleRegions(sheet.tables ?? [], extras, report), ...pivots], theme, report)
}

registerTableStyleEngine(regionStylesFromSource)
