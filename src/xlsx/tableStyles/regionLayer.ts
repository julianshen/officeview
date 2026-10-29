/** Lazy resident worksheet table-style composition. */
import { parseStylesExtras } from './parseStylesExtras'
import { builtinTableStyle } from './builtins'
import { orderRoles, resolveRoleStyle } from './resolve'
import type { RoleRef, SheetRect, StylesExtras, TableStyleDef, TableStyleDiagnostic, TableStyleElementType } from './model'
import type { TableMetadata } from '../formula/types'
import type { ThemeContext } from '../../drawing/style'
import type { XlsxCellStyle, XlsxSheet } from '../types'
import { parseThemeContext } from '../../drawing/style'
import { registerTableStyleEngine } from '../tableStyleLoader'
import { MAX_COLS, MAX_ROWS } from '../view/coordinates'

export interface RegionStyleLayer {
  readonly rects: readonly SheetRect[]
  styleAt(row: number, col: number, own: XlsxCellStyle | undefined): XlsxCellStyle | undefined
}
export type RegionStyleTable = Pick<TableMetadata, 'id' | 'displayName' | 'sheetId' | 'extent' | 'headerRowCount' | 'totalsRowCount' | 'styleInfo'>

function band(element: TableStyleElementType, rect: SheetRect, row: number, col: number): RoleRef {
  return { element, edges: { top: row === rect.minRow, bottom: row === rect.maxRow, left: col === rect.minCol, right: col === rect.maxCol } }
}

function stripeBand(start: number, index: number, first: number, second: number): [number, number] {
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

export function buildTableRegions(tables: readonly RegionStyleTable[], extras: StylesExtras, theme: ThemeContext, report?: (diagnostic: TableStyleDiagnostic) => void): RegionStyleLayer | undefined {
  const regions = tables.flatMap(table => {
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
    return [{ table, def, rect }]
  })
  if (!regions.length) return undefined
  const rects = regions.map(region => region.rect)
  const memo = new Map<string, Map<XlsxCellStyle | undefined, XlsxCellStyle | undefined>>()
  const MAX_MEMO_CELLS = 50_000
  return {
    rects,
    styleAt(row, col, own) {
      const region = regions.find(candidate => row >= candidate.rect.minRow && row <= candidate.rect.maxRow && col >= candidate.rect.minCol && col <= candidate.rect.maxCol)
      if (!region) return own
      const key = `${row}:${col}`
      const byOwn = memo.get(key)
      if (byOwn?.has(own)) return byOwn.get(own)
      const roles = orderRoles(tableCellRoles(region.table, row, col, region.def), 'table')
      const scratch: TableStyleDiagnostic[] = []
      const style = roles.length ? resolveRoleStyle(region.def, roles, own, theme, scratch) : own
      for (const diagnostic of scratch) report?.(diagnostic)
      let entries = byOwn
      if (!entries) {
        // Delete in batches: Map iterator deletion leaves tombstones in some
        // engines, making per-cell FIFO eviction quadratic for large tables.
        if (memo.size >= MAX_MEMO_CELLS) memo.clear()
        entries = new Map<XlsxCellStyle | undefined, XlsxCellStyle | undefined>()
        memo.set(key, entries)
      }
      entries.set(own, style)
      return style
    },
  }
}

export function regionStylesFromSource(sheet: XlsxSheet): RegionStyleLayer | undefined {
  if (!sheet.tables?.length) return undefined
  const source = sheet.tableStyleSource
  const extras: StylesExtras = source ? parseStylesExtras(source.root, source.colors) : { dxfs: [], tableStyles: new Map(), slicerStyles: new Map() }
  const report = (diagnostic: TableStyleDiagnostic) => {
    const list = (sheet.diagnostics ??= []) as TableStyleDiagnostic[]
    if (!list.some(item => item.feature === diagnostic.feature && item.message === diagnostic.message)) list.push(diagnostic)
  }
  return buildTableRegions(sheet.tables, extras, sheet.drawingTheme ?? parseThemeContext(), report)
}

registerTableStyleEngine(regionStylesFromSource)
