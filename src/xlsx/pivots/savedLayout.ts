/**
 * The saved pivot layout (spec §3, §6.5): `location` and the saved
 * `rowItems`/`colItems` mapped onto sheet cells, and each cell's style roles.
 * Lazy chunk `pivots`. Nothing here recomputes a layout: it reads the one Excel
 * saved, and a definition whose indices are invalid is never laid out (§7.6,
 * Ruling AC).
 *
 * Measured (Excel 16.114, docs/validation/xlsx-pivots.md "Element precedence";
 * refs/measure-precedence.png, pivot-subtotals-bottom, pivot-values-on-rows,
 * pivot-layout-outline, pivot-layout-tabular, pivot-page-field):
 * - headerRow is every row above the first data row; firstHeaderCell is the
 *   row-label columns of those rows except the last (the "Row Labels" row);
 * - the column-item rows (from `firstHeaderRow`) carry one column subheading
 *   per level over the item columns, the leaf level included;
 * - lastColumn is the grand-total columns from the column-item rows down, never
 *   the top header row (Ruling AE), and through the grand-total row;
 * - firstColumn is the row-label columns over the location's whole height,
 *   header rows included, under headerRow and firstHeaderCell there
 *   (refs/pivot-light-9.png: its right line at x 77 runs through rows 3–5;
 *   refs/pivot-medium-23.png: its #83CCEB fill in A3:A5, where headerRow has
 *   none; refs/measure-precedence.png: headerRow still wins A5);
 * - a line whose deepest item has children is a row subheading for its level
 *   across the whole row (compact, outline); in tabular form a parent item is a
 *   subheading in its own label column, from its line to its last child line;
 * - every subtotal line is its own band (adjacent lines are two bands); in a
 *   tabular parent's label column its top edge is inner, so the parent's
 *   bottom line shows there;
 * - stripes alternate from the first data row, subheading lines included;
 * - page fields sit above the location, one blank row between.
 * Unmeasured, from ECMA-376 §18.10 / §18.8.40: second and third subtotal rows,
 * subtotal columns, blankRow, the style-info flag gating (showRowHeaders →
 * firstColumn, firstHeaderCell and row subheadings; showColHeaders → headerRow
 * and column subheadings), and more than one page field.
 */
import type { BandEdges, RoleRef, SheetRect, TableStyleDef, TableStyleElementType } from '../tableStyles/model'
import { stripeBand } from '../tableStyles/regionLayer'
import { PIVOT_VALUES_FIELD, type PivotAxisItem, type PivotItemType, type PivotPageField, type PivotTableDefinition } from './model'

/** One saved `rowItems`/`colItems` line with its leading repeats resolved. */
export interface PivotLine {
  type: PivotItemType
  /** The first level this line writes (`r`); its deepest level is `items.length - 1`. */
  first: number
  /** Pivot-field item index per level, 0 … deepest (-1 where nothing precedes a repeat). */
  items: number[]
  /** `i`: the data-field index on the axis carrying the Values field. */
  data: number
}
export interface PivotPageCell { row: number; col: number; field: PivotPageField }
export interface SavedPivotLayout {
  def: PivotTableDefinition
  rect: SheetRect
  /** The location plus each page field's label/value pair. */
  rects: SheetRect[]
  /** Rows: first column-item row, first data row; first data column. */
  itemRow0: number
  dataRow0: number
  dataCol0: number
  rows: PivotLine[]
  cols: PivotLine[]
  /** Sheet column of each row-field level's labels. */
  labelCol: number[]
  /** Page fields' label cells (the value is the next column). */
  pages: PivotPageCell[]
  /** Last data row / column before the grand totals. */
  bodyRowEnd: number
  bodyColEnd: number
  /** Tabular parents (an item sharing its line with its first child), by every line their band covers. */
  parentBands: Map<number, ParentBand[]>
}
/** A parent item's label band: `col`, lines `line` … `last`. */
export interface ParentBand { level: number; col: number; line: number; last: number }

const CELL = /^\$?([A-Z]{1,3})\$?([1-9]\d{0,6})$/
function locationRect(ref: string): SheetRect | undefined {
  const [from, to = from] = ref.split(':')
  const a = CELL.exec(from ?? ''), b = CELL.exec(to)
  if (!a || !b) return undefined
  const col = (letters: string) => [...letters].reduce((n, c) => n * 26 + c.charCodeAt(0) - 64, 0) - 1
  const rect = { minRow: Number(a[2]) - 1, minCol: col(a[1]!), maxRow: Number(b[2]) - 1, maxCol: col(b[1]!) }
  return rect.minRow <= rect.maxRow && rect.minCol <= rect.maxCol ? rect : undefined
}

function axisLines(items: readonly PivotAxisItem[]): PivotLine[] {
  let previous: number[] = []
  return items.map(item => {
    const first = Math.max(0, item.r)
    const resolved = previous.slice(0, first)
    while (resolved.length < first) resolved.push(-1)
    resolved.push(...item.x)
    if (item.type !== 'grand') previous = resolved
    return { type: item.type, first, items: item.type === 'grand' ? [] : resolved, data: item.i }
  })
}

const isSubtotal = (type: PivotItemType) => type !== 'data' && type !== 'grand' && type !== 'blank'

/** Whether the field after `field` shares its label column (compact form). */
export function sharesColumn(def: PivotTableDefinition, field: number): boolean {
  if (field === PIVOT_VALUES_FIELD) return def.compactData && def.outlineData
  const f = def.fields[field]
  return Boolean(f && f.compact && f.outline)
}

/** The saved layout, or undefined when it cannot be laid out: no `rowItems`/
 * `colItems`, an invalid index (Ruling AC), or saved items that do not fill the location. */
export function savedPivotLayout(def: PivotTableDefinition): SavedPivotLayout | undefined {
  if (!def.rowItems || !def.colItems || def.unsupported.some(entry => entry.startsWith('invalid-index:'))) return undefined
  const rect = locationRect(def.location.ref)
  const { firstHeaderRow, firstDataRow, firstDataCol } = def.location
  if (!rect || !(firstHeaderRow >= 0 && firstHeaderRow <= firstDataRow && firstDataCol >= 0)) return undefined
  const bodyRows = rect.maxRow - rect.minRow + 1 - firstDataRow, bodyCols = rect.maxCol - rect.minCol + 1 - firstDataCol
  // An axis with nothing on it saves one empty line (`<i/>`) and may leave no body cells.
  const fits = (lines: number, cells: number) => lines === cells || (cells === 0 && lines <= 1)
  if (bodyRows < 0 || bodyCols < 0 || !fits(def.rowItems.length, bodyRows) || !fits(def.colItems.length, bodyCols)) return undefined
  const dataRow0 = rect.minRow + firstDataRow, dataCol0 = rect.minCol + firstDataCol
  const rows = axisLines(def.rowItems).slice(0, bodyRows), cols = axisLines(def.colItems).slice(0, bodyCols)
  const labelCol: number[] = []
  def.rowFields.forEach((_field, level) => {
    const previous = labelCol[level - 1]
    labelCol.push(previous === undefined ? rect.minCol : Math.min(previous + (sharesColumn(def, def.rowFields[level - 1]!) ? 0 : 1), Math.max(rect.minCol, dataCol0 - 1)))
  })
  const firstGrand = (lines: PivotLine[]) => { const n = lines.findIndex(line => line.type === 'grand'); return n < 0 ? lines.length : n }
  const pages: PivotPageCell[] = []
  const rects: SheetRect[] = [rect]
  const pageRows = def.location.rowPageCount || def.pageFields.length, pageCols = def.location.colPageCount || 1
  def.pageFields.forEach((field, k) => {
    // "Down, then over" by default; `pageOverThenDown` fills across first. A blank row separates them from the table.
    const i = def.pageOverThenDown ? Math.floor(k / pageCols) : k % pageRows
    const j = def.pageOverThenDown ? k % pageCols : Math.floor(k / pageRows)
    const row = rect.minRow - 1 - pageRows + i, col = rect.minCol + 3 * j
    if (row < 0 || i >= pageRows) return
    pages.push({ row, col, field })
    rects.push({ minRow: row, minCol: col, maxRow: row, maxCol: col + 1 })
  })
  const parentBands = new Map<number, ParentBand[]>()
  rows.forEach((line, n) => {
    if (line.type !== 'data') return
    for (let level = line.first; level < line.items.length - 1; level++) {
      let last = n
      while (last + 1 < rows.length && rows[last + 1]!.first > level && rows[last + 1]!.type !== 'grand' && rows[last + 1]!.type !== 'blank') last++
      const parent = { level, col: labelCol[level] ?? rect.minCol, line: n, last }
      for (let k = n; k <= last; k++) {
        const list = parentBands.get(k)
        if (list) list.push(parent)
        else parentBands.set(k, [parent])
      }
    }
  })
  return {
    def, rect, rects, itemRow0: rect.minRow + firstHeaderRow, dataRow0, dataCol0, rows, cols, labelCol, pages,
    bodyRowEnd: dataRow0 + firstGrand(rows) - 1, bodyColEnd: dataCol0 + firstGrand(cols) - 1, parentBands,
  }
}

const NTH = ['first', 'second', 'third'] as const
const nth = (level: number, kind: 'SubtotalRow' | 'SubtotalColumn' | 'RowSubheading' | 'ColumnSubheading') =>
  `${NTH[Math.min(2, Math.max(0, level))]}${kind}` as TableStyleElementType

export interface SavedPivotRoles {
  /** The saved location (wholeTable). */
  rect: SheetRect
  /** Every painted rectangle: the location and the page fields. */
  rects: SheetRect[]
  roleAt(row: number, col: number): RoleRef[]
}

/** Each cell's style roles from the saved layout, with the bands their borders
 * follow (Ruling AE); `style` supplies stripe sizes. Undefined when the layout
 * cannot be laid out: the cells then keep their own `s=` (pivot-layout-unavailable). */
export function savedPivotRoles(def: PivotTableDefinition, style?: TableStyleDef, layout = savedPivotLayout(def)): SavedPivotRoles | undefined {
  if (!layout) return undefined
  const { rect, dataRow0, dataCol0, itemRow0, rows, cols, bodyRowEnd, bodyColEnd } = layout
  const info = def.styleInfo
  const leafLevel = def.rowFields.length - 1
  const band = (element: TableStyleElementType, minRow: number, minCol: number, maxRow: number, maxCol: number, row: number, col: number): RoleRef => {
    const edges: BandEdges = { top: row === minRow, bottom: row === maxRow, left: col === minCol, right: col === maxCol }
    return { element, edges }
  }
  const size = (type: TableStyleElementType) => Math.max(1, style?.elements[type]?.size ?? 1)
  const grandCol0 = bodyColEnd + 1
  return {
    rect,
    rects: layout.rects,
    roleAt(row, col) {
      const page = layout.pages.find(cell => cell.row === row && (cell.col === col || cell.col + 1 === col))
      if (page) return [band(page.col === col ? 'pageFieldLabels' : 'pageFieldValues', row, col, row, col, row, col)]
      if (row < rect.minRow || row > rect.maxRow || col < rect.minCol || col > rect.maxCol) return []
      const out: RoleRef[] = [band('wholeTable', rect.minRow, rect.minCol, rect.maxRow, rect.maxCol, row, col)]
      const line = col >= dataCol0 ? cols[col - dataCol0] : undefined
      if (line?.type === 'grand' && info?.showLastColumn && row >= Math.max(itemRow0, rect.minRow + 1)) out.push(band('lastColumn', Math.max(itemRow0, rect.minRow + 1), grandCol0, rect.maxRow, rect.maxCol, row, col))
      if (line && isSubtotal(line.type) && row >= itemRow0) out.push(band(nth(line.items.length - 1, 'SubtotalColumn'), itemRow0, col, rect.maxRow, col, row, col))
      if (info?.showRowHeaders && col < dataCol0) out.push(band('firstColumn', rect.minRow, rect.minCol, rect.maxRow, dataCol0 - 1, row, col))
      if (row < dataRow0) {
        if (info?.showColHeaders) out.push(band('headerRow', rect.minRow, rect.minCol, dataRow0 - 1, rect.maxCol, row, col))
        if (info?.showRowHeaders && col < dataCol0 && row < dataRow0 - 1) out.push(band('firstHeaderCell', rect.minRow, rect.minCol, dataRow0 - 2, dataCol0 - 1, row, col))
        const level = row - itemRow0
        if (info?.showColHeaders && line && line.type !== 'grand' && level >= 0 && level < def.colFields.length) out.push(band(nth(level, 'ColumnSubheading'), row, dataCol0, row, Math.max(dataCol0, grandCol0 - 1), row, col))
        return out
      }
      const n = row - dataRow0, rowLine = rows[n]
      if (info?.showRowStripes && row <= bodyRowEnd) {
        const first = size('firstRowStripe'), period = first + size('secondRowStripe')
        const [lo, hi] = stripeBand(dataRow0, n, first, size('secondRowStripe'))
        out.push(band(n % period < first ? 'firstRowStripe' : 'secondRowStripe', lo, rect.minCol, Math.min(hi, bodyRowEnd), rect.maxCol, row, col))
      }
      if (info?.showColStripes && row <= bodyRowEnd && col >= dataCol0 && col <= bodyColEnd) {
        const k = col - dataCol0, period = size('firstColumnStripe') + size('secondColumnStripe')
        const [lo, hi] = stripeBand(dataCol0, k, size('firstColumnStripe'), size('secondColumnStripe'))
        out.push(band(k % period < size('firstColumnStripe') ? 'firstColumnStripe' : 'secondColumnStripe', dataRow0, lo, bodyRowEnd, Math.min(hi, bodyColEnd), row, col))
      }
      if (!rowLine) return out
      const level = rowLine.items.length - 1
      if (rowLine.type === 'grand') out.push(band('totalRow', row, rect.minCol, row, rect.maxCol, row, col))
      else if (rowLine.type === 'blank') out.push(band('blankRow', row, rect.minCol, row, rect.maxCol, row, col))
      else if (isSubtotal(rowLine.type)) {
        const role = band(nth(level, 'SubtotalRow'), row, rect.minCol, row, rect.maxCol, row, col)
        // Under a tabular parent's label band the band's bottom edge is kept: the
        // subtotal's top is no outer edge in that column (refs/pivot-layout-tabular.png:
        // #44B3E1 under A8, #156082 from column B on).
        if (typeof role !== 'string' && info?.showRowHeaders && col < dataCol0 && layout.parentBands.get(n - 1)?.some(parent => parent.col === col && parent.last === n - 1)) role.edges.top = false
        out.push(role)
      }
      else if (info?.showRowHeaders && level >= 0 && level < leafLevel) out.push(band(nth(level, 'RowSubheading'), row, rect.minCol, row, rect.maxCol, row, col))
      // A tabular parent: its label column, from its line through its last child line.
      if (info?.showRowHeaders && col < dataCol0) for (const parent of layout.parentBands.get(n) ?? []) {
        if (parent.col === col) out.push(band(nth(parent.level, 'RowSubheading'), dataRow0 + parent.line, col, dataRow0 + parent.last, col, row, col))
      }
      return out
    },
  }
}
