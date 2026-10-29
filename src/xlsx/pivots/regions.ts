/**
 * Pivot regions of the sheet's RegionStyleLayer (spec §6.5): each pivot's saved
 * roles, its `formats` dxf overrides by pivotArea, and the compact-label indent
 * beside ± buttons. Lazy chunk `pivots`; merged with table regions by
 * tableStyles/regionLayer.ts.
 *
 * Precedence: style roles (PIVOT_STYLE_PRECEDENCE), then `formats` in document
 * order, then the cell's own `s=` (measure-cell-override, Ruling P). Excel
 * writes a format twice — as a `<format>` and into the cell's `s=` — so a
 * pivotArea this matcher does not model is skipped with a diagnostic and still
 * shows through the cell's own format.
 *
 * Pivot conditional formats: no conditional-format engine exists in src, so
 * they are reported (`pivot-conditional-format`) and skipped.
 */
import type { ThemeContext } from '../../drawing/style'
import type { XlsxSheet } from '../types'
import type { DxfStyle, StylesExtras, TableStyleDef, TableStyleDiagnostic } from '../tableStyles/model'
import type { StyleRegion } from '../tableStyles/regionLayer'
import { builtinTableStyle } from '../tableStyles/builtins'
import { orderRoles } from '../tableStyles/resolve'
import { REGION_STYLE_REGION_ESTIMATE_BYTES } from '../view/contracts'
import { PIVOT_VALUES_FIELD, type PivotArea, type PivotTableDefinition } from './model'
import { savedPivotLayout, savedPivotRoles, type PivotLine, type SavedPivotLayout } from './savedLayout'
import { compactLabelIndents } from './buttons'

const NO_STYLE: TableStyleDef = { name: '', pivot: true, table: false, elements: {} }
/** Retained per saved line: the resolved item vector and its bookkeeping. */
const LINE_ESTIMATE_BYTES = 64

/** A pivot cell's coordinates: the field items it stands for, its data field, and what kind of cell it is. */
interface PivotCellCoordinates {
  kind: 'data' | 'label'
  fields: Map<number, number>
  data?: number
  /** For a label: the field whose item it shows. */
  labelField?: number
  grandRow: boolean
  grandCol: boolean
}

function lineFields(out: Map<number, number>, axis: readonly number[], line: PivotLine, through: number): number | undefined {
  let data: number | undefined
  for (let level = 0; level <= through && level < line.items.length; level++) {
    const field = axis[level]
    if (field === PIVOT_VALUES_FIELD) data = line.data
    else if (field !== undefined && line.items[level]! >= 0) out.set(field, line.items[level]!)
  }
  return data
}

function coordinates(layout: SavedPivotLayout, row: number, col: number): PivotCellCoordinates | undefined {
  const { def, rect, dataRow0, dataCol0, itemRow0, rows, cols, labelCol } = layout
  if (row < rect.minRow || row > rect.maxRow || col < rect.minCol || col > rect.maxCol) return undefined
  const rowLine = row >= dataRow0 ? rows[row - dataRow0] : undefined
  const colLine = col >= dataCol0 ? cols[col - dataCol0] : undefined
  const valuesOnRows = def.rowFields.includes(PIVOT_VALUES_FIELD), valuesOnCols = def.colFields.includes(PIVOT_VALUES_FIELD)
  if (rowLine && colLine) {
    const fields = new Map<number, number>()
    const rowData = lineFields(fields, def.rowFields, rowLine, Infinity)
    const colData = lineFields(fields, def.colFields, colLine, Infinity)
    const data = valuesOnRows ? rowData : valuesOnCols ? colData : 0
    return { kind: 'data', fields, ...(data !== undefined ? { data } : {}), grandRow: rowLine.type === 'grand', grandCol: colLine.type === 'grand' }
  }
  if (rowLine) {
    if (rowLine.type === 'grand') return col === rect.minCol ? { kind: 'label', fields: new Map(), grandRow: true, grandCol: false } : undefined
    let level = -1
    for (let k = rowLine.first; k < rowLine.items.length; k++) if (labelCol[k] === col) level = k
    if (level < 0) return undefined
    const fields = new Map<number, number>()
    const data = lineFields(fields, def.rowFields, rowLine, level)
    return { kind: 'label', fields, ...(data !== undefined ? { data } : {}), labelField: def.rowFields[level], grandRow: false, grandCol: false }
  }
  if (colLine && row >= itemRow0) {
    const level = row - itemRow0
    if (colLine.type === 'grand') return level === 0 ? { kind: 'label', fields: new Map(), grandRow: false, grandCol: true } : undefined
    if (level < colLine.first || level >= colLine.items.length) return undefined
    const fields = new Map<number, number>()
    const data = lineFields(fields, def.colFields, colLine, level)
    return { kind: 'label', fields, ...(data !== undefined ? { data } : {}), labelField: def.colFields[level], grandRow: false, grandCol: false }
  }
  return undefined
}

/** A pivotArea as a cell predicate, or undefined for a construct not modelled
 * here: offsets, cache-indexed or positional references, subtotal references,
 * button/topEnd/topRight areas, and a normal area naming no cells. */
export function pivotAreaMatcher(layout: SavedPivotLayout, area: PivotArea): ((row: number, col: number) => boolean) | undefined {
  if (area.offset !== undefined || area.cacheIndex) return undefined
  if (area.references.some(ref => ref.field === undefined || ref.byPosition || ref.relative || ref.defaultSubtotal || ref.subtotals.length)) return undefined
  const { rect, dataRow0, dataCol0 } = layout
  const inside = (row: number, col: number) => row >= rect.minRow && row <= rect.maxRow && col >= rect.minCol && col <= rect.maxCol
  switch (area.type) {
    case 'none': return () => false
    case 'origin': return (row, col) => inside(row, col) && row < dataRow0 && col < dataCol0
    case 'data': return (row, col) => inside(row, col) && row >= dataRow0 && col >= dataCol0
    case 'all': return (row, col) => {
      if (!inside(row, col)) return false
      const data = row >= dataRow0 && col >= dataCol0
      return area.dataOnly ? data : area.labelOnly ? !data : true
    }
    case 'normal': break
    default: return undefined
  }
  if (!area.references.length && !(area.labelOnly && area.field !== undefined)) return undefined
  return (row, col) => {
    const cell = coordinates(layout, row, col)
    if (!cell) return false
    if (area.dataOnly && cell.kind !== 'data') return false
    if (area.labelOnly && cell.kind !== 'label') return false
    if ((cell.grandRow && !area.grandRow) || (cell.grandCol && !area.grandCol)) return false
    if (area.field !== undefined && cell.kind === 'label' && cell.labelField !== area.field) return false
    return area.references.every(ref => {
      const item = ref.field === PIVOT_VALUES_FIELD ? cell.data : cell.fields.get(ref.field!)
      return item !== undefined && (!ref.x.length || ref.x.includes(item))
    })
  }
}

export function pivotStyleRegions(sheet: XlsxSheet, extras: StylesExtras, _theme: ThemeContext, report: (diagnostic: TableStyleDiagnostic) => void): StyleRegion[] {
  return (sheet.pivotTables ?? []).flatMap((def: PivotTableDefinition): StyleRegion[] => {
    const name = def.styleInfo?.name
    // A named custom style replaces a built-in of the same name; a pivot without one has no style elements.
    let style = name ? extras.tableStyles.get(name) ?? builtinTableStyle(name) : undefined
    if (name && (!style || !style.pivot)) {
      report({ kind: 'unsupported', feature: 'pivot-style-unknown', message: `Pivot table "${def.name}" uses unknown pivot style "${name}"; stored cell formatting is used.` })
      style = undefined
    }
    if (def.conditionalFormats.length) report({ kind: 'unsupported', feature: 'pivot-conditional-format', message: `Pivot table "${def.name}" has ${def.conditionalFormats.length} conditional format(s); they are not shown.` })
    const layout = savedPivotLayout(def)
    const roles = savedPivotRoles(def, style, layout)
    if (!layout || !roles) {
      report({ kind: 'unsupported', feature: 'pivot-layout-unavailable', message: `Pivot table "${def.name}" has no usable saved layout; its cells paint with their stored formatting only.` })
      return []
    }
    let skipped = 0
    const formats = def.formats.flatMap((format, n) => {
      if (format.action !== 'formatting' || format.dxfId === undefined) return []
      const dxf: DxfStyle | undefined = extras.dxfs[format.dxfId]
      const match = dxf && pivotAreaMatcher(layout, format.pivotArea)
      if (!dxf || !match) { skipped++; return [] }
      return [{ n, dxf, match }]
    })
    if (skipped) report({ kind: 'approximate', feature: 'pivot-format-area-unsupported', message: `Pivot table "${def.name}": ${skipped} of ${def.formats.length} format(s) are not applied; those cells keep their own formatting.` })
    const indent = compactLabelIndents(layout)
    return [{
      rects: roles.rects,
      def: style ?? NO_STYLE,
      kind: 'pivot' as const,
      retainedBytes: REGION_STYLE_REGION_ESTIMATE_BYTES + (layout.rows.length + layout.cols.length) * LINE_ESTIMATE_BYTES,
      cell(row, col, mode) {
        const hits = formats.filter(format => format.match(row, col))
        const indentPx = indent(row, col, mode)
        return {
          roles: orderRoles(roles.roleAt(row, col), 'pivot'),
          ...(hits.length ? { overrides: hits.map(hit => hit.dxf), overrideKey: hits.map(hit => hit.n).join(',') } : {}),
          ...(indentPx !== undefined ? { indentPx } : {}),
        }
      },
    }]
  })
}
