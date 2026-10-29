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
import { MAX_COLS } from '../view/coordinates'
import { PIVOT_VALUES_FIELD, type PivotArea, type PivotTableDefinition } from './model'
import { savedPivotLayout, savedPivotRoles, type PivotLine, type SavedPivotLayout } from './savedLayout'
import { compactLabelIndents } from './buttons'

const NO_STYLE: TableStyleDef = { name: '', pivot: true, table: false, elements: {} }
/** Retained per saved line: the resolved item vector and its bookkeeping. */
const LINE_ESTIMATE_BYTES = 64
/** Cells whose `formats` overrides are kept, and the estimate per kept cell (its key, list and joined ids). */
const FORMAT_MEMO_ENTRIES = 4_096
const FORMAT_MEMO_ENTRY_ESTIMATE_BYTES = 128

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

/** Where a cell sits in the location; the rectangle area types depend on nothing else. */
type Zone = 'outside' | 'origin' | 'data' | 'label'
const ZONES: readonly Zone[] = ['outside', 'origin', 'data', 'label']
function zoneOf(layout: SavedPivotLayout, row: number, col: number): Zone {
  const { rect, dataRow0, dataCol0 } = layout
  if (row < rect.minRow || row > rect.maxRow || col < rect.minCol || col > rect.maxCol) return 'outside'
  if (row >= dataRow0 && col >= dataCol0) return 'data'
  return row < dataRow0 && col < dataCol0 ? 'origin' : 'label'
}

/** A compiled pivotArea: a rectangle type decided by the cell's zone, or a test over the cell's coordinates
 * indexed by its most selective reference (or, for a label-only area without references, its field). */
type CompiledArea =
  | { kind: 'zone'; zones: ReadonlySet<Zone> }
  | { kind: 'cell'; test: (cell: PivotCellCoordinates) => boolean; field?: number; items?: readonly number[]; labelField?: number }

/** A pivotArea compiled once per layout, or undefined for a construct not modelled here: offsets, cache-indexed or
 * positional references, subtotal references, button/topEnd/topRight areas, and a normal area naming no cells. */
function compileArea(area: PivotArea): CompiledArea | undefined {
  if (area.offset !== undefined || area.cacheIndex) return undefined
  if (area.references.some(ref => ref.field === undefined || ref.byPosition || ref.relative || ref.defaultSubtotal || ref.subtotals.length)) return undefined
  switch (area.type) {
    case 'none': return { kind: 'zone', zones: new Set() }
    case 'origin': return { kind: 'zone', zones: new Set(['origin']) }
    case 'data': return { kind: 'zone', zones: new Set(['data']) }
    case 'all': return { kind: 'zone', zones: new Set<Zone>(area.dataOnly ? ['data'] : area.labelOnly ? ['origin', 'label'] : ['origin', 'data', 'label']) }
    case 'normal': break
    default: return undefined
  }
  if (!area.references.length && !(area.labelOnly && area.field !== undefined)) return undefined
  const test = (cell: PivotCellCoordinates) => {
    if (area.dataOnly && cell.kind !== 'data') return false
    if (area.labelOnly && cell.kind !== 'label') return false
    if ((cell.grandRow && !area.grandRow) || (cell.grandCol && !area.grandCol)) return false
    if (area.field !== undefined && cell.kind === 'label' && cell.labelField !== area.field) return false
    return area.references.every(ref => {
      const item = ref.field === PIVOT_VALUES_FIELD ? cell.data : cell.fields.get(ref.field!)
      return item !== undefined && (!ref.x.length || ref.x.includes(item))
    })
  }
  // Indexed by its most selective reference: a field other than Values (Excel writes the Values reference
  // first, and it splits cells only by data field), naming the fewest items (none means any).
  const rank = (ref: PivotArea['references'][number]) => (ref.field === PIVOT_VALUES_FIELD ? 2 : 0) + (ref.x.length ? 0 : 4) + Math.min(1, ref.x.length / 1e6)
  const key = area.references.reduce<PivotArea['references'][number] | undefined>((best, ref) => !best || rank(ref) < rank(best) ? ref : best, undefined)
  return key ? { kind: 'cell', test, field: key.field!, ...(key.x.length ? { items: key.x } : {}) } : { kind: 'cell', test, labelField: area.field! }
}

/** A pivotArea as a cell predicate (see compileArea). */
export function pivotAreaMatcher(layout: SavedPivotLayout, area: PivotArea): ((row: number, col: number) => boolean) | undefined {
  const compiled = compileArea(area)
  if (!compiled) return undefined
  if (compiled.kind === 'zone') return (row, col) => compiled.zones.has(zoneOf(layout, row, col))
  return (row, col) => {
    const cell = coordinates(layout, row, col)
    return Boolean(cell && compiled.test(cell))
  }
}

interface FormatHit { n: number; dxf: DxfStyle }
/** The formats over each cell, in document order, at a cost independent of how many formats the pivot has
 * (spec §8: resolved once per layout, not per tile): the rectangle types are grouped by zone, and the
 * others are found through one reference's (field, item) before their full test. */
function formatIndex(layout: SavedPivotLayout, formats: ReadonlyArray<FormatHit & { area: CompiledArea }>): (row: number, col: number) => FormatHit[] {
  const byZone = new Map<Zone, FormatHit[]>(ZONES.map(zone => [zone, []]))
  const byItem = new Map<number, Map<number, Array<FormatHit & { test: (cell: PivotCellCoordinates) => boolean }>>>()
  const byField = new Map<number, Array<FormatHit & { test: (cell: PivotCellCoordinates) => boolean }>>()
  const byLabel = new Map<number, Array<FormatHit & { test: (cell: PivotCellCoordinates) => boolean }>>()
  const push = <K, V>(map: Map<K, V[]>, key: K, value: V) => { const list = map.get(key); if (list) list.push(value); else map.set(key, [value]) }
  for (const { n, dxf, area } of formats) {
    if (area.kind === 'zone') { for (const zone of area.zones) byZone.get(zone)!.push({ n, dxf }); continue }
    const entry = { n, dxf, test: area.test }
    if (area.labelField !== undefined) push(byLabel, area.labelField, entry)
    else if (!area.items) push(byField, area.field!, entry)
    else {
      let items = byItem.get(area.field!)
      if (!items) byItem.set(area.field!, items = new Map())
      for (const item of new Set(area.items)) push(items, item, entry)
    }
  }
  const indexed = byItem.size > 0 || byField.size > 0 || byLabel.size > 0
  return (row, col) => {
    const zone = byZone.get(zoneOf(layout, row, col))!
    if (!indexed) return zone
    const cell = coordinates(layout, row, col)
    if (!cell) return zone
    const hits = new Map<number, FormatHit>()
    const consider = (list: ReadonlyArray<FormatHit & { test: (cell: PivotCellCoordinates) => boolean }> | undefined) => {
      for (const entry of list ?? []) if (!hits.has(entry.n) && entry.test(cell)) hits.set(entry.n, { n: entry.n, dxf: entry.dxf })
    }
    const at = (field: number, item: number | undefined) => {
      if (item === undefined) return
      consider(byItem.get(field)?.get(item))
      consider(byField.get(field))
    }
    for (const [field, item] of cell.fields) at(field, item)
    at(PIVOT_VALUES_FIELD, cell.data)
    if (cell.labelField !== undefined) consider(byLabel.get(cell.labelField))
    if (!hits.size) return zone
    for (const hit of zone) hits.set(hit.n, hit)
    return [...hits.values()].sort((a, b) => a.n - b.n)
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
      const area = dxf && compileArea(format.pivotArea)
      if (!dxf || !area) { skipped++; return [] }
      return [{ n, dxf, area }]
    })
    const formatsAt = formatIndex(layout, formats)
    // A tile asks for each cell and its four neighbours: the overrides per cell are kept (bounded).
    type Overrides = { overrides: readonly DxfStyle[]; overrideKey: string } | undefined
    const overridesMemo = new Map<number, Overrides>()
    const overridesAt = (row: number, col: number): Overrides => {
      const key = row * MAX_COLS + col
      if (overridesMemo.has(key)) return overridesMemo.get(key)
      const hits = formatsAt(row, col)
      const out = hits.length ? { overrides: hits.map(hit => hit.dxf), overrideKey: hits.map(hit => hit.n).join(',') } : undefined
      if (overridesMemo.size >= FORMAT_MEMO_ENTRIES) overridesMemo.clear()
      overridesMemo.set(key, out)
      return out
    }
    if (skipped) report({ kind: 'approximate', feature: 'pivot-format-area-unsupported', message: `Pivot table "${def.name}": ${skipped} of ${def.formats.length} format(s) are not applied; those cells keep their own formatting.` })
    const indent = compactLabelIndents(layout)
    return [{
      rects: roles.rects,
      def: style ?? NO_STYLE,
      kind: 'pivot' as const,
      retainedBytes: REGION_STYLE_REGION_ESTIMATE_BYTES + (layout.rows.length + layout.cols.length) * LINE_ESTIMATE_BYTES
        + (formats.length ? FORMAT_MEMO_ENTRIES * FORMAT_MEMO_ENTRY_ESTIMATE_BYTES : 0),
      cell(row, col, mode) {
        const hits = formats.length ? overridesAt(row, col) : undefined
        const indentPx = indent(row, col, mode)
        return {
          roles: orderRoles(roles.roleAt(row, col), 'pivot'),
          ...(hits ?? {}),
          ...(indentPx !== undefined ? { indentPx } : {}),
        }
      },
    }]
  })
}
