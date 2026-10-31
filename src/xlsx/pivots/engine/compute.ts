/** Complete Stage 2 pivot computation: axes, values and placement. */
import type { PivotCacheDefinition, PivotDiagnostic, PivotTableDefinition } from '../model'
import { computeAxes, type EngineResult } from './axes'
import type { PivotLayout, PlaceOptions } from './place'
import { pageFieldsFit, placeCells } from './place'
import { computeValues } from './values'
import type { RecordStore } from './records'
import type { PivotViewState } from './view'

/** A worksheet's rows and columns (zero-based rect bounds must stay below these). */
const SHEET_ROWS = 1_048_576, SHEET_COLS = 16_384

const unsupported = (feature: string, message: string): EngineResult<never> => ({
  ok: false, diagnostic: { kind: 'unsupported', feature, message } satisfies PivotDiagnostic,
})

/** Compose the Stage 2 library engine. `date1904` is explicit because the cache definition has no date-system field. */
export function computePivot(store: RecordStore, cache: PivotCacheDefinition, def: PivotTableDefinition,
  view: PivotViewState, options: PlaceOptions): EngineResult<PivotLayout> {
  if (def.unsupported.length) return unsupported('pivot-engine-pending', `unsupported pivot definition: ${def.unsupported.join(', ')}`)
  if (cache.olap || cache.source.type !== 'worksheet') return unsupported('pivot-engine-pending', 'OLAP and non-worksheet sources are not computed')
  // Date-filter criteria are read on the 1900 epoch (filters.ts) while records carry the workbook's serials.
  if (options.date1904 && view.filters.some(filter => filter.type.startsWith('date')))
    return unsupported('pivot-engine-unmeasured', 'date-filter criteria in a 1904 workbook are not measured (the oracle is 1900-based)')
  const axes = computeAxes(store, cache, def, view)
  if (!axes.ok) return axes
  const values = computeValues(store, cache, def, view, axes.value)
  if (!values.ok) return values
  // Page fields that would sit above row 1 would be dropped while their selections still filter the records.
  if (!pageFieldsFit(def, view)) return unsupported('pivot-values-limit', 'the page fields do not fit above the pivot location')
  const layout = placeCells(def, view, cache, axes.value, values.value, options)
  // The axis caps don't count the anchor or the header rows: a pivot near the sheet's edge can still outgrow it.
  const outside = (row: number, col: number) => row < 0 || col < 0 || row >= SHEET_ROWS || col >= SHEET_COLS
  // Page cells sit outside `rect` (above it, wrapping across), so every placed cell is checked too.
  if (outside(layout.rect.minRow, layout.rect.minCol) || outside(layout.rect.maxRow, layout.rect.maxCol) ||
    layout.cells.some(cell => outside(cell.row, cell.col)))
    return unsupported('pivot-values-limit', 'the pivot layout would extend beyond the worksheet')
  return { ok: true, value: layout }
}
