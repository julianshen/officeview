/** Complete Stage 2 pivot computation: axes, values and placement. */
import type { PivotCacheDefinition, PivotDiagnostic, PivotTableDefinition } from '../model'
import { computeAxes, type EngineResult } from './axes'
import type { PivotCandidateHold, PivotLayout, PlaceOptions } from './place'
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
  view: PivotViewState, options: PlaceOptions & { admitLayoutBytes?: (bytes: number) => PivotCandidateHold | undefined }): EngineResult<PivotLayout> {
  if (def.unsupported.length) return unsupported('pivot-engine-pending', `unsupported pivot definition: ${def.unsupported.join(', ')}`)
  if (cache.olap || cache.source.type !== 'worksheet') return unsupported('pivot-engine-pending', 'OLAP and non-worksheet sources are not computed')
  // Date-filter criteria are read on the 1900 epoch (filters.ts) while records carry the workbook's serials.
  if (options.date1904 && view.filters.some(filter => filter.type.startsWith('date')))
    return unsupported('pivot-engine-unmeasured', 'date-filter criteria in a 1904 workbook are not measured (the oracle is 1900-based)')
  // Cover the record-key maps, per-record axis keys and hierarchy paths before constructing axes. Their exact line
  // counts are not known until the scan; the later resize adds the dense layout upper bound before values/placement.
  const axisDepth = Math.max(1, view.rows.length + view.cols.length)
  const axisUpper = store.count * (64 + axisDepth * 64) + cache.fields.length * 128
    + (view.rows.length + view.cols.length + view.pages.length + view.data.length) * 512 + 4096
  if (!Number.isSafeInteger(axisUpper)) return unsupported('pivot-records-limit', 'The candidate layout exceeds the memory budget.')
  const candidateHold = options.admitLayoutBytes?.(axisUpper)
  if (options.admitLayoutBytes && !candidateHold) return unsupported('pivot-records-limit', 'The candidate layout exceeds the memory budget.')
  const releaseCandidate = () => candidateHold?.release()
  let axes: ReturnType<typeof computeAxes>
  try { axes = computeAxes(store, cache, def, view) }
  catch (error) { releaseCandidate(); throw error }
  if (!axes.ok) { releaseCandidate(); return axes }
  // The axes give a no-allocation upper bound for the dense values and placed cells. Reserve it before either is built.
  const rowCount = axes.value.rows.length, colCount = axes.value.cols.length
  const product = rowCount * colCount
  const upperBytes = store.count * 8 + product * 256 + (rowCount + colCount) * 256 + 2048
  if (!Number.isSafeInteger(upperBytes)) { releaseCandidate(); return unsupported('pivot-records-limit', 'The candidate layout exceeds the memory budget.') }
  try {
    if (candidateHold && upperBytes > candidateHold.bytes && candidateHold.resize && !candidateHold.resize(upperBytes)) {
      releaseCandidate()
      return unsupported('pivot-records-limit', 'The candidate layout exceeds the memory budget.')
    }
  } catch (error) {
    releaseCandidate()
    throw error
  }
  let values: ReturnType<typeof computeValues>
  try { values = computeValues(store, cache, def, view, axes.value) }
  catch (error) { releaseCandidate(); throw error }
  if (!values.ok) { releaseCandidate(); return values }
  // Page fields that would sit above row 1 would be dropped while their selections still filter the records.
  if (!pageFieldsFit(def, view)) { releaseCandidate(); return unsupported('pivot-values-limit', 'the page fields do not fit above the pivot location') }
  let layout: PivotLayout
  try { layout = placeCells(def, view, cache, axes.value, values.value, options) }
  catch (error) { releaseCandidate(); throw error }
  // The axis caps don't count the anchor or the header rows: a pivot near the sheet's edge can still outgrow it.
  const outside = (row: number, col: number) => row < 0 || col < 0 || row >= SHEET_ROWS || col >= SHEET_COLS
  // Page cells sit outside `rect` (above it, wrapping across), so every placed cell is checked too.
  if (outside(layout.rect.minRow, layout.rect.minCol) || outside(layout.rect.maxRow, layout.rect.maxCol) ||
    layout.cells.some(cell => outside(cell.row, cell.col)))
    { releaseCandidate(); return unsupported('pivot-values-limit', 'the pivot layout would extend beyond the worksheet') }
  if (candidateHold) layout = { ...layout, candidateHold }
  return { ok: true, value: layout }
}
