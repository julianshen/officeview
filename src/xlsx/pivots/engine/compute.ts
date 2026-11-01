/** Complete Stage 2 pivot computation: axes, values and placement. */
import { PIVOT_VALUES_FIELD, type PivotCacheDefinition, type PivotDiagnostic, type PivotTableDefinition } from '../model'
import { computeAxes, subtotalTypes, type EngineResult } from './axes'
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

function axisExpansionUpperBytes(def: PivotTableDefinition, cache: PivotCacheDefinition, view: PivotViewState,
  fields: readonly number[], lineCap: number, which: 'rows' | 'cols'): number {
  if (!fields.some(field => field !== PIVOT_VALUES_FIELD && view.fields[field]?.showAll)) return 0
  const dataCount = Math.max(1, view.data.length)
  const valuesLevel = fields.indexOf(PIVOT_VALUES_FIELD)
  let prefixes = 1, depth = 0, lines = 0, itemSlots = 0
  const add = (a: number, b: number) => Math.min(lineCap, a + b)
  const multiply = (a: number, b: number) => Math.min(lineCap, a * b)
  for (let level = 0; level < fields.length; level++) {
    const field = fields[level]!
    if (field === PIVOT_VALUES_FIELD) {
      prefixes = multiply(prefixes, dataCount)
      depth++
      if (which === 'rows' && def.outlineData && level < fields.length - 1) {
        lines = add(lines, prefixes)
        itemSlots += prefixes * depth
      }
      continue
    }
    const pivotItems = def.fields[field]?.items
    const itemCount = pivotItems ? pivotItems.reduce((count, item) => count + (item.type === 'data' ? 1 : 0), 0)
      : cache.fields[field]?.fieldGroup?.groupItems?.length ?? cache.fields[field]?.sharedItems?.items.length ?? 0
    const state = view.fields[field]
    let hidden = 0
    if (state) for (let item = 0; item < Math.min(itemCount, state.hidden.length); item++) if (state.hidden[item]) hidden++
    prefixes = multiply(prefixes, Math.max(0, itemCount - hidden))
    depth++
    if (!state || prefixes === 0) continue
    const valuesBelow = valuesLevel > level
    const subtotalCount = subtotalTypes(state).length * (valuesBelow ? dataCount : 1)
    const extraLines = (which === 'rows' && state.outline ? 1 : 0) + subtotalCount
      + (state.insertBlankRow ? 1 : 0) + (state.collapsed.some(Boolean) ? dataCount + 1 : 0)
    if (extraLines) {
      const count = multiply(prefixes, extraLines)
      lines = add(lines, count)
      itemSlots += count * depth
    }
  }
  // Every surviving leaf contributes at least one line. `lineCap` is the sheet axis limit used by computeAxes.
  lines = add(lines, prefixes)
  itemSlots += prefixes * depth
  // Reserve the retained line objects and path arrays plus headroom for array growth/transient construction.
  const bytes = 2 * (lines * 64 + itemSlots * 8) + 1024
  return Number.isSafeInteger(bytes) ? bytes : Number.POSITIVE_INFINITY
}

/** Compose the Stage 2 library engine. `date1904` is explicit because the cache definition has no date-system field. */
export function computePivot(store: RecordStore, cache: PivotCacheDefinition, def: PivotTableDefinition,
  view: PivotViewState, options: PlaceOptions & { admitLayoutBytes?: (bytes: number) => PivotCandidateHold | undefined }): EngineResult<PivotLayout> {
  if (def.unsupported.length) return unsupported('pivot-engine-pending', `unsupported pivot definition: ${def.unsupported.join(', ')}`)
  if (cache.olap || cache.source.type !== 'worksheet') return unsupported('pivot-engine-pending', 'OLAP and non-worksheet sources are not computed')
  // Date-filter criteria are read on the 1900 epoch (filters.ts) while records carry the workbook's serials.
  if (options.date1904 && view.filters.some(filter => filter.type.startsWith('date')))
    return unsupported('pivot-engine-unmeasured', 'date-filter criteria in a 1904 workbook are not measured (the oracle is 1900-based)')
  // Cover the record-key maps, per-record axis keys and hierarchy paths before constructing axes. Their exact line
  // counts are not known until the scan. Show All can fan out beyond records, so add a metadata-only upper bound for
  // those output lines and their hierarchy path arrays before computeAxes builds them.
  const axisDepth = Math.max(1, view.rows.length + view.cols.length)
  const showAllAxisUpper = axisExpansionUpperBytes(def, cache, view, view.rows, SHEET_ROWS, 'rows')
    + axisExpansionUpperBytes(def, cache, view, view.cols, SHEET_COLS, 'cols')
  const axisUpper = store.count * (64 + axisDepth * 64) + cache.fields.length * 128 + showAllAxisUpper
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
  let itemSlots = 0
  for (const line of axes.value.rows) itemSlots += line.items.length
  for (const line of axes.value.cols) itemSlots += line.items.length
  const upperBytes = store.count * 8 + product * 256 + (rowCount + colCount) * 256 + itemSlots * 16 + 2048
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
