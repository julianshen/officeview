/** Complete Stage 2 pivot computation: axes, values and placement. */
import type { PivotCacheDefinition, PivotDiagnostic, PivotTableDefinition } from '../model'
import { computeAxes, type EngineResult } from './axes'
import type { PivotLayout, PlaceOptions } from './place'
import { placeCells } from './place'
import { computeValues } from './values'
import type { RecordStore } from './records'
import type { PivotViewState } from './view'

const unsupported = (feature: string, message: string): EngineResult<never> => ({
  ok: false, diagnostic: { kind: 'unsupported', feature, message } satisfies PivotDiagnostic,
})

/** Compose the Stage 2 library engine. `date1904` is explicit because the cache definition has no date-system field. */
export function computePivot(store: RecordStore, cache: PivotCacheDefinition, def: PivotTableDefinition,
  view: PivotViewState, options: PlaceOptions): EngineResult<PivotLayout> {
  if (def.unsupported.length) return unsupported('pivot-engine-pending', `unsupported pivot definition: ${def.unsupported.join(', ')}`)
  if (cache.olap || cache.source.type !== 'worksheet') return unsupported('pivot-engine-pending', 'OLAP and non-worksheet sources are not computed')
  const axes = computeAxes(store, cache, def, view)
  if (!axes.ok) return axes
  const values = computeValues(store, cache, def, view, axes.value)
  if (!values.ok) return values
  return { ok: true, value: placeCells(def, view, cache, axes.value, values.value, options) }
}
