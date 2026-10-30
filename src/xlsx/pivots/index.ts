/** Entry of the lazy pivots chunk: definitions (loaded at parse time) and their display, registered with the eager pivotLoader. */
import { registerPivotEngine } from '../pivotLoader'
import { pivotStyleRegions } from './regions'
import { paintPivotButtons } from './buttons'
import { loadSheetControls } from './slicers/load'
import { paintControl } from './slicers/paint'

export { loadPivotWorkbook } from './load'
// The Stage 2 engine (library only: no runtime caller yet). Exported here so it lives in this lazy chunk.
export { decodeRecords, isoToSerial, recordEvents, RECORD_LIMITS, VALUE_TAG } from './engine/records'
export type { RecordColumn, RecordLimits, RecordResult, RecordStore } from './engine/records'
export { initialViewState } from './engine/view'
export type { DataFieldState, FieldState, PageState, PivotViewState } from './engine/view'
export { computeAxes, lineCarriesValues, NO_LINE, subtotalTypes } from './engine/axes'
export type { EngineResult, PivotAxes } from './engine/axes'
export { liveDefinition, placeCells } from './engine/place'
export type { PivotCellValue, PivotLayout, PivotLayoutCell, PivotValueGrid, PlaceOptions } from './engine/place'

registerPivotEngine({ regions: pivotStyleRegions, paintButtons: paintPivotButtons, sheetControls: loadSheetControls, paintControl })
