/** Entry of the lazy pivots chunk: definitions (loaded at parse time) and their display, registered with the eager pivotLoader. */
import { registerPivotEngine } from '../pivotLoader'
import { pivotStyleRegions } from './regions'
import { paintPivotButtons } from './buttons'
import { loadSheetControls } from './slicers/load'
import { paintControl } from './slicers/paint'

export { loadPivotWorkbook } from './load'

registerPivotEngine({ regions: pivotStyleRegions, paintButtons: paintPivotButtons, sheetControls: loadSheetControls, paintControl })
