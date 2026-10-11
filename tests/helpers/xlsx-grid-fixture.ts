import type { GridSemanticsProps } from '../../src/components/xlsx/GridSemantics'
import { createAxisMetrics } from '../../src/xlsx/view/axes'
import { createSheetMetrics } from '../../src/xlsx/view/projection'
import { createRowOrder } from '../../src/xlsx/view/order'
import { createIntervalSet, visibilityFromHidden } from '../../src/xlsx/view/visibility'
import { MAX_ROWS, MAX_COLS } from '../../src/xlsx/view/coordinates'
import { createScrollWindow } from '../../src/xlsx/view/scrollWindow'
import { emptyGridSelection, selectCell } from '../../src/xlsx/view/selection'

export function gridFixture(overrides: Partial<GridSemanticsProps> = {}): GridSemanticsProps {
  const hidden = createIntervalSet([{ r0: 2, r1: 2 }]), visibility = visibilityFromHidden(hidden)
  const projection = createSheetMetrics({ rows: createAxisMetrics({ logicalCount: MAX_ROWS, defaultSizeCss: 24 }), cols: createAxisMetrics({ logicalCount: MAX_COLS, defaultSizeCss: 80, collapsed: [{ start: 1, end: 1 }] }), order: createRowOrder([{ r0: 0, r1: 3, sourceAtOffset: Uint32Array.of(0, 3, 1, 2) }]), visibility })
  const viewport = { width: 240, height: 96 }, selection = selectCell(emptyGridSelection(), projection, { row: 3, col: 0 })
  return { owner: { docId: 'doc', sheetId: 's', valueGen: 1, fontGen: 1, viewGen: 1 }, projection,
    view: { baseHidden: createIntervalSet([]), visibility, filter: null, sort: null }, selection,
    window: createScrollWindow(projection, viewport), viewport, body: { r0: 1, r1: 3 }, onChange: () => {}, cellPresentation: () => ({ text: '10' }), ...overrides }
}

