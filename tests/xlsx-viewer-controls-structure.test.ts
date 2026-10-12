/**
 * S10.1 reproduced gap (real Chromium, worker-backed F2): the viewer controls
 * port snapshots the sheet's structural metadata once, when the sheet opens. A
 * paged workbook is still importing at that moment, so the worker answers
 * `completeness: 'loading'`, and every later header sort/filter refused with
 * "Source range is incomplete." forever — no worker command was ever issued
 * again (STORE-04 / LIFE-03: a loading state must resolve, not stick).
 */
import { describe, expect, it } from 'vitest'
import { viewerControls } from '../src/components/xlsx/viewerControlsClient'
import { createObjectWorkbookSource } from '../src/xlsx/view/source'
import { createIdentityOrder } from '../src/xlsx/view/order'
import { createIntervalSet, visibilityFromHidden } from '../src/xlsx/view/visibility'
import { emptyGridSelection } from '../src/xlsx/view/selection'
import type { SheetView, StructuralResult } from '../src/xlsx/view/contracts'
import type { XlsxDocument } from '../src/xlsx/types'

const num = (row: number, col: number, value: number | string) => ({ ref: '', row, col, value, styleIndex: 0, style: undefined } as never)
const doc: XlsxDocument = {
  sheets: [{
    name: 'S', sheetId: '1', workbookIndex: 0, cols: [], merges: [], mergeRanges: [],
    rows: [
      { index: 0, cells: [num(0, 0, 'n')] },
      { index: 1, cells: [num(1, 0, 3)] },
      { index: 2, cells: [num(2, 0, 1)] },
      { index: 3, cells: [num(3, 0, 2)] },
    ],
  } as never],
  tables: [],
}

function loadingFirst() {
  const source = createObjectWorkbookSource(doc, 'doc-1')
  const store = source.store('1')
  let calls = 0
  const queryStructural = store.queryStructural.bind(store)
  // The first answer models a worker import still in flight; later ones are complete.
  store.queryStructural = async (rect, options): Promise<StructuralResult> => {
    calls++
    const real = await queryStructural(rect, options)
    return calls === 1 ? { ...real, completeness: 'loading', missing: [rect] } : real
  }
  return { source, calls: () => calls }
}

describe('viewer controls structural snapshot', () => {
  it('a structure that was still loading when the sheet opened is re-queried before a sort resolves', async () => {
    const { source, calls } = loadingFirst()
    const order = createIdentityOrder()
    const view: SheetView = { sheetId: '1', revision: 0, order, baseHidden: createIntervalSet([]), visibility: visibilityFromHidden(createIntervalSet([])), sort: null, filter: null, selection: emptyGridSelection() }
    const owner = { docId: 'doc-1', sheetId: '1', valueGen: 0, fontGen: 0, viewGen: 0 }
    const snapshot = () => ({ owner, filter: null, sort: null, records: { visible: 3, total: 3, completeness: 'complete' as const } })
    const controls = await viewerControls({ doc, source, sheetId: '1', snapshot, view: () => view, publish: snapshot }, new AbortController().signal)
    const resolved = await controls.client.resolve(owner, 'sort', new AbortController().signal, undefined, 0)
    expect(resolved).toMatchObject({ kind: 'complete', range: { r0: 0, r1: 3, c0: 0, c1: 0, headerRows: 1 } })
    expect(calls()).toBe(2)
    // Once complete, the snapshot is reused: no structural query per gesture.
    await controls.client.resolve(owner, 'sort', new AbortController().signal, undefined, 0)
    expect(calls()).toBe(2)
    controls.dispose()
  })

  /**
   * S10.1 reproduced gap (Chromium/WebKit, F2 resident and worker-backed): after
   * Go To A1, the funnel on column D refused "Contiguous selection does not
   * contain the active cell". A header gesture moves the active cell to its
   * column but kept the old single-cell selection, which detection (S7.1,
   * deliberately) treats as an explicit selection away from the active cell.
   */
  it('a header gesture on another column ignores a leftover single-cell selection', async () => {
    const wide: XlsxDocument = { ...doc, sheets: [{ ...doc.sheets[0], rows: [0, 1, 2, 3].map(r => ({ index: r, cells: [0, 1, 2].map(c => num(r, c, r === 0 ? `h${c}` : r * 10 + c)) })) }] }
    const source = createObjectWorkbookSource(wide, 'doc-1')
    const order = createIdentityOrder()
    const a1 = { rows: { kind: 'intervals' as const, intervals: [{ r0: 0, r1: 0 }] }, cols: [{ c0: 0, c1: 0 }] }
    const view: SheetView = { sheetId: '1', revision: 0, order, baseHidden: createIntervalSet([]), visibility: visibilityFromHidden(createIntervalSet([])), sort: null, filter: null,
      selection: { active: { row: 0, col: 0 }, anchor: { row: 0, col: 0 }, ranges: [a1] } }
    const owner = { docId: 'doc-1', sheetId: '1', valueGen: 0, fontGen: 0, viewGen: 0 }
    const snapshot = () => ({ owner, filter: null, sort: null, records: { visible: 3, total: 3, completeness: 'complete' as const } })
    const controls = await viewerControls({ doc: wide, source, sheetId: '1', snapshot, view: () => view, publish: snapshot }, new AbortController().signal)
    const resolved = await controls.client.resolve(owner, 'filter', new AbortController().signal, undefined, 2)
    expect(resolved).toMatchObject({ kind: 'complete', range: { r0: 0, r1: 3, c0: 0, c1: 2, headerRows: 1 } })
    controls.dispose()
  })
})
