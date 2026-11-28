/** The sort guard's outside-column scan (`guardSort` → `findOutsideColumnData` → `DocumentSession.populatedSource`)
 * settles every `readRows` page it reads: a page's main-thread charge ends when the scan is done with it, not when
 * the collector happens to reclaim it. A one-column sort over a full 1,048,576-row sheet then stays within the
 * main-thread ledger. The ledger's collection watcher never reports a collection, so only explicit settlement ends a
 * page's charge, whenever the host's collector runs. */
import { expect, test } from 'vitest'
import { viewerControls } from '../src/components/xlsx/viewerControlsClient'
import type { CacheIdentity, SheetView } from '../src/xlsx/view/contracts'
import { MainMemoryLedger } from '../src/xlsx/view/mainLedger'
import { createIdentityOrder } from '../src/xlsx/view/order'
import { emptyGridSelection } from '../src/xlsx/view/selection'
import { semanticsFor } from '../src/xlsx/view/semanticsFor'
import { createObjectWorkbookSource } from '../src/xlsx/view/source'
import { MAIN_THREAD_BUDGETS } from '../src/xlsx/view/tileCache'
import { createIntervalSet, visibilityFromHidden } from '../src/xlsx/view/visibility'
import type { XlsxDocument } from '../src/xlsx/types'

const ROWS = 1_048_576

test('a full-column sort guard over 1,048,576 rows stays within the main-thread ledger', async () => {
  const rows = Array.from({ length: ROWS }, (_, r) => ({ index: r, cells: [{ row: r, col: 0, ref: '', value: ROWS - r, styleIndex: 0, style: { numFmtId: 0 } }] }))
  const doc = { sheets: [{ name: 'S', sheetId: 's', workbookIndex: 0, rows, cols: [], merges: [], mergeRanges: [], semantics: semanticsFor('1900') }], tables: [] } as unknown as XlsxDocument
  const source = createObjectWorkbookSource(doc, 'd')
  const store = source.store('s')
  const ledger = new MainMemoryLedger({ register() {} })
  let peak = 0
  const set = ledger.set.bind(ledger)
  ledger.set = (pool, owner, bytes) => { set(pool, owner, bytes); peak = Math.max(peak, ledger.bytes) }
  const hidden = createIntervalSet([])
  let viewGen = 0
  let view: SheetView = { sheetId: 's', revision: 0, order: createIdentityOrder(), baseHidden: hidden, visibility: visibilityFromHidden(hidden), sort: null, filter: null, selection: emptyGridSelection() }
  const owner = (): CacheIdentity => ({ ...store.stamp, fontGen: 0, viewGen })
  const snapshot = () => ({ owner: owner(), filter: view.filter, sort: view.sort, records: { visible: ROWS, total: ROWS, completeness: 'complete' as const } })
  const publish = (next: Pick<SheetView, 'order' | 'visibility' | 'filter' | 'sort'>) => { view = { ...view, ...next }; viewGen++; return snapshot() }
  const signal = new AbortController().signal
  const controls = await viewerControls({ doc, source, sheetId: 's', snapshot, view: () => view, publish, ledger }, signal)
  try {
    const result = await controls.client.sort(owner(), { range: { r0: 0, r1: ROWS - 1, c0: 0, c1: 0, headerRows: 0, totalsRows: 0 }, keys: [{ col: 0, order: 'asc' }], caseSensitive: false }, signal)
    expect(result.kind).toBe('committed')
    expect(view.order.bands[0]!.sourceAtOffset[0]).toBe(ROWS - 1)
    expect(peak).toBeLessThanOrEqual(MAIN_THREAD_BUDGETS.D)
  } finally {
    controls.dispose()
    await source.release()
  }
}, 120_000)
