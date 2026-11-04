import { afterEach, expect, test, vi } from 'vitest'
import { cleanup, render, waitFor } from '@testing-library/react'
import InteractiveSheetDoc from '../src/components/xlsx/InteractiveSheetDoc'
import * as sources from '../src/xlsx/view/source'
import * as controls from '../src/components/xlsx/viewerControlsClient'
import * as filters from '../src/xlsx/view/filter'
import { DisplayTextLaneCache } from '../src/worker/displayTextLane'
import { semanticsFor } from '../src/xlsx/view/semanticsFor'

afterEach(() => { cleanup(); vi.restoreAllMocks() })

for (const profile of ['D', 'M'] as const) test(`local-controls-header-ready-accounting-fixed ${profile} first/reopen has no eager columns or display builds`, async () => {
  const styles = [{ numFmtId: 0 }, { numFmtId: 0, formatCode: '"' + 'z'.repeat(3000) + '"0' }]
  const n = 4096
  const doc: any = { sheets: [{ name: 'S', sheetId: 's', workbookIndex: 0,
    rows: Array.from({ length: n + 1 }, (_, row) => ({ index: row, cells: Array.from({ length: 3 }, (_, col) => ({ row, col, ref: '', value: row === 0 ? 'H' + col : col === 2 && row === n ? 'label' : col === 2 ? 0 : n - row, styleIndex: col === 2 ? 1 : 0, style: styles[col === 2 ? 1 : 0] })) })),
    cols: [], merges: [], mergeRanges: [], semantics: semanticsFor('1900') }], tables: [] }
  let columns = 0, structural = 0
  const make = sources.createObjectWorkbookSource, open = controls.viewerControls
  vi.spyOn(sources, 'createObjectWorkbookSource').mockImplementation((...args) => {
    const source = make(...args), store = source.store('s'), query = store.queryColumn.bind(store), structure = store.queryStructural.bind(store)
    store.queryColumn = async (...request) => { columns++; return query(...request) }
    store.queryStructural = async (...request) => { structural++; return structure(...request) }
    return source
  })
  const classify = vi.spyOn(filters, 'classifyColumn'), lease = vi.spyOn(DisplayTextLaneCache.prototype, 'lease')
  let resources: controls.ViewerControlResources | undefined
  vi.spyOn(controls, 'viewerControls').mockImplementation(async (...args) => { resources = args[0].resources; return open(...args) })
  for (let attempt = 0; attempt < 2; attempt++) {
    const ui = render(<InteractiveSheetDoc doc={doc} profile={profile} />)
    await waitFor(() => expect((ui.getByRole('button', { name: /^Filter column A/ }) as HTMLButtonElement).disabled).toBe(false))
    expect(columns).toBe(0)
    expect(structural).toBe(attempt + 1)
    expect(classify).not.toHaveBeenCalled()
    expect(lease).not.toHaveBeenCalled()
    expect(resources!.cache.size).toBe(0)
    expect(resources!.cache.usedBytes).toBe(0)
    expect(resources!.memory.ledger.peakBytes).toBeLessThanOrEqual(resources!.memory.limitBytes)
    ui.unmount()
  }
}, 30000)
