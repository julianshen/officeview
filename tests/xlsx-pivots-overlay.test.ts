import { describe, expect, it } from 'vitest'
import { createObjectWorkbookSource } from '../src/xlsx/view/source'
import { buildSheetGeometry } from '../src/xlsx/view/viewport'
import { createPagedWorkbookSource } from '../src/xlsx/view/pagedSource'
import { openImportSession } from '../src/xlsx/storage/pagedImport'
import { SourcePageStore, memoryPageBackend } from '../src/xlsx/storage/sourcePages'
import { xlsxBytes } from './helpers/xlsx-package'
import { StructuralIndex, fromStoreScan } from '../src/xlsx/view/structure'
import { guardFilter, guardSort } from '../src/xlsx/view/operationGuard'
import { invalidateRegionStyles } from '../src/xlsx/tableStyleLoader'
import { pagedSheetGeometry } from '../src/components/xlsx/pagedGeometry'
import { createAxisMetrics } from '../src/xlsx/view/axes'
import { createSheetMetrics } from '../src/xlsx/view/projection'
import { createIdentityOrder } from '../src/xlsx/view/order'
import { createIntervalSet, visibilityFromHidden } from '../src/xlsx/view/visibility'
import { MAX_COLS, MAX_ROWS } from '../src/xlsx/view/coordinates'
import { drain, findInputOf, workbookOf } from './helpers/xlsx-find'
import { WorkbookFind } from '../src/xlsx/view/find'
import { serializeSelection } from '../src/xlsx/view/copy'
import { initialWorkbookOwner } from '../src/worker/viewerProtocol'
import { interruptibleStore } from '../src/worker/viewerAbort'
import { createCanvas } from 'canvas'
import { renderSheetViewport } from '../src/xlsx/view/renderViewport'
import { LayoutCache } from '../src/xlsx/view/textLayout'
import { SOURCE_ALLOCATION } from '../src/xlsx/view/allocationAccount'
import type { ValueOverlay } from '../src/xlsx/view/contracts'
import { RemoteWorkbook } from '../src/worker/viewerRemoteSource'

const options = (store: any) => ({ signal: new AbortController().signal, stamp: { ...store.stamp } })
const overlay = {
  rects: [{ r0: 0, r1: 2, c0: 0, c1: 1 }], protectedRects: [{ r0: 0, r1: 2, c0: 0, c1: 1 }], merges: [],
  cell(row: number, col: number) {
    if (row === 0 && col === 0) return { value: 100, numFmtId: 14 }
    if (row === 2 && col === 1) return { value: 'new', numFmtId: 0 }
    return null
  },
}

describe('pivot value overlay', () => {
  it('does not repeat overlay rows across resident and paged row cursors', async () => {
    const rows = Array.from({ length: 405 }, (_, row) => ({ row, col: 8, value: `base-${row}` }))
    const doc = workbookOf([{ sheetId: 's', cells: rows }])
    const source = createObjectWorkbookSource(doc, 'overlay-pages')
    const store: any = source.store('s')
    store.setValueOverlay({ rects: [{ r0: 0, r1: 2, c0: 0, c1: 0 }], protectedRects: [{ r0: 0, r1: 2, c0: 0, c1: 0 }], merges: [],
      cell: (row: number, col: number) => col === 0 && row < 3 ? { value: `over-${row}` } : undefined })
    const seen: number[] = []
    let cursor: string | undefined
    do {
      const page: any = await store.readRows({ r0: 0, r1: 404, c0: 0, c1: 8 }, cursor, options(store))
      seen.push(...page.items.map((item: any) => item.row))
      cursor = page.nextCursor
      expect(page.error).toBeUndefined()
    } while (cursor)
    expect(seen).toHaveLength(405)
    expect(seen.filter(row => row < 3)).toEqual([0, 1, 2])
    const find = findInputOf(doc, { source, docId: 'overlay-pages' })
    const baseOwner = initialWorkbookOwner('overlay-pages', source.sheets)
    const owner = { ...baseOwner, sheetVector: baseOwner.sheetVector.map(entry => entry.sheetId === 's' ? { ...entry, valueGen: store.stamp.valueGen } : entry) }
    const search = new WorkbookFind({ ...find.input, owner, sheets: [{ ...find.input.sheets[0], identity: { ...find.input.sheets[0].identity, valueGen: store.stamp.valueGen } }] }, { text: 'over-', caseSensitive: false })
    expect((await drain(search)).map(hit => [hit.source.row, hit.text])).toEqual([[0, 'over-0'], [1, 'over-1'], [2, 'over-2']])
    const copyOwner = { ...store.stamp, fontGen: 0, viewGen: 0 }
    const copied = await serializeSelection({ store, owner: copyOwner, currentOwner: () => copyOwner,
      selection: { active: null, anchor: null, ranges: [{ rows: { kind: 'intervals', intervals: [{ r0: 0, r1: 404 }] }, cols: [{ c0: 0, c1: 8 }] }] },
      order: createIdentityOrder(), visibility: visibilityFromHidden(createIntervalSet([])), columns: visibilityFromHidden(createIntervalSet([]), MAX_COLS),
      allowed: () => true, signal: new AbortController().signal, format: cell => source.displayText!(cell, 's') })
    expect(copied.kind).toBe('ready')
    if (copied.kind === 'ready') {
      expect(copied.text.split('\n')).toHaveLength(405)
      expect(copied.text.match(/over-[012]/g)).toHaveLength(3)
    }
    await source.release()
  })

  it('bounds dense overlay-only pages by row and cell limits', async () => {
    const rows = 20_000, cols = 14
    const doc = workbookOf([{ sheetId: 's', cells: [] }])
    const source = createObjectWorkbookSource(doc, 'overlay-big')
    const store: any = source.store('s')
    store.setValueOverlay({ rects: [{ r0: 0, r1: rows - 1, c0: 0, c1: cols - 1 }], protectedRects: [{ r0: 0, r1: rows - 1, c0: 0, c1: cols - 1 }], merges: [],
      cell: (row: number, col: number) => ({ value: row * cols + col }) })
    let count = 0, pages = 0, cursor: string | undefined
    do {
      const page: any = await store.readRows({ r0: 0, r1: rows - 1, c0: 0, c1: cols - 1 }, cursor, options(store))
      expect(page.error).toBeUndefined()
      expect(page.items.length).toBeLessThanOrEqual(200)
      const cells = page.items.reduce((sum: number, row: any) => sum + row.cells.length, 0)
      expect(cells).toBeLessThanOrEqual(262_144)
      count += cells; pages++; cursor = page.nextCursor
    } while (cursor)
    expect(count).toBe(rows * cols)
    expect(pages).toBeGreaterThan(1)
    let scanned = 0
    for await (const page of store.scanPopulated({ range: { r0: 0, r1: rows - 1, c0: 0, c1: cols - 1 }, limit: 5000 }, options(store))) {
      expect(page.error).toBeUndefined()
      scanned += page.items.length
    }
    expect(scanned).toBe(rows * cols)
    const find = findInputOf(doc, { source, docId: 'overlay-big' })
    const baseOwner = initialWorkbookOwner('overlay-big', source.sheets)
    const owner = { ...baseOwner, sheetVector: baseOwner.sheetVector.map(entry => entry.sheetId === 's' ? { ...entry, valueGen: store.stamp.valueGen } : entry) }
    const found = await drain(new WorkbookFind({ ...find.input, owner, sheets: [{ ...find.input.sheets[0], identity: { ...find.input.sheets[0].identity, valueGen: store.stamp.valueGen } }] }, { text: '279999', caseSensitive: false }))
    expect(found.map(hit => [hit.source.row, hit.source.col])).toEqual([[19_999, 13]])
    const copyOwner = { ...store.stamp, fontGen: 0, viewGen: 0 }
    const copied = await serializeSelection({ store, owner: copyOwner, currentOwner: () => copyOwner,
      selection: { active: null, anchor: null, ranges: [{ rows: { kind: 'intervals', intervals: [{ r0: 0, r1: rows - 1 }] }, cols: [{ c0: 0, c1: cols - 1 }] }] },
      order: createIdentityOrder(), visibility: visibilityFromHidden(createIntervalSet([])), columns: visibilityFromHidden(createIntervalSet([]), MAX_COLS),
      allowed: () => true, signal: new AbortController().signal, format: cell => source.displayText!(cell, 's') })
    expect(copied.kind).toBe('ready')
    if (copied.kind === 'ready') expect(copied.text.split('\n')).toHaveLength(rows)
    await source.release()
  })

  it('bounds paged overlay rows and resumes after the cursor without duplicates', async () => {
    const xml = Array.from({ length: 405 }, (_, i) => `<row r="${i + 1}"><c r="I${i + 1}"><v>${i + 1}</v></c></row>`).join('')
    const bytes = xlsxBytes({ sheetData: xml })
    const pages = new SourcePageStore({ docId: 'overlay-paged-pages', backend: memoryPageBackend() })
    const session = await openImportSession({ kind: 'blob', blob: new Blob([bytes.slice().buffer as ArrayBuffer]) }, { docId: 'overlay-paged-pages', profile: 'D', store: pages, pageRows: 500 })
    expect((await session.done).kind).toBe('complete')
    const source = createPagedWorkbookSource(session)
    const store: any = source.store('1')
    store.setValueOverlay({ rects: [{ r0: 0, r1: 2, c0: 0, c1: 0 }], protectedRects: [{ r0: 0, r1: 2, c0: 0, c1: 0 }], merges: [],
      cell: (row: number, col: number) => col === 0 && row < 3 ? { value: `over-${row}` } : undefined })
    let cursor: string | undefined, seen: number[] = []
    do {
      const page = await store.readRows({ r0: 0, r1: 404, c0: 0, c1: 8 }, cursor, options(store))
      expect(page.error).toBeUndefined()
      seen.push(...page.items.map((item: any) => item.row))
      cursor = page.nextCursor
    } while (cursor)
    expect(seen).toHaveLength(405)
    expect(seen).toEqual(Array.from({ length: 405 }, (_, row) => row))
    const pagedScan: Array<[number, number, unknown]> = []
    for await (const page of store.scanPopulated({ range: { r0: 0, r1: 404, c0: 0, c1: 8 }, limit: 64 }, options(store))) {
      pagedScan.push(...page.items.map((cell: any) => [cell.row, cell.col, cell.value]))
    }
    expect(pagedScan.filter(([, col]) => col === 0)).toEqual([[0, 0, 'over-0'], [1, 0, 'over-1'], [2, 0, 'over-2']])
    const baseOwner = initialWorkbookOwner('overlay-paged-pages', source.sheets)
    const owner = { ...baseOwner, sheetVector: baseOwner.sheetVector.map(entry => entry.sheetId === '1' ? { ...entry, valueGen: store.stamp.valueGen } : entry) }
    const visibility = visibilityFromHidden(createIntervalSet([]))
    const found = await drain(new WorkbookFind({ owner, sheets: [{ sheetId: '1', store,
      identity: { docId: store.stamp.docId, sheetId: '1', valueGen: store.stamp.valueGen, fontGen: owner.fontGen, viewGen: owner.sheetVector[0].revision },
      order: createIdentityOrder(), visibility }] }, { text: 'over-', caseSensitive: false }))
    expect(found.map(hit => [hit.source.row, hit.text])).toEqual([[0, 'over-0'], [1, 'over-1'], [2, 'over-2']])
    const copyOwner = { ...store.stamp, fontGen: 0, viewGen: 0 }
    const copied = await serializeSelection({ store, owner: copyOwner, currentOwner: () => copyOwner,
      selection: { active: null, anchor: null, ranges: [{ rows: { kind: 'intervals', intervals: [{ r0: 0, r1: 404 }] }, cols: [{ c0: 0, c1: 8 }] }] },
      order: createIdentityOrder(), visibility, columns: visibilityFromHidden(createIntervalSet([]), MAX_COLS), allowed: () => true,
      signal: new AbortController().signal, format: cell => source.displayText!(cell, '1') })
    expect(copied.kind).toBe('ready')
    if (copied.kind === 'ready') {
      expect(copied.text.split('\n')).toHaveLength(405)
      expect(copied.text.match(/over-[012]/g)).toHaveLength(3)
    }
    store.setValueOverlay({ rects: [{ r0: 0, r1: 19_999, c0: 0, c1: 13 }], protectedRects: [{ r0: 0, r1: 19_999, c0: 0, c1: 13 }], merges: [],
      cell: (row: number, col: number) => ({ value: row * 14 + col }) })
    let count = 0, pageCount = 0
    cursor = undefined
    do {
      const page: any = await store.readRows({ r0: 0, r1: 19_999, c0: 0, c1: 13 }, cursor, options(store))
      expect(page.error).toBeUndefined()
      expect(page.items.length).toBeLessThanOrEqual(200)
      const cells = page.items.reduce((sum: number, item: any) => sum + item.cells.length, 0)
      expect(cells).toBeLessThanOrEqual(262_144)
      count += cells; pageCount++; cursor = page.nextCursor
    } while (cursor)
    expect(count).toBe(280_000)
    expect(pageCount).toBeGreaterThan(1)
    let scanCount = 0
    for await (const page of store.scanPopulated({ range: { r0: 0, r1: 19_999, c0: 0, c1: 13 }, limit: 5000 }, options(store))) {
      expect(page.error).toBeUndefined()
      scanCount += page.items.length
    }
    expect(scanCount).toBe(280_000)
    {
      const largeBaseOwner = initialWorkbookOwner('overlay-paged-pages', source.sheets)
      const largeOwner = { ...largeBaseOwner, sheetVector: largeBaseOwner.sheetVector.map(entry => entry.sheetId === '1' ? { ...entry, valueGen: store.stamp.valueGen } : entry) }
      const largeVisibility = visibilityFromHidden(createIntervalSet([]))
      const largeFound = await drain(new WorkbookFind({ owner: largeOwner, sheets: [{ sheetId: '1', store,
        identity: { docId: store.stamp.docId, sheetId: '1', valueGen: store.stamp.valueGen, fontGen: largeOwner.fontGen, viewGen: largeOwner.sheetVector[0].revision },
        order: createIdentityOrder(), visibility: largeVisibility }] }, { text: '279999', caseSensitive: false }))
      expect(largeFound.map(hit => [hit.source.row, hit.source.col])).toEqual([[19_999, 13]])
      const largeCopyOwner = { ...store.stamp, fontGen: 0, viewGen: 0 }
      const largeCopy = await serializeSelection({ store, owner: largeCopyOwner, currentOwner: () => largeCopyOwner,
        selection: { active: null, anchor: null, ranges: [{ rows: { kind: 'intervals', intervals: [{ r0: 0, r1: 19_999 }] }, cols: [{ c0: 0, c1: 13 }] }] },
        order: createIdentityOrder(), visibility: largeVisibility, columns: visibilityFromHidden(createIntervalSet([]), MAX_COLS), allowed: () => true,
        signal: new AbortController().signal, format: cell => source.displayText!(cell, '1') })
      expect(largeCopy.kind).toBe('ready')
      if (largeCopy.kind === 'ready') expect(largeCopy.text.split('\n')).toHaveLength(20_000)
    }
    await source.release()
  })
  it('keeps the no-overlay path to a single undefined check', async () => {
    const doc = workbookOf([{ sheetId: 's', cells: [{ row: 0, col: 0, value: 1 }] }])
    const source = createObjectWorkbookSource(doc, 'no-overlay')
    const store: any = source.store('s')
    expect(Object.prototype.hasOwnProperty.call(store, 'valueOverlay')).toBe(false)
    expect(Object.prototype.hasOwnProperty.call(store, 'currentStructureGen')).toBe(false)
    let storeChecks = 0
    Object.defineProperty(store, 'valueOverlay', { configurable: true, get: () => { storeChecks++; return undefined } })
    expect(await store.readCell({ row: 0, col: 0 }, options(store))).toMatchObject({ kind: 'cell', cell: { value: 1 } })
    expect(storeChecks).toBe(1)
    const range = { r0: 0, r1: 0, c0: 0, c1: 0 }
    storeChecks = 0
    await store.readRows(range, undefined, options(store)); expect(storeChecks).toBe(1)
    storeChecks = 0
    await store.queryColumn({ col: 0, rows: { kind: 'intervals', intervals: [{ r0: 0, r1: 0 }] }, limit: 2 }, options(store)); expect(storeChecks).toBe(1)
    storeChecks = 0
    for await (const _page of store.scanPopulated({ range, limit: 2 }, options(store))) { /* drain */ }
    expect(storeChecks).toBe(1)
    storeChecks = 0
    for await (const _page of store.scanOccupied({ range, limit: 2 }, options(store))) { /* drain */ }
    expect(storeChecks).toBe(1)
    let geometryChecks = 0
    const geometry = buildSheetGeometry(doc.sheets[0], Object.defineProperty({}, 'overlay', { get: () => { geometryChecks++; return undefined } }))
    expect(geometry.kind).toBe('ready')
    expect(geometryChecks).toBe(1)
    const noPivotStructure = new StructuralIndex({
      stamp: store.stamp, completeness: 'complete', missing: [], merges: [], tables: [], spills: [], drawings: [],
    })
    expect(Object.prototype.hasOwnProperty.call(noPivotStructure, 'pivots')).toBe(false)
  })

  it('interruptible stores forward overlay publication and structure generation', async () => {
    const source = createObjectWorkbookSource(workbookOf([{ sheetId: 's', cells: [] }]), 'overlay-facade')
    const raw: any = source.store('s')
    const wrapped: any = interruptibleStore(raw)
    expect(wrapped.structureGen).toBe(0)
    wrapped.setValueOverlay({ rects: [{ r0: 1, r1: 1, c0: 1, c1: 1 }], protectedRects: [{ r0: 1, r1: 1, c0: 1, c1: 1 }], merges: [], cell: () => ({ value: 9 }) })
    expect(wrapped.structureGen).toBe(0)
    expect((await wrapped.readCell({ row: 1, col: 1 }, options(wrapped))).kind).toBe('cell')
    await source.release()
  })

  it('overrides source reads, blanks owned cells, grows extent and advances valueGen', async () => {
    const source = createObjectWorkbookSource(workbookOf([{ sheetId: 's', cells: [{ row: 0, col: 0, value: 1 }, { row: 1, col: 0, value: 2 }] }]), 'overlay')
    const store: any = source.store('s')
    const before = { ...store.stamp }
    const structureBefore = store.structureGen
    store.setValueOverlay(overlay)
    expect(store.stamp.valueGen).toBe(before.valueGen + 1)
    expect(store.structureGen).toBe(structureBefore)
    const overlaid = await store.readCell({ row: 0, col: 0 }, options(store))
    expect(overlaid).toMatchObject({ kind: 'cell', cell: { value: 100, valueProvenance: 'engine' } })
    if (overlaid.kind === 'cell') expect(source.displayText!(overlaid.cell, 's')).not.toBe('100')
    expect(await store.readCell({ row: 0, col: 1 }, options(store))).toEqual({ kind: 'blank' })
    expect((await store.readCell({ row: 2, col: 1 }, options(store)))).toMatchObject({ kind: 'cell', cell: { value: 'new' } })
    expect(store.extent.used).toEqual({ r0: 0, r1: 2, c0: 0, c1: 1 })
    const rows = await store.readRows({ r0: 0, r1: 2, c0: 0, c1: 1 }, undefined, options(store))
    expect(rows.items).toEqual([
      { row: 0, cells: [expect.objectContaining({ row: 0, col: 0, value: 100 })] },
      { row: 2, cells: [expect.objectContaining({ row: 2, col: 1, value: 'new' })] },
    ])
    expect(await store.readCell({ row: 0, col: 0 }, { signal: new AbortController().signal, stamp: before })).toMatchObject({ kind: 'error', code: 'stale-source' })
    const merged = { ...overlay, merges: [{ r0: 0, r1: 1, c0: 0, c1: 1 }] }
    store.setValueOverlay(merged)
    expect(store.structureGen).toBe(structureBefore + 1)
    store.setValueOverlay({ ...merged, rects: [...merged.rects] })
    expect(store.structureGen).toBe(structureBefore + 1)
    store.setValueOverlay(undefined)
    expect(store.structureGen).toBe(structureBefore + 2)
  })

  it('keeps overlay extent anchored at its true top-left and returns to the base extent on collapse', async () => {
    const doc = workbookOf([{ sheetId: 's', cells: [{ row: 8, col: 6, value: 1 }] }])
    doc.sheets[0].pivotRegions = [{ partPath: 'pivotTable1.xml', ref: 'G9:J13' }]
    const source = createObjectWorkbookSource(doc, 'overlay-extent')
    const store: any = source.store('s')
    const base = store.extent.used
    const expanded = { rects: [{ r0: 8, r1: 12, c0: 6, c1: 9 }], protectedRects: [{ r0: 8, r1: 12, c0: 6, c1: 9 }], merges: [], cell: () => null }
    store.setValueOverlay(expanded)
    expect(store.extent.used).toEqual({ r0: 8, r1: 12, c0: 6, c1: 9 })
    store.setValueOverlay({ ...expanded, rects: [{ r0: 8, r1: 8, c0: 6, c1: 6 }], protectedRects: [{ r0: 8, r1: 8, c0: 6, c1: 6 }]})
    expect(store.extent.used).toEqual(base)
    const structure = new StructuralIndex(await store.queryStructural({ r0: 8, r1: 12, c0: 6, c1: 9 }, options(store)))
    expect(structure.firstWhere({ r0: 12, r1: 12, c0: 9, c1: 9 }, hit => hit.kind === 'pivot')).toBeUndefined()
    expect(await store.readCell({ row: 12, col: 9 }, options(store))).toEqual({ kind: 'blank' })
    await source.release()
  })

  it('does not advance structureGen when overlay merges are only reordered', () => {
    const store: any = createObjectWorkbookSource(workbookOf([{ sheetId: 's', cells: [] }]), 'merge-order').store('s')
    const a = { r0: 0, r1: 1, c0: 0, c1: 1 }, b = { r0: 4, r1: 5, c0: 4, c1: 5 }
    store.setValueOverlay({ rects: [], protectedRects: [], merges: [a, b], cell: () => undefined })
    const generation = store.structureGen
    store.setValueOverlay({ rects: [], protectedRects: [], merges: [b, a], cell: () => undefined })
    expect(store.structureGen).toBe(generation)
  })

  it('formats an overlay without a format as General, never the base cell format', async () => {
    const doc = workbookOf([{ sheetId: 's', cells: [{ row: 0, col: 0, value: 0.5 }] }])
    doc.sheets[0].rows[0].cells[0].style = { numFmtId: 10, formatCode: '0.00%' }
    const source = createObjectWorkbookSource(doc, 'overlay-style')
    const store: any = source.store('s')
    store.setValueOverlay({ rects: [{ r0: 0, r1: 0, c0: 0, c1: 0 }], protectedRects: [{ r0: 0, r1: 0, c0: 0, c1: 0 }], merges: [],
      cell: () => ({ value: 0.25 }) })
    const read = await store.readCell({ row: 0, col: 0 }, options(store))
    expect(read.kind).toBe('cell')
    if (read.kind === 'cell') expect(source.displayText!(read.cell, 's')).toBe('0.25')
    await source.release()
  })

  it('preserves overlay errors and omits styled blanks from occupied scans', async () => {
    const doc = workbookOf([{ sheetId: 's', cells: [{ row: 0, col: 0, value: 1 }, { row: 4, col: 2, value: 0 }] }])
    doc.sheets[0].rows[1].cells[0] = { ref: 'C5', row: 4, col: 2, value: null, styleIndex: 7 }
    const source = createObjectWorkbookSource(doc, 'overlay-error')
    const store: any = source.store('s')
    store.setValueOverlay({ rects: [{ r0: 0, r1: 0, c0: 0, c1: 2 }], protectedRects: [{ r0: 0, r1: 0, c0: 0, c1: 2 }], merges: [],
      cell: (row: number, col: number) => row === 0 && col === 1 ? { value: '#N/A', valueIsError: true } : undefined })
    const error = await store.readCell({ row: 0, col: 1 }, options(store))
    expect(error).toMatchObject({ kind: 'cell', cell: { value: '#N/A', valueIsError: true } })
    const found: Array<[number, number]> = []
    for await (const page of store.scanOccupied({ range: { r0: 0, r1: 4, c0: 0, c1: 2 }, limit: 50 }, options(store))) {
      found.push(...page.items.map((cell: any) => [cell.row, cell.col]))
    }
    expect(found).toEqual([[0, 0], [0, 1]])
    await source.release()
  })

  it('charges overlay scan output pages through the caller allocation context', async () => {
    const source = createObjectWorkbookSource(workbookOf([{ sheetId: 's', cells: [] }]), 'overlay-scan-charge')
    const store: any = source.store('s')
    store.setValueOverlay({ rects: [{ r0: 0, r1: 2, c0: 0, c1: 0 }], protectedRects: [{ r0: 0, r1: 2, c0: 0, c1: 0 }], merges: [], cell: (row: number) => ({ value: row + 1 }) })
    let reserves = 0, adopted = 0
    const allocation: any = {
      limitBytes: 1_000_000, usedBytes: 0, owns: () => false, canReserve: () => true, settle: () => {},
      reserve() {
        reserves++; let bytes = 0
        return { owner: `scan-${reserves}`, get bytes() { return bytes }, charge(amount: number) { bytes += amount }, trySet(amount: number) { bytes = amount; return true }, release() {}, adopt(_targets: object[], amount?: number) { adopted += amount ?? bytes } }
      },
    }
    for await (const _page of store.scanPopulated({ range: { r0: 0, r1: 2, c0: 0, c1: 0 }, limit: 1 }, { ...options(store), [SOURCE_ALLOCATION]: allocation })) { /* drain */ }
    expect(reserves).toBeGreaterThanOrEqual(4)
    expect(adopted).toBeGreaterThan(0)
    await source.release()
  })

  it('resident geometry exposes overlay cells and replaces merges only inside owned rectangles', () => {
    const doc = workbookOf([{ sheetId: 's', cells: [{ row: 0, col: 0, value: 1 }] }])
    doc.sheets[0].mergeRanges = [{ minRow: 0, maxRow: 0, minCol: 0, maxCol: 1 }]
    const sheet = doc.sheets[0]
    const geometry = buildSheetGeometry(sheet, { overlay } as any)
    expect(geometry.kind).toBe('ready')
    if (geometry.kind !== 'ready') return
    expect(geometry.cellsOf(0)).toEqual([expect.objectContaining({ row: 0, col: 0, value: 100, style: { numFmtId: 14 } })])
    expect(geometry.cellsOf(2)).toEqual([expect.objectContaining({ row: 2, col: 1, value: 'new' })])
    expect(geometry.merges.collect({ r0: 0, r1: 0, c0: 0, c1: 1 })).toEqual([])
  })

  it('clears a vacated host across both 175% / DPR 2 tiles that intersect it', () => {
    const doc = workbookOf([{ sheetId: 's', cells: [{ row: 0, col: 2, value: 'vacated' }] }])
    const sheet = doc.sheets[0]
    const owner = { docId: 'tile-overlay', sheetId: 's', valueGen: 0, fontGen: 0, viewGen: 0 }
    const paint = (geometry: any, x: number) => {
      const canvas = createCanvas(512, 96)
      const ctx = canvas.getContext('2d')
      const texts: string[] = []
      const fillText = ctx.fillText.bind(ctx)
      ctx.fillText = ((text: string, x: number, y: number, maxWidth?: number) => {
        texts.push(text)
        return maxWidth === undefined ? fillText(text, x, y) : fillText(text, x, y, maxWidth)
      }) as typeof ctx.fillText
      renderSheetViewport(sheet, ctx as never, geometry, {
        box: { x, y: 0, width: 512 / (1.75 * 2), height: 96 / (1.75 * 2) }, zoom: 1.75, deviceScale: 2,
      }, {}, { owner, cache: new LayoutCache() })
      return texts
    }
    const before = buildSheetGeometry(sheet, {})
    expect(before.kind).toBe('ready')
    if (before.kind !== 'ready') return
    const leftBefore = paint(before, 0)
    const rightBefore = paint(before, 146)
    expect(leftBefore).toContain('vacated')
    expect(rightBefore).toContain('vacated')

    const clearingOverlay = { ...overlay, rects: [{ r0: 0, r1: 0, c0: 2, c1: 2 }], protectedRects: [], cell: () => null }
    const after = buildSheetGeometry(sheet, { overlay: clearingOverlay } as any)
    expect(after.kind).toBe('ready')
    if (after.kind !== 'ready') return
    expect(paint(after, 0)).not.toContain('vacated')
    expect(paint(after, 146)).not.toContain('vacated')
  })

  it('paged geometry exposes overlay cells and replaces intersecting merges', () => {
    const base = { id: 0, ref: 'A1', row: 0, col: 0, value: 1, styleIndex: 0 }
    const stableRow = [{ id: 5, ref: 'A6', row: 5, col: 0, value: 6, styleIndex: 0 }]
    const axes = { rows: createAxisMetrics({ logicalCount: MAX_ROWS, defaultSizeCss: 20 }), cols: createAxisMetrics({ logicalCount: MAX_COLS, defaultSizeCss: 64 }) }
    const remote: any = {
      geometry: () => axes,
      store: () => ({ cellsOf: (row: number) => row === 0 ? [base] : row === 5 ? stableRow : undefined }),
      merges: () => ({ merges: [{ minRow: 0, maxRow: 0, minCol: 0, maxCol: 1 }] }),
      drawings: () => ({ drawingsIn: () => [], drawings: [], diagnostics: [] }),
      overflowFor: () => ({}), profile: 'D', mainLedger: { bytes: 0 },
    }
    const geometry = pagedSheetGeometry(remote, 's', createIdentityOrder(), visibilityFromHidden(createIntervalSet([])), {} as any, undefined, true, overlay)
    expect(geometry.cellsOf(0)).toEqual([expect.objectContaining({ row: 0, col: 0, value: 100, style: { numFmtId: 14 } })])
    expect(geometry.cellsOf(2)).toEqual([expect.objectContaining({ row: 2, col: 1, value: 'new' })])
    expect(geometry.cellsOf(5)).toBe(stableRow)
    expect(geometry.merges.collect({ r0: 0, r1: 0, c0: 0, c1: 1 })).toEqual([])
  })

  it('column queries and populated scans publish overlay values in source order', async () => {
    const source = createObjectWorkbookSource(workbookOf([{ sheetId: 's', cells: [{ row: 0, col: 0, value: 1 }, { row: 1, col: 0, value: 2 }] }]), 'overlay')
    const store: any = source.store('s')
    store.setValueOverlay(overlay)
    const rows = { kind: 'intervals' as const, intervals: [{ r0: 0, r1: 2 }] }
    const column = await store.queryColumn({ col: 0, rows, limit: 10 }, options(store))
    expect(column.items).toEqual([
      { kind: 'cell', row: 0, cell: expect.objectContaining({ value: 100, valueProvenance: 'engine' }) },
      { kind: 'blank', r0: 1, r1: 2 },
    ])
    const seen = [] as Array<{ row: number; col: number; value: unknown }>
    for await (const page of store.scanPopulated({ range: { r0: 0, r1: 2, c0: 0, c1: 1 }, limit: 10 }, options(store))) {
      seen.push(...page.items.map((cell: any) => ({ row: cell.row, col: cell.col, value: cell.value })))
    }
    expect(seen).toEqual([{ row: 0, col: 0, value: 100 }, { row: 2, col: 1, value: 'new' }])
  })

  it('find and copy consume the overlaid value view', async () => {
    const doc = workbookOf([{ sheetId: 's', cells: [{ row: 0, col: 0, value: 'before' }] }])
    const source = createObjectWorkbookSource(doc, 'overlay-consumers')
    const store: any = source.store('s')
    store.setValueOverlay(overlay)
    const find = findInputOf(doc, { source, docId: 'overlay-consumers' })
    const baseOwner = initialWorkbookOwner('overlay-consumers', source.sheets)
    const owner = { ...baseOwner, sheetVector: baseOwner.sheetVector.map(entry => entry.sheetId === 's' ? { ...entry, valueGen: store.stamp.valueGen } : entry) }
    const sheet = find.input.sheets[0]
    const session = new WorkbookFind({ ...find.input, owner, sheets: [{ ...sheet, identity: { ...sheet.identity, valueGen: store.stamp.valueGen } }] }, { text: 'new', caseSensitive: false })
    expect((await drain(session)).map(hit => [hit.source.row, hit.source.col, hit.text])).toEqual([[2, 1, 'new']])

    const copyOwner = { ...store.stamp, fontGen: 0, viewGen: 0 }
    const copied = await serializeSelection({
      store, owner: copyOwner, currentOwner: () => copyOwner,
      selection: { active: null, anchor: null, ranges: [{ rows: { kind: 'intervals', intervals: [{ r0: 0, r1: 2 }] }, cols: [{ c0: 0, c1: 1 }] }] },
      order: createIdentityOrder(), visibility: visibilityFromHidden(createIntervalSet([])), columns: visibilityFromHidden(createIntervalSet([]), MAX_COLS),
      allowed: () => true, signal: new AbortController().signal, format: cell => source.displayText!(cell, 's'),
    })
    expect(copied).toMatchObject({ kind: 'ready', text: '4/9/1900\t\n\t\n\tnew' })
  })

  it('paged reads, scans and column lanes consult the committed overlay', async () => {
    const bytes = xlsxBytes({ sheetData: '<row r="1"><c r="A1"><v>1</v></c></row><row r="2"><c r="A2"><v>2</v></c></row>' })
    const pages = new SourcePageStore({ docId: 'overlay-paged', backend: memoryPageBackend() })
    const session = await openImportSession({ kind: 'blob', blob: new Blob([bytes.slice().buffer as ArrayBuffer]) }, { docId: 'overlay-paged', profile: 'D', store: pages, pageRows: 2 })
    expect((await session.done).kind).toBe('complete')
    const source = createPagedWorkbookSource(session)
    const store = source.store('1') as any
    store.setValueOverlay(overlay)
    const overlaid = await store.readCell({ row: 0, col: 0 }, options(store))
    expect(overlaid).toMatchObject({ kind: 'cell', cell: { value: 100, valueProvenance: 'engine' } })
    if (overlaid.kind === 'cell') expect(source.displayText!(overlaid.cell, '1')).not.toBe('100')
    const rowPage = await store.readRows({ r0: 0, r1: 2, c0: 0, c1: 1 }, undefined, options(store))
    expect(rowPage.items).toEqual([
      { row: 0, cells: [expect.objectContaining({ row: 0, col: 0, value: 100, valueProvenance: 'engine' })] },
      { row: 2, cells: [expect.objectContaining({ row: 2, col: 1, value: 'new' })] },
    ])
    const rows = { kind: 'intervals' as const, intervals: [{ r0: 0, r1: 2 }] }
    const column = await store.queryColumn({ col: 0, rows, limit: 10 }, options(store))
    expect(column.items).toEqual([
      { kind: 'cell', row: 0, cell: expect.objectContaining({ value: 100 }) },
      { kind: 'blank', r0: 1, r1: 2 },
    ])
    const lanes = await store.queryColumn({ col: 0, rows, limit: 3, encoding: 'lanes' }, options(store))
    expect(lanes.items).toMatchObject([{ kind: 'lanes', rows: new Uint32Array([0]), numbers: new Float64Array([100]) }])
    const seen: unknown[] = []
    for await (const page of store.scanPopulated({ range: { r0: 0, r1: 2, c0: 0, c1: 1 }, limit: 10 }, options(store))) {
      seen.push(...page.items.map((cell: any) => [cell.row, cell.col, cell.value]))
    }
    expect(seen).toEqual([[0, 0, 100], [2, 1, 'new']])
    const baseOwner = initialWorkbookOwner('overlay-paged', source.sheets)
    const owner = { ...baseOwner, sheetVector: baseOwner.sheetVector.map(entry => entry.sheetId === '1' ? { ...entry, valueGen: store.stamp.valueGen } : entry) }
    const visibility = visibilityFromHidden(createIntervalSet([]))
    const found = await drain(new WorkbookFind({ owner, sheets: [{ sheetId: '1', store, identity: { docId: store.stamp.docId, sheetId: '1', valueGen: store.stamp.valueGen, fontGen: owner.fontGen, viewGen: owner.sheetVector[0].revision }, order: createIdentityOrder(), visibility }] }, { text: 'new', caseSensitive: false }))
    expect(found.map(hit => [hit.source.row, hit.source.col, hit.text])).toEqual([[2, 1, 'new']])
    const copyOwner = { ...store.stamp, fontGen: 0, viewGen: 0 }
    const copied = await serializeSelection({
      store, owner: copyOwner, currentOwner: () => copyOwner,
      selection: { active: null, anchor: null, ranges: [{ rows: { kind: 'intervals', intervals: [{ r0: 0, r1: 2 }] }, cols: [{ c0: 0, c1: 1 }] }] },
      order: createIdentityOrder(), visibility, columns: visibilityFromHidden(createIntervalSet([]), MAX_COLS),
      allowed: () => true, signal: new AbortController().signal,
      format: cell => source.displayText!(cell, '1'),
    })
    expect(copied).toMatchObject({ kind: 'ready', text: '4/9/1900\t\n\t\n\tnew' })
    await source.release()
  })

  it('applies an overlay to answered column entries while preserving a loading suffix', async () => {
    const bytes = xlsxBytes({ sheetData: '<row r="1"><c r="A1"><v>1</v></c></row><row r="2"><c r="A2"><v>2</v></c></row>' })
    const pages = new SourcePageStore({ docId: 'overlay-column-partial', backend: memoryPageBackend() })
    const session = await openImportSession({ kind: 'blob', blob: new Blob([bytes.slice().buffer as ArrayBuffer]) }, { docId: 'overlay-column-partial', profile: 'D', store: pages, pageRows: 2 })
    expect((await session.done).kind).toBe('complete')
    const sheet = session.workbook.sheets[0] as any
    sheet.state = 'loading'; sheet.importedThroughRow = 0
    const source = createPagedWorkbookSource(session)
    const store: any = source.store('1')
    store.setValueOverlay({ rects: [{ r0: 0, r1: 0, c0: 0, c1: 0 }], protectedRects: [{ r0: 0, r1: 0, c0: 0, c1: 0 }], merges: [],
      cell: () => ({ value: 99 }) })
    const result = await store.queryColumn({ col: 0, rows: { kind: 'intervals', intervals: [{ r0: 0, r1: 1 }] }, limit: 10 }, options(store))
    expect(result.completeness).toBe('loading')
    expect(result.items).toMatchObject([{ kind: 'cell', row: 0, cell: { value: 99 } }, { kind: 'unknown', r0: 1, r1: 1 }])
    await source.release()
  })

  it('refuses sorts intersecting an overlay pivot and allows an adjacent band', async () => {
    const source = createObjectWorkbookSource(workbookOf([{ sheetId: 's', cells: [{ row: 0, col: 0, value: 1 }] }]), 'guard')
    const store: any = source.store('s')
    const protectedOverlay = { ...overlay, rects: [{ r0: 0, r1: 2, c0: 0, c1: 1 }], protectedRects: [{ r0: 0, r1: 2, c0: 0, c1: 1 }] }
    store.setValueOverlay(protectedOverlay)
    const structure = new StructuralIndex(await store.queryStructural({ r0: 0, r1: 4, c0: 0, c1: 2 }, options(store)))
    const pivotRange = { r0: 0, r1: 2, c0: 0, c1: 1, headerRows: 0 as const, totalsRows: 0 as const }
    const adjacentRange = { r0: 3, r1: 4, c0: 0, c1: 1, headerRows: 0 as const, totalsRows: 0 as const }
    const occupancy = fromStoreScan(store, options(store))
    expect(await guardSort(pivotRange, structure, occupancy)).toMatchObject({ ok: false, refusal: { reason: 'pivot', range: protectedOverlay.rects[0] } })
    expect(await guardSort(adjacentRange, structure, occupancy)).toEqual({ ok: true })
  })

  it('protects saved pivots before an overlay exists for both sort and filter', async () => {
    const doc = workbookOf([{ sheetId: 's', cells: [{ row: 0, col: 0, value: 1 }] }])
    doc.sheets[0].pivotRegions = [{ partPath: 'pivotTable1.xml', ref: 'A1:B2' }]
    const source = createObjectWorkbookSource(doc, 'saved-pivot-guard')
    const store: any = source.store('s')
    const structure = new StructuralIndex(await store.queryStructural({ r0: 0, r1: 3, c0: 0, c1: 3 }, options(store)))
    expect(structure.firstWhere({ r0: 0, r1: 1, c0: 0, c1: 1 }, hit => hit.kind === 'pivot')).toBeDefined()
    const pivotRange = { r0: 0, r1: 1, c0: 0, c1: 1, headerRows: 0 as const, totalsRows: 0 as const }
    const sort = await guardSort(pivotRange, structure, fromStoreScan(store, options(store)))
    expect(sort).toMatchObject({ ok: false, refusal: { reason: 'pivot' } })
    const axes = { rows: createAxisMetrics({ logicalCount: MAX_ROWS, defaultSizeCss: 20 }), cols: createAxisMetrics({ logicalCount: MAX_COLS, defaultSizeCss: 64 }) }
    const beforeVisibility = visibilityFromHidden(createIntervalSet([]))
    const afterVisibility = visibilityFromHidden(createIntervalSet([{ r0: 0, r1: 0 }]))
    const before = createSheetMetrics({ ...axes, order: createIdentityOrder(), visibility: beforeVisibility })
    const after = createSheetMetrics({ ...axes, order: createIdentityOrder(), visibility: afterVisibility })
    const filter = await guardFilter({ range: pivotRange, candidateVisibility: afterVisibility, currentVisibility: beforeVisibility, projectionBefore: before, projectionAfter: after, structure })
    expect(filter).toMatchObject({ ok: false, refusal: { reason: 'pivot' } })
    await source.release()
  })

  it('invalidates a cached table/pivot region style layer after pivot definitions change', () => {
    const sheet = { name: 'S', pivotTables: [{}], regionStyles: {}, regionStylesResolved: true } as any
    invalidateRegionStyles(sheet)
    expect(sheet.regionStyles).toBeUndefined()
    expect(sheet.regionStylesResolved).toBe(false)
    const remoteCalls: string[] = []
    invalidateRegionStyles({ invalidateRegionStyles: (sheetId: string) => remoteCalls.push(sheetId) }, 'paged-sheet' as any)
    expect(remoteCalls).toEqual(['paged-sheet'])
  })
})

describe('pivot value overlay (fix round 4)', () => {
  /** An allocation context that records each adoption (pages it charges and adopts). */
  const recordingAllocation = () => {
    const seen = { adopted: 0, adopts: 0 }
    const allocation: any = {
      limitBytes: 1e12, usedBytes: 0, canReserve: () => true, owns: () => false, settle: () => {},
      reserve() {
        let held = 0
        return { owner: 'scan', get bytes() { return held }, trySet(bytes: number) { held = bytes; return true }, charge(bytes: number) { held += bytes }, release(bytes = held) { held -= bytes },
          adopt(_targets: object[], bytes = held) { seen.adopted += bytes; seen.adopts++; held -= bytes } }
      },
    }
    return { allocation, seen }
  }

  it('N2: a resident overlay scan charges each cell once, through its own output pages only', async () => {
    const N = 1_000
    const scan = async (withOverlay: boolean) => {
      const cells = withOverlay ? [{ row: 5_000, col: 5, value: 1 }] : Array.from({ length: N }, (_, row) => ({ row, col: 0, value: row }))
      const source = createObjectWorkbookSource(workbookOf([{ sheetId: 's', cells }]), `n2-${withOverlay}`)
      const store: any = source.store('s')
      if (withOverlay) store.setValueOverlay({ rects: [{ r0: 0, r1: N - 1, c0: 0, c1: 0 }], protectedRects: [{ r0: 0, r1: N - 1, c0: 0, c1: 0 }], merges: [], cell: (row: number) => ({ value: row }) })
      const { allocation, seen } = recordingAllocation()
      let count = 0, pages = 0
      for await (const page of store.scanPopulated({ range: { r0: 0, r1: N - 1, c0: 0, c1: 0 }, limit: 10_000 }, { ...options(store), [SOURCE_ALLOCATION]: allocation })) { count += page.items.length; pages++ }
      await source.release()
      return { count, pages, ...seen }
    }
    const base = await scan(false), overlaid = await scan(true)
    expect(overlaid.count).toBe(N)
    // One adoption per published page, as on the base path; no inner readRows page is charged or adopted.
    expect(overlaid.adopts).toBe(overlaid.pages)
    expect(overlaid.adopted).toBeLessThanOrEqual(base.adopted * 1.2)
  })

  it('N1: paged geometry re-merges a row whose base cells arrive, change or are evicted after it was painted', () => {
    const axes = { rows: createAxisMetrics({ logicalCount: MAX_ROWS, defaultSizeCss: 20 }), cols: createAxisMetrics({ logicalCount: MAX_COLS, defaultSizeCss: 64 }) }
    const beside = { id: 9, ref: 'J1', row: 0, col: 9, value: 'beside', styleIndex: 0 }
    let row0: any[] | undefined
    const remote: any = {
      geometry: () => axes,
      store: () => ({ cellsOf: (row: number) => row === 0 ? row0 : undefined }),
      merges: () => ({ merges: [] }),
      drawings: () => ({ drawingsIn: () => [], drawings: [], diagnostics: [] }),
      overflowFor: () => ({}), profile: 'D', mainLedger: { bytes: 0 },
    }
    const geometry = pagedSheetGeometry(remote, 's', createIdentityOrder(), visibilityFromHidden(createIntervalSet([])), {} as any, undefined, true, overlay)
    const values = () => geometry.cellsOf(0)?.map(cell => cell.value)
    expect(values()).toEqual([100])
    row0 = [beside]
    expect(values()).toEqual([100, 'beside'])
    // A repeated read of the same base row is memoized (one merged array per base row).
    expect(geometry.cellsOf(0)).toBe(geometry.cellsOf(0))
    row0 = [{ ...beside, value: 'refetched' }]
    expect(values()).toEqual([100, 'refetched'])
    row0 = undefined
    expect(values()).toEqual([100])
  })

  it('N6: an overlay names its protected bounds; vacated cells are ordinary empty cells on both stores', async () => {
    // @ts-expect-error protectedRects is required: a fallback to rects would protect vacated cells
    const missing: ValueOverlay = { rects: [], merges: [], cell: () => undefined }
    // @ts-expect-error the overlay carries no row/column counts; the extent comes from protectedRects
    const counted: ValueOverlay = { rects: [], protectedRects: [], merges: [], cell: () => undefined, usedRows: 1, usedCols: 1 }
    void missing; void counted
    // The pivot shrank from A1:B3 to A1:A1: B1:B3 and A2:A3 are vacated (owned, null, unprotected).
    const shrunk: ValueOverlay = { rects: [{ r0: 0, r1: 2, c0: 0, c1: 1 }], protectedRects: [{ r0: 0, r1: 0, c0: 0, c1: 0 }], merges: [], cell: (row, col) => row === 0 && col === 0 ? { value: 7 } : null }
    const resident = createObjectWorkbookSource(workbookOf([{ sheetId: 's', cells: [{ row: 0, col: 0, value: 1 }] }]), 'n6')
    const bytes = xlsxBytes({ sheetData: '<row r="1"><c r="A1"><v>1</v></c></row>' })
    const session = await openImportSession({ kind: 'blob', blob: new Blob([bytes.slice().buffer as ArrayBuffer]) }, { docId: 'n6-paged', profile: 'D', store: new SourcePageStore({ docId: 'n6-paged', backend: memoryPageBackend() }), pageRows: 2 })
    expect((await session.done).kind).toBe('complete')
    const paged = createPagedWorkbookSource(session)
    for (const store of [resident.store('s'), paged.store('1')] as any[]) {
      store.setValueOverlay(shrunk)
      expect(store.extent.used).toEqual({ r0: 0, r1: 0, c0: 0, c1: 0 })
      expect(await store.readCell({ row: 2, col: 1 }, options(store))).toEqual({ kind: 'blank' })
      const structure = new StructuralIndex(await store.queryStructural({ r0: 0, r1: 3, c0: 0, c1: 2 }, options(store)))
      const vacated = { r0: 1, r1: 2, c0: 0, c1: 1, headerRows: 0 as const, totalsRows: 0 as const }
      expect(await guardSort(vacated, structure, fromStoreScan(store, options(store)))).toEqual({ ok: true })
      const pivot = { r0: 0, r1: 1, c0: 0, c1: 0, headerRows: 0 as const, totalsRows: 0 as const }
      expect(await guardSort(pivot, structure, fromStoreScan(store, options(store)))).toMatchObject({ ok: false, refusal: { reason: 'pivot' } })
    }
    await resident.release(); await paged.release()
  })

  it('N5: remote invalidation drops the pivot display layer too, and a remote call without a sheet id throws', () => {
    const invalidated: string[] = []
    const remote = Object.create(RemoteWorkbook.prototype)
    remote.regionStyleCache = { invalidate: (sheetId: string) => invalidated.push(`regionStyles:${sheetId}`) }
    remote.pivotDisplayCache = { invalidate: (sheetId: string) => invalidated.push(`pivotDisplay:${sheetId}`) }
    invalidateRegionStyles(remote, 's')
    expect(invalidated).toEqual(['regionStyles:s', 'pivotDisplay:s'])
    expect(() => invalidateRegionStyles(remote)).toThrow(/sheet id/i)
  })
})
