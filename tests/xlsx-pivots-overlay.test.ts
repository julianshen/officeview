import { describe, expect, it } from 'vitest'
import { createObjectWorkbookSource } from '../src/xlsx/view/source'
import { buildSheetGeometry } from '../src/xlsx/view/viewport'
import { createPagedWorkbookSource } from '../src/xlsx/view/pagedSource'
import { openImportSession } from '../src/xlsx/storage/pagedImport'
import { SourcePageStore, memoryPageBackend } from '../src/xlsx/storage/sourcePages'
import { xlsxBytes } from './helpers/xlsx-package'
import { StructuralIndex, fromStoreScan } from '../src/xlsx/view/structure'
import { guardSort } from '../src/xlsx/view/operationGuard'
import { invalidateRegionStyles } from '../src/xlsx/tableStyleLoader'
import { pagedSheetGeometry } from '../src/components/xlsx/pagedGeometry'
import { createAxisMetrics } from '../src/xlsx/view/axes'
import { createIdentityOrder } from '../src/xlsx/view/order'
import { createIntervalSet, visibilityFromHidden } from '../src/xlsx/view/visibility'
import { MAX_COLS, MAX_ROWS } from '../src/xlsx/view/coordinates'
import { drain, findInputOf, workbookOf } from './helpers/xlsx-find'
import { WorkbookFind } from '../src/xlsx/view/find'
import { serializeSelection } from '../src/xlsx/view/copy'
import { initialWorkbookOwner } from '../src/worker/viewerProtocol'
import { createCanvas } from 'canvas'
import { renderSheetViewport } from '../src/xlsx/view/renderViewport'
import { LayoutCache } from '../src/xlsx/view/textLayout'

const options = (store: any) => ({ signal: new AbortController().signal, stamp: { ...store.stamp } })
const overlay = {
  rects: [{ r0: 0, r1: 2, c0: 0, c1: 1 }], merges: [], usedRows: 3, usedCols: 2,
  cell(row: number, col: number) {
    if (row === 0 && col === 0) return { value: 100, numFmtId: 14 }
    if (row === 2 && col === 1) return { value: 'new', numFmtId: 0 }
    return null
  },
}

describe('pivot value overlay', () => {
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
    let geometryChecks = 0
    const geometry = buildSheetGeometry(doc.sheets[0], Object.defineProperty({}, 'overlay', { get: () => { geometryChecks++; return undefined } }))
    expect(geometry.kind).toBe('ready')
    expect(geometryChecks).toBe(1)
    const noPivotStructure = new StructuralIndex({
      stamp: store.stamp, completeness: 'complete', missing: [], merges: [], tables: [], spills: [], drawings: [],
    })
    expect(Object.prototype.hasOwnProperty.call(noPivotStructure, 'pivots')).toBe(false)
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
    store.setValueOverlay({ ...merged, usedRows: 4 })
    expect(store.structureGen).toBe(structureBefore + 1)
    store.setValueOverlay(undefined)
    expect(store.structureGen).toBe(structureBefore + 2)
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

    const clearingOverlay = { ...overlay, rects: [{ r0: 0, r1: 0, c0: 2, c1: 2 }], usedRows: 1, usedCols: 3, cell: () => null }
    const after = buildSheetGeometry(sheet, { overlay: clearingOverlay } as any)
    expect(after.kind).toBe('ready')
    if (after.kind !== 'ready') return
    expect(paint(after, 0)).not.toContain('vacated')
    expect(paint(after, 146)).not.toContain('vacated')
  })

  it('paged geometry exposes overlay cells and replaces intersecting merges', () => {
    const base = { id: 0, ref: 'A1', row: 0, col: 0, value: 1, styleIndex: 0 }
    const axes = { rows: createAxisMetrics({ logicalCount: MAX_ROWS, defaultSizeCss: 20 }), cols: createAxisMetrics({ logicalCount: MAX_COLS, defaultSizeCss: 64 }) }
    const remote: any = {
      geometry: () => axes,
      store: () => ({ cellsOf: (row: number) => row === 0 ? [base] : undefined }),
      merges: () => ({ merges: [{ minRow: 0, maxRow: 0, minCol: 0, maxCol: 1 }] }),
      drawings: () => ({ drawingsIn: () => [], drawings: [], diagnostics: [] }),
      overflowFor: () => ({}), profile: 'D', mainLedger: { bytes: 0 },
    }
    const geometry = pagedSheetGeometry(remote, 's', createIdentityOrder(), visibilityFromHidden(createIntervalSet([])), {} as any, undefined, true, overlay)
    expect(geometry.cellsOf(0)).toEqual([expect.objectContaining({ row: 0, col: 0, value: 100, style: { numFmtId: 14 } })])
    expect(geometry.cellsOf(2)).toEqual([expect.objectContaining({ row: 2, col: 1, value: 'new' })])
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

  it('refuses sorts intersecting an overlay pivot and allows an adjacent band', async () => {
    const source = createObjectWorkbookSource(workbookOf([{ sheetId: 's', cells: [{ row: 0, col: 0, value: 1 }] }]), 'guard')
    const store: any = source.store('s')
    const protectedOverlay = { ...overlay, rects: [{ r0: 0, r1: 2, c0: 0, c1: 1 }] }
    store.setValueOverlay(protectedOverlay)
    const structure = new StructuralIndex(await store.queryStructural({ r0: 0, r1: 4, c0: 0, c1: 2 }, options(store)))
    const pivotRange = { r0: 0, r1: 2, c0: 0, c1: 1, headerRows: 0 as const, totalsRows: 0 as const }
    const adjacentRange = { r0: 3, r1: 4, c0: 0, c1: 1, headerRows: 0 as const, totalsRows: 0 as const }
    const occupancy = fromStoreScan(store, options(store))
    expect(await guardSort(pivotRange, structure, occupancy)).toMatchObject({ ok: false, refusal: { reason: 'pivot', range: protectedOverlay.rects[0] } })
    expect(await guardSort(adjacentRange, structure, occupancy)).toEqual({ ok: true })
  })

  it('invalidates a cached table/pivot region style layer after pivot definitions change', () => {
    const sheet = { name: 'S', pivotTables: [{}], regionStyles: {}, regionStylesResolved: true } as any
    invalidateRegionStyles(sheet)
    expect(sheet.regionStyles).toBeUndefined()
    expect(sheet.regionStylesResolved).toBe(false)
  })
})
