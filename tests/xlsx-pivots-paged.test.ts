// @vitest-environment node
import { readFileSync } from 'node:fs'
import { resolve } from 'node:path'
import { createCanvas } from 'canvas'
import JSZip from 'jszip'
import { describe, expect, it } from 'vitest'
import type { CacheIdentity, RemotePivotDisplay, SourceStore, WorkerReply } from '../src/xlsx/view/contracts'
import { openImportSession } from '../src/xlsx/storage/pagedImport'
import { createPagedWorkbookSource } from '../src/xlsx/view/pagedSource'
import { RemoteRegionStyleCache } from '../src/worker/viewerRemoteRegionStyles'
import { OfficePackage } from '../src/core/zip'
import { parseXlsx } from '../src/xlsx/parse'
import { computeMetrics, renderSheet } from '../src/xlsx/render'
import { ensurePivotEngine } from '../src/xlsx/pivotLoader'
import { ensureTableStyleEngine } from '../src/xlsx/tableStyleLoader'
import { sheetWithRegionStyles } from '../src/worker/viewerRemoteRegionStyles'
import { sheetWithDrawings } from '../src/worker/viewerRemoteDrawings'
import { ViewerKernel, connectInProcess } from '../src/worker/viewerKernel'
import { ViewerClient } from '../src/worker/viewerClient'
import { openRemoteWorkbook, RemoteWorkbook } from '../src/worker/viewerRemoteSource'
import { renderSheetViewport } from '../src/xlsx/view/renderViewport'
import { pagedSheetGeometry } from '../src/components/xlsx/pagedGeometry'
import { createIdentityOrder } from '../src/xlsx/view/order'
import { createIntervalSet, visibilityFromHidden } from '../src/xlsx/view/visibility'
import { validateOverflowProof } from '../src/xlsx/view/viewportOverflow'
import { formatA1 } from '../src/xlsx/view/coordinates'
import { emptyWork, type ViewportPlan } from '../src/xlsx/view/viewport'
import { interruptibleStore } from '../src/worker/viewerAbort'

const bytes = (name: string) => new Uint8Array(readFileSync(resolve(process.cwd(), `tests/fixtures/xlsx-pivots/${name}.xlsx`)))
function pixels(sheet: Parameters<typeof computeMetrics>[0]): Buffer {
  const metrics = computeMetrics(sheet), canvas = createCanvas(metrics.widthPx, metrics.heightPx)
  renderSheet(sheet, canvas.getContext('2d') as unknown as CanvasRenderingContext2D, metrics)
  return Buffer.from(canvas.getContext('2d').getImageData(0, 0, canvas.width, canvas.height).data.buffer)
}

describe('paged pivot display descriptors', () => {
  it('forwards the active profile through the interruptible store wrapper', async () => {
    const stamp = { docId: 'profile-forward', sheetId: 'sheet-1', valueGen: 0 }
    let received: string | undefined
    const payload: RemotePivotDisplay = { sheetId: stamp.sheetId, stamp, tables: [], pivotTables: [], slicers: [], timelines: [], diagnostics: [] }
    const store = {
      stamp, extent: { rows: 1, cols: 1 }, valuePolicy: 'cached',
      readPivotDisplay: async (_options: unknown, _grantBytes: number, _memory: unknown, profile?: string) => { received = profile; return payload },
      release: async () => {},
    } as unknown as SourceStore
    const wrapped = interruptibleStore(store)
    await wrapped.readPivotDisplay!({ signal: new AbortController().signal, stamp }, 4096, undefined, 'M')
    expect(received).toBe('M')
  })

  it('exposes the receiver-granted pivot display read for a sheet with pivot or control metadata', async () => {
    const input = bytes('pivot-layout-compact')
    const session = await openImportSession({ kind: 'blob', blob: new Blob([input]) }, { docId: 'pivot-display-contract', profile: 'D' })
    await session.done
    const source = createPagedWorkbookSource(session)
    try {
      const sheet = source.sheets[0]!
      expect(source.hasPivotDisplay?.(sheet.sheetId)).toBe(true)
      const store = source.store(sheet.sheetId)
      expect(store.readPivotDisplay).toBeTypeOf('function')
      const result = await store.readPivotDisplay!({ signal: new AbortController().signal, stamp: store.stamp }, 1024 * 1024, { set: () => true })
      expect(result.pivotTables.length).toBeGreaterThan(0)
      expect('buttons' in result).toBe(false)
      expect(result.slicers).toEqual([])
      expect(result.timelines).toEqual([])
      expect(result.diagnostics).not.toContainEqual(expect.objectContaining({ feature: 'pivot-display-reply-limit' }))
    } finally {
      await source.release()
      await session.release()
    }
  })

  it.each([
    ['slicer-light1', 'slicers'],
    ['timeline-light1-months', 'timelines'],
  ] as const)('resolves %s controls in the worker display descriptor', async (name, key) => {
    const input = bytes(name)
    const session = await openImportSession({ kind: 'blob', blob: new Blob([input]) }, { docId: `pivot-display-${name}`, profile: 'D' })
    await session.done
    const source = createPagedWorkbookSource(session)
    try {
      const sheet = source.sheets[0]!, store = source.store(sheet.sheetId)
      expect(source.hasPivotDisplay?.(sheet.sheetId)).toBe(true)
      const result = await store.readPivotDisplay!({ signal: new AbortController().signal, stamp: store.stamp }, 1024 * 1024, { set: () => true })
      expect(result[key].length).toBeGreaterThan(0)
      expect(result.diagnostics).not.toContainEqual(expect.objectContaining({ feature: 'pivot-display-reply-limit' }))
    } finally { await source.release(); await session.release() }
  })

  it('retains a referenced custom table style after more than 2 MiB of unused definitions', async () => {
    const archive = await JSZip.loadAsync(bytes('table-medium-2'))
    const tablePath = Object.keys(archive.files).find(path => /^xl\/tables\/table\d+\.xml$/.test(path))!
    const tableText = await archive.file(tablePath)!.async('string')
    archive.file(tablePath, tableText.replace('TableStyleMedium2', 'LateReferenced'))
    const stylesPath = 'xl/styles.xml', stylesText = await archive.file(stylesPath)!.async('string')
    const unused = Array.from({ length: 12_000 }, (_, i) => `<tableStyle name="UnusedStyle${i}" pivot="1" table="1"><tableStyleElement type="wholeTable" dxfId="0"/></tableStyle>`).join('')
    const used = '<tableStyle name="LateReferenced" pivot="1" table="1"><tableStyleElement type="wholeTable" dxfId="0"/></tableStyle>'
    const mutated = stylesText
      .replace('<dxfs count="0"/>', '<dxfs count="1"><dxf><fill><patternFill patternType="solid"><fgColor rgb="FFFF0000"/></patternFill></fill></dxf></dxfs>')
      .replace(/<tableStyles count="0"([^>]*)\/>/, `<tableStyles count="12001"$1>${unused}${used}</tableStyles>`)
    archive.file(stylesPath, mutated)
    const input = new Uint8Array(await archive.generateAsync({ type: 'uint8array' }))
    const session = await openImportSession({ kind: 'blob', blob: new Blob([input]) }, { docId: 'referenced-style-priority', profile: 'D' })
    await session.done
    const source = createPagedWorkbookSource(session)
    try {
      expect(session.workbook.tableStyleDefs?.has('LateReferenced')).toBe(true)
      expect(session.workbook.tableStyleDefs?.get('LateReferenced')?.elements.wholeTable?.dxf.fill?.fg).toEqual({ rgb: 'FFFF0000' })
      const sheet = source.sheets[0]!, store = source.store(sheet.sheetId)
      const display = await store.readPivotDisplay!({ signal: new AbortController().signal, stamp: store.stamp }, 1024 * 1024, { set: () => true })
      expect(display.tables[0]?.def.name).toBe('LateReferenced')
      expect(display.tables[0]?.def.elements.wholeTable?.dxf.fill?.fg).toEqual({ rgb: 'FFFF0000' })
    } finally { await source.release(); await session.release() }
  })

  it.each(['pivot-layout-compact', 'slicer-light1', 'timeline-light1-months'])('matches resident pixels through paged worker descriptors and drawings for %s', async name => {
    const input = bytes(name)
    const session = await openImportSession({ kind: 'blob', blob: new Blob([input]) }, { docId: `pivot-pixels-${name}`, profile: 'D' })
    await session.done
    const pagedSource = createPagedWorkbookSource(session)
    const kernel = new ViewerKernel({ parse: async () => { throw new Error('expected supplied paged source') }, openSource: async () => pagedSource }, { yield: async () => {} })
    const client = new ViewerClient(connectInProcess(kernel), `pivot-pixels-${name}`)
    const opened = await openRemoteWorkbook(client, `pivot-pixels-${name}`, new Blob([input]), 'D')
    if (!('layouts' in opened)) throw new Error('paged workbook did not open')
    const remote = new RemoteWorkbook(client, opened), sheetId = remote.sheets[0]!.sheetId
    try {
      for (let i = 0; i < 100 && remote.loadState(sheetId).completeness !== 'complete'; i++) await new Promise(resolveTick => setTimeout(resolveTick, 2))
      remote.pivotDisplay(sheetId)
      for (let i = 0; i < 200 && !['ready', 'failed'].includes(remote.pivotDisplay(sheetId).state); i++) await new Promise(resolveTick => setTimeout(resolveTick, 2))
      const display = remote.pivotDisplay(sheetId)
      expect(display.state, 'worker descriptor cache').toBe('ready')
      remote.drawings(sheetId)
      for (let i = 0; i < 200 && !['ready', 'failed'].includes(remote.drawings(sheetId).state); i++) await new Promise(resolveTick => setTimeout(resolveTick, 2))
      expect(remote.drawings(sheetId).state, 'worker drawing/control cache').toBe('ready')
      const resident = await parseXlsx(await OfficePackage.load(input), { valuePolicy: 'cached' })
      const residentSheet = resident.sheets[0]!
      await ensureTableStyleEngine([residentSheet])
      await ensurePivotEngine([residentSheet])
      const pagedSheet = sheetWithRegionStyles(sheetWithDrawings(residentSheet, remote.drawings(sheetId)), display)
      expect(pixels(pagedSheet)).toEqual(pixels(residentSheet))
      // Compare a real worker viewport plan painted by the interactive compositor.
      const window = { box: { x: 0, y: 0, width: 480, height: 360 }, zoom: 1, deviceScale: 1 }
      const geometry = pagedSheetGeometry(remote, sheetId, createIdentityOrder(), visibilityFromHidden(createIntervalSet([])))
      if (geometry.kind !== 'ready') throw new Error('paged viewport geometry is not ready')
      const frame = remote.currentPageFrame, tile = `pivot-parity:${sheetId}`, pageLimit = remote.pageLimits.tile
      expect(remote.admitPages(frame, tile, pageLimit, false, false)).toBe(true)
      expect(remote.admitPages(frame, tile, pageLimit, false)).toBe(true)
      const viewport = await remote.viewport(remote.identity(sheetId), window, new AbortController().signal, 'tile')
      expect(typeof viewport).toBe('object')
      if (typeof viewport !== 'object' || viewport.kind !== 'complete') throw new Error('paged worker viewport did not complete')
      const result = viewport.result
      if (result.content.kind !== 'detail') throw new Error('paged worker viewport was not detailed')
      const proof = validateOverflowProof(result, remote.identity(sheetId), window, 'D', 'tile', remote.columns(sheetId))
      const cells = result.content.cells.flatMap(item => {
        if (item.kind !== 'cell' && item.kind !== 'unavailable') return []
        const overflow = proof.get(item.source.row * 16384 + item.source.col)!
        const style = remote.styleOf(item.cell.styleIndex)
        return [{ source: item.source, x: item.bounds.x, y: item.bounds.y, width: item.bounds.width, height: item.bounds.height,
          overflow: { left: overflow.left, right: overflow.right }, cell: { ...item.cell, ref: formatA1(item.source), ...(style ? { style } : {}) } }]
      })
      const axes = remote.geometry(sheetId), extent = { width: axes.cols.totalCss(), height: axes.rows.totalCss() }
      const plan: Extract<ViewportPlan, { kind: 'detail' }> = { kind: 'detail', cells, merges: (result.merges ?? []).map(merge => ({ range: merge.range, ...merge.bounds })), drawings: [...remote.drawings(sheetId).drawings], rowLines: result.rowLines ?? [], colLines: result.colLines ?? [], extent, diagnostics: [], coarse: result.coarse ?? [], work: result.work ?? emptyWork() }
      const paint = (sheet: typeof pagedSheet) => {
        const canvas = createCanvas(window.box.width, window.box.height)
        const ctx = canvas.getContext('2d') as unknown as CanvasRenderingContext2D
        renderSheetViewport(sheet, ctx, geometry, window, undefined, { preparedPlan: plan })
        return Buffer.from(ctx.getImageData(0, 0, canvas.width, canvas.height).data.buffer)
      }
      expect(paint(pagedSheet)).toEqual(paint(residentSheet))
      remote.mainLedger.settle(viewport)
    } finally { await remote.close(); await pagedSource.release(); await session.release() }
  })

  it('returns a typed pivot-display-reply-limit and exact retry size when the receiver grant is too small', async () => {
    const input = bytes('pivot-layout-compact')
    const session = await openImportSession({ kind: 'blob', blob: new Blob([input]) }, { docId: 'pivot-display-small-grant', profile: 'D' })
    await session.done
    const source = createPagedWorkbookSource(session)
    try {
      const sheet = source.sheets[0]!, store = source.store(sheet.sheetId)
      const result = await store.readPivotDisplay!({ signal: new AbortController().signal, stamp: store.stamp }, 1, { set: () => true })
      expect(result.pivotTables).toEqual([])
      expect(result.requiredBytes).toBeGreaterThan(1)
      expect(result.diagnostics).toContainEqual(expect.objectContaining({ feature: 'pivot-display-reply-limit' }))
    } finally { await source.release(); await session.release() }
  })

  it.each([
    ['pivot-layout-compact', /xl\/pivotTables\/pivotTable[^/]*\.xml$/i, /<pivotTableDefinition\b/],
    ['slicer-light1', /xl\/slicerCaches\/slicerCache[^/]*\.xml$/i, /<slicerCacheDefinition\b/],
  ] as const)('reports pivot-part-too-large for over-cap package parts (%s)', async (fixture, pathPattern, rootPattern) => {
    const archive = await JSZip.loadAsync(bytes(fixture))
    const path = Object.keys(archive.files).find(name => pathPattern.test(name))
    expect(path, 'fixture part').toBeTruthy()
    const original = await archive.file(path!)!.async('string')
    expect(original).toMatch(rootPattern)
    const rootEnd = original.indexOf('>')
    archive.file(path!, `${original.slice(0, rootEnd + 1)}<!--${' '.repeat(70 * 1024)}-->${original.slice(rootEnd + 1)}`)
    const inflated = new Uint8Array(await archive.generateAsync({ type: 'uint8array' }))
    const session = await openImportSession({ kind: 'blob', blob: new Blob([inflated]) }, { docId: `pivot-cap-${fixture}`, profile: 'D' })
    await session.done
    const source = createPagedWorkbookSource(session)
    try {
      const sheet = source.sheets[0]!, store = source.store(sheet.sheetId)
      const result = await store.readPivotDisplay!({ signal: new AbortController().signal, stamp: store.stamp }, 4 * 1024 * 1024, { set: () => true })
      expect(result.diagnostics).toContainEqual(expect.objectContaining({ feature: 'pivot-part-too-large' }))
    } finally { await source.release(); await session.release() }
  })

  it('releases retained descriptors on sheet-owner change and close, and settles error replies', async () => {
    const owner: CacheIdentity = { docId: 'pivot-cache', sheetId: 's', valueGen: 1, fontGen: 0, viewGen: 0 }
    let currentOwner = { ...owner }
    let released = 0, resolve!: (value: WorkerReply<'read'>) => void
    const success = (stamp = currentOwner) => ({ kind: 'complete', commandKind: 'read', requestId: 'r', owner: stamp, diagnostics: [], replyBytes: 64,
      result: { kind: 'read', result: { op: 'pivotDisplay', result: { sheetId: 's', stamp, tables: [], pivotTables: [], slicers: [], timelines: [], diagnostics: [] } } } }) as unknown as WorkerReply<'read'>
    const cache = new RemoteRegionStyleCache({
      identity: () => currentOwner, hasRegionStyles: () => false, hasPivotDisplay: () => true, loadState: () => ({ completeness: 'complete' }),
      grant: () => 4096, releaseGrant: () => {}, request: () => new Promise<WorkerReply<'read'>>(done => { resolve = done }),
      nextRequestId: () => 'r', release: () => { released++ }, admitLayer: () => () => {}, changed: () => {},
    }, 4096, 'pivotDisplay')
    cache.get('s'); resolve(success())
    await new Promise(resolveTick => setTimeout(resolveTick, 0))
    expect(cache.get('s').state).toBe('ready')
    currentOwner = { ...currentOwner, valueGen: 2 }
    cache.get('s')
    expect(released).toBe(1)
    resolve(success(currentOwner))
    await new Promise(resolveTick => setTimeout(resolveTick, 0))
    cache.closeSheet('s')
    expect(released).toBe(2)
    cache.close()

    const failing = new RemoteRegionStyleCache({
      identity: () => currentOwner, hasRegionStyles: () => false, hasPivotDisplay: () => true, loadState: () => ({ completeness: 'complete' }),
      grant: () => 4096, releaseGrant: () => {}, request: async () => ({ kind: 'error', commandKind: 'read', requestId: 'err', owner, code: 'parse-error', retryable: false, diagnostics: [] } as unknown as WorkerReply<'read'>),
      nextRequestId: () => 'err', release: () => { released++ }, admitLayer: () => () => {}, changed: () => {},
    }, 4096, 'pivotDisplay')
    failing.get('s')
    await new Promise(resolveTick => setTimeout(resolveTick, 0))
    expect(released).toBe(3)
    failing.close()
  })
})
