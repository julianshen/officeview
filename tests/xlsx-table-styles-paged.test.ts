// @vitest-environment node
import { readFileSync } from 'node:fs'
import { resolve } from 'node:path'
import { createCanvas } from 'canvas'
import JSZip from 'jszip'
import { describe, expect, it } from 'vitest'
import { OfficePackage } from '../src/core/zip'
import { parseXlsx } from '../src/xlsx/parse'
import { computeMetrics, renderSheet } from '../src/xlsx/render'
import { openImportSession } from '../src/xlsx/storage/pagedImport'
import { createPagedWorkbookSource } from '../src/xlsx/view/pagedSource'
import type { RemoteRegionStyles } from '../src/xlsx/view/contracts'
import type { CacheIdentity, WorkerReply } from '../src/xlsx/view/contracts'
import { buildTableRegions } from '../src/xlsx/tableStyles/regionLayer'
import type { StylesExtras } from '../src/xlsx/tableStyles/model'
import { RemoteRegionStyleCache } from '../src/worker/viewerRemoteRegionStyles'
import { ViewerClient } from '../src/worker/viewerClient'
import { RemoteWorkbook } from '../src/worker/viewerRemoteSource'
import type { WorkerCommand } from '../src/xlsx/view/contracts'
import { builtinTableStyle } from '../src/xlsx/tableStyles/builtins'

const bytes = (name: string) => new Uint8Array(readFileSync(resolve(process.cwd(), `tests/fixtures/xlsx-pivots/${name}.xlsx`)))

async function remoteLayer(name: string, grantBytes = 1024 * 1024, override?: Uint8Array) {
  const input = override ?? bytes(name)
  const session = await openImportSession({ kind: 'blob', blob: new Blob([new Uint8Array(input).buffer as ArrayBuffer]) }, { docId: `table-style-${name}`, profile: 'D' })
  try {
    await session.done
    const source = createPagedWorkbookSource(session), sheet = source.sheets[0], store = source.store(sheet.sheetId)
    const result = await (store as unknown as { readRegionStyles(options: { signal: AbortSignal; stamp: typeof store.stamp }, grant: number, memory: { set(bytes: number): boolean }): Promise<RemoteRegionStyles> }).readRegionStyles({ signal: new AbortController().signal, stamp: store.stamp }, grantBytes, { set: () => true })
    return { sheet, result, release: () => source.release(), session }
  } catch (error) { await session.release(); throw error }
}

function render(sheet: Parameters<typeof computeMetrics>[0]): Buffer {
  const metrics = computeMetrics(sheet), canvas = createCanvas(metrics.widthPx, metrics.heightPx)
  renderSheet(sheet, canvas.getContext('2d') as unknown as CanvasRenderingContext2D, metrics)
  return Buffer.from(canvas.getContext('2d').getImageData(0, 0, canvas.width, canvas.height).data.buffer)
}

describe('paged table-style transfer', () => {
  it.each(['table-medium-2', 'table-totals-row', 'table-custom-dxf'])(`matches resident pixels for %s`, async name => {
    let input = bytes(name)
    if (name === 'table-custom-dxf') {
      const zip = await JSZip.loadAsync(input), styles = await zip.file('xl/styles.xml')!.async('string')
      const customStyle = '<tableStyle name="OfficeViewCustom" pivot="0" table="1" count="4"><tableStyleElement type="wholeTable" dxfId="0"/><tableStyleElement type="headerRow" dxfId="0"/><tableStyleElement type="firstRowStripe" dxfId="0" size="2"/><tableStyleElement type="secondRowStripe" dxfId="0" size="1"/></tableStyle>'
      zip.file('xl/styles.xml', styles.replace(/<tableStyle name="OfficeViewCustom"[^>]*>.*?<\/tableStyle>/, customStyle))
      input = new Uint8Array(await zip.generateAsync({ type: 'uint8array' }))
    }
    const source = await remoteLayer(name, 1024 * 1024, input)
    try {
      expect(source.result.tables.length).toBeGreaterThan(0)
      const resident = await parseXlsx(await OfficePackage.load(input), { valuePolicy: 'cached' })
      const tableStyles = new Map(source.result.tables.map(item => [item.def.name, item.def]))
      const extras: StylesExtras = { dxfs: [], tableStyles, slicerStyles: new Map() }
      const residentSheet = resident.sheets[0]!
      const pagedSheet = { ...residentSheet, tables: source.result.tables.map(item => ({ ...item.table, name: item.table.displayName, partPath: '', columns: [] })), regionStyles: buildTableRegions(source.result.tables.map(item => item.table), extras, source.result.theme!), regionStylesResolved: true }
      expect(render(pagedSheet)).toEqual(render(residentSheet))
      expect(pagedSheet.regionStyles?.rects.length).toBe(source.result.tables.length)
    } finally { await source.release(); await source.session.release() }
  })

  it('refuses a too-small receiver grant with an explicit diagnostic and no styled layer', async () => {
    const source = await remoteLayer('table-medium-2', 1)
    try {
      expect(source.result.tables).toEqual([])
      expect(source.result.diagnostics).toContainEqual(expect.objectContaining({ feature: 'table-style-reply-limit' }))
      const resident = await parseXlsx(await OfficePackage.load(bytes('table-medium-2')), { valuePolicy: 'cached' })
      const noLayer = { ...resident.sheets[0]!, tables: [], regionStyles: undefined, regionStylesResolved: true }
      expect(render(noLayer)).toEqual(render({ ...resident.sheets[0]!, tables: undefined, tableStyleSource: undefined, regionStyles: undefined, regionStylesResolved: true }))
    } finally { await source.release(); await source.session.release() }
  })

  it('releases the granted reply and decoded layer on sheet and document close', async () => {
    let released = 0, decoded = 0
    const owner: CacheIdentity = { docId: 'cache', sheetId: 's', valueGen: 0, fontGen: 0, viewGen: 0 }
    const regionTables: RemoteRegionStyles['tables'] = [{ table: { id: 't', displayName: 'T', sheetId: 's', extent: { sheetId: 's', firstRow: 0, firstCol: 0, rows: 5, cols: 2 }, headerRowCount: 1, totalsRowCount: 0, styleInfo: { name: 'TableStyleMedium2', showFirstColumn: false, showLastColumn: false, showRowStripes: true, showColumnStripes: false } }, def: builtinTableStyle('TableStyleMedium2')! }]
    const response = (stamp = owner): WorkerReply<'read'> => ({ kind: 'complete', commandKind: 'read', requestId: 'r', owner, diagnostics: [], replyBytes: 128,
      result: { kind: 'read', result: { op: 'regionStyles', result: { sheetId: 's', stamp, tables: regionTables, cells: new Int32Array(), diagnostics: [] } } } } as unknown as WorkerReply<'read'>)
    const host = {
      identity: () => owner, hasRegionStyles: () => true, loadState: () => ({ completeness: 'complete' as const }), grant: () => 4096, releaseGrant: () => {},
      request: async () => response(), nextRequestId: () => 'r', release: () => { released++ },
      admitLayer: (_sheet: string, _request: string, bytes: number) => { decoded += bytes; return () => { decoded -= bytes } }, changed: () => {},
    }
    const cache = new RemoteRegionStyleCache(host, 8 * 1024 * 1024)
    cache.get('s')
    await new Promise(resolve => setTimeout(resolve, 0))
    expect(cache.get('s').state).toBe('ready')
    expect(cache.get('s').layer?.rects).toHaveLength(1)
    expect(released).toBe(0); expect(decoded).toBeGreaterThan(0)
    cache.closeSheet('s')
    expect(released).toBe(1); expect(decoded).toBe(0)
    cache.get('s')
    await new Promise(resolve => setTimeout(resolve, 0))
    cache.close()
    expect(released).toBe(2); expect(decoded).toBe(0)
  })

  it('discards an old-stamp response and refunds its reply charge', async () => {
    const owner: CacheIdentity = { docId: 'stale', sheetId: 's', valueGen: 2, fontGen: 0, viewGen: 0 }
    let released = 0
    const old: CacheIdentity = { ...owner, valueGen: 1 }
    const reply = { kind: 'complete', commandKind: 'read', requestId: 'r', owner, diagnostics: [], replyBytes: 128,
      result: { kind: 'read', result: { op: 'regionStyles', result: { sheetId: 's', stamp: old, tables: [], cells: new Int32Array(), diagnostics: [] } } } } as unknown as WorkerReply<'read'>
    const cache = new RemoteRegionStyleCache({ identity: () => owner, hasRegionStyles: () => true, loadState: () => ({ completeness: 'complete' }), grant: () => 4096, releaseGrant: () => {}, request: async () => reply, nextRequestId: () => 'r', release: () => { released++ }, admitLayer: () => () => {}, changed: () => {} }, 8 * 1024 * 1024)
    cache.get('s'); await new Promise(resolve => setTimeout(resolve, 0))
    expect(released).toBe(1)
    expect(cache.retainedEntries).toBe(0)
    expect(cache.retainedBytes).toBe(0)
    cache.close()
  })

  it('aborts a pending read and releases its receiver grant on sheet close', async () => {
    const owner: CacheIdentity = { docId: 'abort', sheetId: 's', valueGen: 0, fontGen: 0, viewGen: 0 }
    let grantHeld = true
    const cache = new RemoteRegionStyleCache({ identity: () => owner, hasRegionStyles: () => true, loadState: () => ({ completeness: 'complete' }), grant: () => 4096, releaseGrant: () => { grantHeld = false },
      request: (_command, signal) => new Promise(resolve => signal!.addEventListener('abort', () => { grantHeld = false; resolve(Symbol('cancelled')) }, { once: true })),
      nextRequestId: () => 'r', release: () => {}, admitLayer: () => () => {}, changed: () => {} }, 8 * 1024 * 1024)
    cache.get('s'); await Promise.resolve()
    cache.closeSheet('s'); await new Promise(resolve => setTimeout(resolve, 0))
    expect(grantHeld).toBe(false)
    expect(cache.retainedEntries).toBe(0)
    cache.close()
  })

  it('sends a real RemoteWorkbook grant and settles a late reply after sheet close', async () => {
    const listeners = new Set<(event: MessageEvent) => void>()
    let read: Extract<WorkerCommand, { kind: 'read' }> | undefined
    const port = {
      postMessage(command: WorkerCommand) { if (command.kind === 'read' && command.query.op === 'regionStyles') read = command },
      addEventListener(_type: string, listener: (event: MessageEvent) => void) { listeners.add(listener) },
      removeEventListener(_type: string, listener: (event: MessageEvent) => void) { listeners.delete(listener) },
    }
    const client = new ViewerClient(port as never, 'region-route')
    const remote = new RemoteWorkbook(client, { docId: 'region-route', valuePolicy: 'cached', sheets: [{ sheetId: 's', name: 'S' }], diagnostics: [], layouts: new Map([['s', { rowOverrides: [], colOverrides: [], hidden: [], hasRegionStyles: true }]]) })
    const internals = remote as unknown as { completed: Set<string>; structures: Map<string, { gen: number; merges: never[]; state: 'done' }> }
    internals.completed.add('s'); internals.structures.set('s', { gen: 0, merges: [], state: 'done' })
    remote.regionStyles('s')
    await Promise.resolve()
    expect(read?.query.op).toBe('regionStyles')
    if (read?.query.op !== 'regionStyles') throw new Error('region-style read was not dispatched')
    const grant = read.query.grantBytes
    expect(grant).toBeGreaterThan(0)
    remote.closeSheet('s')
    const payload: RemoteRegionStyles = { sheetId: 's', stamp: read.owner, tables: [], cells: new Int32Array(), diagnostics: [] }
    const reply = { kind: 'complete', commandKind: 'read', requestId: read.requestId, owner: read.owner, diagnostics: [], replyBytes: 128,
      result: { kind: 'read', result: { op: 'regionStyles', result: payload } } } as unknown as WorkerReply<'read'>
    for (const listener of listeners) listener({ data: reply } as MessageEvent)
    await new Promise(resolve => setTimeout(resolve, 0))
    expect(remote.mainLedger.snapshot().byPool.replies).toBe(0)
    expect(remote.mainLedger.ownerBytes('replies', `region-style-grant:${read.requestId}`)).toBe(0)
    await remote.close()
  })
})
