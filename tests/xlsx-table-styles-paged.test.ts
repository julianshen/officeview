// @vitest-environment node
import { readFileSync } from 'node:fs'
import { resolve } from 'node:path'
import { createCanvas } from 'canvas'
import JSZip from 'jszip'
import { describe, expect, it, vi } from 'vitest'
import { OfficePackage } from '../src/core/zip'
import { parseXlsx } from '../src/xlsx/parse'
import { computeMetrics, renderSheet } from '../src/xlsx/render'
import { openImportSession } from '../src/xlsx/storage/pagedImport'
import { createPagedWorkbookSource } from '../src/xlsx/view/pagedSource'
import type { RemoteRegionStyles } from '../src/xlsx/view/contracts'
import type { CacheIdentity, WorkerReply } from '../src/xlsx/view/contracts'
import { buildTableRegions } from '../src/xlsx/tableStyles/regionLayer'
import type { StylesExtras } from '../src/xlsx/tableStyles/model'
import { RemoteRegionStyleCache, sheetWithRegionStyles } from '../src/worker/viewerRemoteRegionStyles'
import { ViewerClient } from '../src/worker/viewerClient'
import { RemoteWorkbook } from '../src/worker/viewerRemoteSource'
import type { WorkerCommand } from '../src/xlsx/view/contracts'
import { builtinTableStyle } from '../src/xlsx/tableStyles/builtins'
import { ViewerKernel, connectInProcess } from '../src/worker/viewerKernel'
import { openRemoteWorkbook } from '../src/worker/viewerRemoteSource'
import { buildSheetGeometry } from '../src/xlsx/view/viewport'
import { renderSheetViewport } from '../src/xlsx/view/renderViewport'
import * as tableStyleLoader from '../src/xlsx/tableStyleLoader'

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

  it('defers and contains a throwing change observer without releasing a reply twice', async () => {
    const owner: CacheIdentity = { docId: 'observer', sheetId: 's', valueGen: 0, fontGen: 0, viewGen: 0 }
    let released = 0, changed = 0
    const cache = new RemoteRegionStyleCache({
      identity: () => owner, hasRegionStyles: () => true, loadState: () => ({ completeness: 'complete' }), grant: () => 4096, releaseGrant: () => {},
      request: async () => ({ kind: 'complete', commandKind: 'read', requestId: 'r', owner, diagnostics: [], replyBytes: 64,
        result: { kind: 'read', result: { op: 'regionStyles', result: { sheetId: 's', stamp: owner, tables: [], cells: new Int32Array(), diagnostics: [] } } } } as unknown as WorkerReply<'read'>),
      nextRequestId: () => 'r', release: () => { released++ }, admitLayer: () => () => {}, changed: () => { changed++; throw new Error('observer') },
    }, 4096)
    cache.get('s')
    await Promise.resolve()
    await Promise.resolve()
    expect(changed).toBe(1)
    expect(released).toBe(0)
    cache.close()
    expect(released).toBe(1)
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

  it('evicts and rebuilds decoded layers so five mobile sheets paint in turn', async () => {
    const owners = new Map(Array.from({ length: 5 }, (_, i) => [`s${i}`, { docId: 'five', sheetId: `s${i}`, valueGen: 0, fontGen: 0, viewGen: 0 } satisfies CacheIdentity]))
    let requests = 0, builds = 0, evictions = 0
    const table = (sheetId: string): RemoteRegionStyles['tables'][number] => ({ table: { id: `t-${sheetId}`, displayName: 'T', sheetId, extent: { sheetId, firstRow: 0, firstCol: 0, rows: 5, cols: 2 }, headerRowCount: 1, totalsRowCount: 0, styleInfo: { name: 'TableStyleMedium2', showFirstColumn: false, showLastColumn: false, showRowStripes: true, showColumnStripes: false } }, def: builtinTableStyle('TableStyleMedium2')! })
    const host = {
      identity: (id: string) => owners.get(id)!, hasRegionStyles: () => true, loadState: () => ({ completeness: 'complete' as const }),
      grant: () => 4096, releaseGrant: () => {}, nextRequestId: () => `r${requests}`,
      request: async (command: { owner: CacheIdentity }) => {
        requests++
        const payload: RemoteRegionStyles = { sheetId: command.owner.sheetId, stamp: command.owner, tables: [table(command.owner.sheetId)], cells: new Int32Array(), diagnostics: [] }
        return { kind: 'complete', commandKind: 'read', requestId: `r${requests}`, owner: command.owner, diagnostics: [], replyBytes: 128,
          result: { kind: 'read', result: { op: 'regionStyles', result: payload } } } as unknown as WorkerReply<'read'>
      },
      release: () => {}, admitLayer: () => { builds++; return () => {} }, changed: (_id: string, reason?: string) => { if (reason === 'evicted') evictions++ },
    }
    const cache = new RemoteRegionStyleCache(host, 8 * 1024 * 1024)
    const flush = () => new Promise(resolve => setTimeout(resolve, 0))
    for (let round = 0; round < 2; round++) for (let i = 0; i < 5; i++) {
      const state = cache.get(`s${i}`)
      if (state.state !== 'ready' || !state.layer) await flush()
      let ready = cache.get(`s${i}`)
      for (let wait = 0; wait < 20 && !ready.layer; wait++) { await flush(); ready = cache.get(`s${i}`) }
      expect(ready.state).toBe('ready')
      expect(ready.layer?.rects).toHaveLength(1)
    }
    expect(evictions).toBeGreaterThan(0)
    expect(builds).toBeGreaterThan(5)
    expect(requests).toBe(5)
    cache.close()
  })

  it('does not request styles for a plain sheet and retries null grants and retryable failures', async () => {
    let plainRequests = 0
    const plain = new RemoteRegionStyleCache({ identity: () => ({ docId: 'plain', sheetId: 's', valueGen: 0, fontGen: 0, viewGen: 0 }), hasRegionStyles: () => false, loadState: () => ({ completeness: 'complete' }), grant: () => 4096, releaseGrant: () => {}, request: async () => { plainRequests++; return Symbol() }, nextRequestId: () => 'plain', release: () => {}, admitLayer: () => () => {}, changed: () => {} }, 1024)
    expect(plain.get('s').state).toBe('waiting')
    await new Promise(resolve => setTimeout(resolve, 0))
    expect(plainRequests).toBe(0)
    plain.close()

    const owner: CacheIdentity = { docId: 'retry', sheetId: 's', valueGen: 0, fontGen: 0, viewGen: 0 }
    let grants = 0, requests = 0
    const retry = new RemoteRegionStyleCache({ identity: () => owner, hasRegionStyles: () => true, loadState: () => ({ completeness: 'complete' }), grant: () => ++grants === 1 ? null : 4096, releaseGrant: () => {},
      request: async () => { requests++; return requests === 1 ? ({ kind: 'error', commandKind: 'read', requestId: 'retry', owner, code: 'stale-source', retryable: true, diagnostics: [] } as unknown as WorkerReply<'read'>) : ({ kind: 'complete', commandKind: 'read', requestId: 'retry', owner, diagnostics: [], replyBytes: 64, result: { kind: 'read', result: { op: 'regionStyles', result: { sheetId: 's', stamp: owner, tables: [], cells: new Int32Array(), diagnostics: [] } } } } as unknown as WorkerReply<'read'>) },
      nextRequestId: () => 'retry', release: () => {}, admitLayer: () => () => {}, changed: () => {} }, 8 * 1024 * 1024)
    retry.get('s')
    await new Promise(resolve => setTimeout(resolve, 900))
    expect(grants).toBeGreaterThanOrEqual(3)
    expect(requests).toBeGreaterThanOrEqual(1)
    expect(retry.get('s').state).toBe('ready')
    retry.close()
  })

  it('retries a cancelled result and a timed-out read instead of pinning failure', async () => {
    vi.useFakeTimers()
    try {
      const owner: CacheIdentity = { docId: 'cancel-timeout', sheetId: 's', valueGen: 0, fontGen: 0, viewGen: 0 }
      let requests = 0
      const success = (): WorkerReply<'read'> => ({ kind: 'complete', commandKind: 'read', requestId: `r${requests}`, owner, diagnostics: [], replyBytes: 64, result: { kind: 'read', result: { op: 'regionStyles', result: { sheetId: 's', stamp: owner, tables: [], cells: new Int32Array(), diagnostics: [] } } } } as unknown as WorkerReply<'read'>)
      const cache = new RemoteRegionStyleCache({ identity: () => owner, hasRegionStyles: () => true, loadState: () => ({ completeness: 'complete' }), grant: () => 4096, releaseGrant: () => {},
        request: (_command, signal) => {
          requests++
          if (requests === 1) return Promise.resolve(Symbol('cancelled'))
          if (requests === 2) return new Promise(resolve => signal!.addEventListener('abort', () => resolve(Symbol('timeout')), { once: true }))
          return Promise.resolve(success())
        }, nextRequestId: () => `r${requests}`, release: () => {}, admitLayer: () => () => {}, changed: () => {} }, 8 * 1024 * 1024)
      cache.get('s')
      await vi.advanceTimersByTimeAsync(250)
      await vi.advanceTimersByTimeAsync(10_000)
      await vi.advanceTimersByTimeAsync(500)
      expect(requests).toBe(3)
      expect(cache.get('s').state).toBe('ready')
      cache.close()
    } finally { vi.useRealTimers() }
  })

  it('has a custom built-in override ready before the first sheet completes', async () => {
    const zip = await JSZip.loadAsync(bytes('table-medium-2'))
    const styles = await zip.file('xl/styles.xml')!.async('string')
    const custom = '<dxfs count="1"><dxf><fill><patternFill patternType="solid"><fgColor rgb="FFFF0000"/></patternFill></fill></dxf></dxfs><tableStyles count="1" defaultTableStyle="TableStyleMedium2"><tableStyle name="TableStyleMedium2" pivot="0" table="1" count="1"><tableStyleElement type="wholeTable" dxfId="0"/></tableStyle></tableStyles>'
    const replaced = styles.replace(/<dxfs\b[^>]*(?:\/>|>[\s\S]*?<\/dxfs>)/, '').replace(/<tableStyles\b[^>]*(?:\/>|>[\s\S]*?<\/tableStyles>)/, '')
    zip.file('xl/styles.xml', replaced.replace('</styleSheet>', `${custom}</styleSheet>`))
    const workbookXml = await zip.file('xl/workbook.xml')!.async('string')
    zip.file('xl/workbook.xml', workbookXml.replace('</sheets>', '<sheet name="Later" sheetId="2" r:id="rId2"/></sheets>'))
    const rels = await zip.file('xl/_rels/workbook.xml.rels')!.async('string')
    zip.file('xl/_rels/workbook.xml.rels', rels.replace('</Relationships>', '<Relationship Id="rId2" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet" Target="worksheets/sheet2.xml"/></Relationships>'))
    const types = await zip.file('[Content_Types].xml')!.async('string')
    zip.file('[Content_Types].xml', types.replace('</Types>', '<Override PartName="/xl/worksheets/sheet2.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml"/></Types>'))
    zip.file('xl/worksheets/sheet2.xml', '<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"><sheetData/></worksheet>')
    const input = new Uint8Array(await zip.generateAsync({ type: 'uint8array' }))
    let session: Awaited<ReturnType<typeof openImportSession>> | undefined
    let releaseSecond!: () => void
    let blockedSecond!: () => void
    const secondGate = new Promise<void>(resolve => { releaseSecond = resolve })
    const secondStarted = new Promise<void>(resolve => { blockedSecond = resolve })
    session = await openImportSession({ kind: 'blob', blob: new Blob([input.buffer as ArrayBuffer]) }, { docId: 'style-race', profile: 'D', beforeRead: async () => {
      if (session?.workbook.sheets[0]?.state === 'complete' && session.workbook.sheets[1]?.state !== 'complete') { blockedSecond(); await secondGate }
    } })
    const source = createPagedWorkbookSource(session)
    try {
      await secondStarted
      expect(session.workbook.sheets[1]?.state).not.toBe('complete')
      const store = source.store(source.sheets[0]!.sheetId)
      const result = await store.readRegionStyles!({ signal: new AbortController().signal, stamp: store.stamp }, 64 * 1024, { set: () => true })
      const override = result.tables.find(item => item.def.name === 'TableStyleMedium2')?.def
      expect(override?.elements.wholeTable?.dxf.fill?.fg).toEqual({ rgb: 'FFFF0000' })
      expect(result.diagnostics).not.toContainEqual(expect.objectContaining({ feature: 'table-style-unknown' }))
    } finally { releaseSecond(); await session.done; await source.release(); await session.release() }
  })

  it('uses the real paged worker reply and receiver cache in viewport pixel parity', async () => {
    const input = bytes('table-medium-2')
    const session = await openImportSession({ kind: 'blob', blob: new Blob([new Uint8Array(input).buffer as ArrayBuffer]) }, { docId: 'table-integration', profile: 'D' })
    await session.done
    const pagedSource = createPagedWorkbookSource(session)
    const kernel = new ViewerKernel({ parse: async () => { throw new Error('expected the supplied paged source') }, openSource: async () => pagedSource }, { yield: async () => {} })
    const client = new ViewerClient(connectInProcess(kernel), 'table-integration')
    const opened = await openRemoteWorkbook(client, 'table-integration', new Blob([input]), 'D')
    if (!('layouts' in opened)) throw new Error(`paged open failed: ${JSON.stringify(opened)}`)
    const remote = new RemoteWorkbook(client, opened), sheetId = remote.sheets[0]!.sheetId
    try {
      remote.merges(sheetId)
      for (let i = 0; i < 100 && remote.loadState(sheetId).completeness !== 'complete'; i++) await new Promise(resolve => setTimeout(resolve, 2))
      expect(remote.layout(sheetId).hasRegionStyles).toBe(true)
      expect(remote.layout(sheetId).regionStyleTableCount).toBe(1)
      remote.regionStyles(sheetId)
      for (let i = 0; i < 100 && remote.regionStyles(sheetId).state !== 'ready'; i++) await new Promise(resolve => setTimeout(resolve, 2))
      const state = remote.regionStyles(sheetId)
      expect(state.state).toBe('ready')
      expect(state.payload?.tables).toHaveLength(1)
      expect(state.layer?.rects.length).toBeGreaterThan(0)
      const resident = (await parseXlsx(await OfficePackage.load(input), { valuePolicy: 'cached' })).sheets[0]!
      await tableStyleLoader.ensureTableStyleEngine([resident])
      const paged = sheetWithRegionStyles({ ...resident, tables: [] }, state)
      const paint = (sheet: typeof resident) => {
        const geometry = buildSheetGeometry(sheet)
        if (geometry.kind !== 'ready') throw new Error('fixture geometry should be ready')
        const canvas = createCanvas(320, 180), ctx = canvas.getContext('2d') as unknown as CanvasRenderingContext2D
        renderSheetViewport(sheet, ctx, geometry, { box: { x: 0, y: 0, width: 320, height: 180 }, zoom: 1, deviceScale: 1 })
        return Buffer.from(ctx.getImageData(0, 0, canvas.width, canvas.height).data)
      }
      expect(paint(paged)).toEqual(paint(resident))
    } finally { await remote.close(); await pagedSource.release(); await session.release() }
  })
})
