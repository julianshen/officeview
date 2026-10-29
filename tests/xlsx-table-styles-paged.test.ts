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
import { emptyWork, type ViewportPlan } from '../src/xlsx/view/viewport'
import { renderSheetViewport } from '../src/xlsx/view/renderViewport'
import { pagedSheetGeometry } from '../src/components/xlsx/pagedGeometry'
import { createIdentityOrder } from '../src/xlsx/view/order'
import { createIntervalSet, visibilityFromHidden } from '../src/xlsx/view/visibility'
import { validateOverflowProof } from '../src/xlsx/view/viewportOverflow'
import { formatA1 } from '../src/xlsx/view/coordinates'
import { pageTileKey } from '../src/worker/viewerRemoteSource'
import * as tableStyleLoader from '../src/xlsx/tableStyleLoader'
import { MAX_REGION_STYLE_MEMO_ENTRIES, REGION_STYLE_DESCRIPTOR_MAX_BYTES_PER_TABLE, REGION_STYLE_REPLY_OVERHEAD_BYTES } from '../src/xlsx/view/contracts'
import { estimatePayloadBytes } from '../src/xlsx/view/payloadBytes'

const bytes = (name: string) => new Uint8Array(readFileSync(resolve(process.cwd(), `tests/fixtures/xlsx-pivots/${name}.xlsx`)))

async function remoteLayer(name: string, grantBytes = 1024 * 1024, override?: Uint8Array) {
  const input = override ?? bytes(name)
  const session = await openImportSession({ kind: 'blob', blob: new Blob([new Uint8Array(input).buffer as ArrayBuffer]) }, { docId: `table-style-${name}`, profile: 'D' })
  try {
    await session.done
    const source = createPagedWorkbookSource(session), sheet = source.sheets[0], store = source.store(sheet.sheetId)
    const result = await (store as unknown as { readRegionStyles(options: { signal: AbortSignal; stamp: typeof store.stamp }, grant: number, memory: { set(bytes: number): boolean }): Promise<RemoteRegionStyles> }).readRegionStyles({ signal: new AbortController().signal, stamp: store.stamp }, grantBytes, { set: () => true })
    return { sheet, store, result, release: () => source.release(), session }
  } catch (error) { await session.release(); throw error }
}

function render(sheet: Parameters<typeof computeMetrics>[0]): Buffer {
  const metrics = computeMetrics(sheet), canvas = createCanvas(metrics.widthPx, metrics.heightPx)
  renderSheet(sheet, canvas.getContext('2d') as unknown as CanvasRenderingContext2D, metrics)
  return Buffer.from(canvas.getContext('2d').getImageData(0, 0, canvas.width, canvas.height).data.buffer)
}

describe('paged table-style transfer', () => {
  it('round-trips every built-in table style within the receiver-computed descriptor grant', async () => {
    const files = ['table.ts', 'pivot.ts']
    const allNames = [...new Set(files.flatMap(file => [...readFileSync(resolve(process.cwd(), `src/xlsx/tableStyles/builtins/${file}`), 'utf8').matchAll(/\"((?:Table|Pivot)Style(?:Light|Medium|Dark)\d+)\":/g)].map(match => match[1])))]
    expect(allNames).toHaveLength(144)
    const source = await remoteLayer('table-medium-2')
    try {
      const sheet = source.session.workbook.sheets[0]!
      const table = sheet.tables![0]!
      const store = source.store
      const grant = REGION_STYLE_REPLY_OVERHEAD_BYTES + REGION_STYLE_DESCRIPTOR_MAX_BYTES_PER_TABLE
      const stamp = store.stamp
      const read = (name: string) => {
        table.styleInfo!.name = name
        return store.readRegionStyles!({ signal: new AbortController().signal, stamp }, grant, { set: () => true })
      }
      const theme = source.result.theme!
      const probeTable = { ...source.result.tables[0]!.table }
      const sizes = allNames.map(name => estimatePayloadBytes({ sheetId: sheet.sheetId, stamp, tables: [{ table: { ...probeTable, styleInfo: { ...probeTable.styleInfo, name } }, def: builtinTableStyle(name)! }], theme, cells: new Int32Array(), diagnostics: [] }))
      const maximum = Math.max(...sizes)
      expect(maximum).toBeLessThan(REGION_STYLE_DESCRIPTOR_MAX_BYTES_PER_TABLE)
      const tableNames = allNames.filter(name => name.startsWith('TableStyle'))
      expect(tableNames).toHaveLength(60)
      for (const name of tableNames) {
        const result = await read(name)
        expect(result.tables, name).toHaveLength(1)
        expect(result.diagnostics, name).not.toContainEqual(expect.objectContaining({ feature: 'table-style-reply-limit' }))
      }
      expect(MAX_REGION_STYLE_MEMO_ENTRIES).toBe(4096)
    } finally { await source.release(); await source.session.release() }
  })

  it('retries once with the worker-reported size for an oversized custom descriptor', async () => {
    const zip = await JSZip.loadAsync(bytes('table-medium-2'))
    const styles = await zip.file('xl/styles.xml')!.async('string')
    const longFontName = 'CustomFont'.repeat(3_000)
    const custom = `<dxfs count="1"><dxf><font><name val="${longFontName}"/></font></dxf></dxfs><tableStyles count="1" defaultTableStyle="TableStyleMedium2"><tableStyle name="TableStyleMedium2" pivot="0" table="1" count="1"><tableStyleElement type="wholeTable" dxfId="0"/></tableStyle></tableStyles>`
    const replaced = styles.replace(/<dxfs\b[^>]*(?:\/>|>[\s\S]*?<\/dxfs>)/, '').replace(/<tableStyles\b[^>]*(?:\/>|>[\s\S]*?<\/tableStyles>)/, '')
    zip.file('xl/styles.xml', replaced.replace('</styleSheet>', `${custom}</styleSheet>`))
    const input = new Uint8Array(await zip.generateAsync({ type: 'uint8array' }))
    const session = await openImportSession({ kind: 'blob', blob: new Blob([input.buffer as ArrayBuffer]) }, { docId: 'large-style-retry', profile: 'D' })
    await session.done
    const pagedSource = createPagedWorkbookSource(session)
    const kernel = new ViewerKernel({ parse: async () => { throw new Error('expected supplied paged source') }, openSource: async () => pagedSource }, { yield: async () => {} })
    const client = new ViewerClient(connectInProcess(kernel), 'large-style-retry')
    const opened = await openRemoteWorkbook(client, 'large-style-retry', new Blob([input]), 'D')
    if (!('layouts' in opened)) throw new Error(`paged open failed: ${JSON.stringify(opened)}`)
    const remote = new RemoteWorkbook(client, opened), sheetId = remote.sheets[0]!.sheetId
    try {
      remote.regionStyles(sheetId)
      for (let i = 0; i < 100 && remote.regionStyles(sheetId).state !== 'ready' && remote.regionStyles(sheetId).state !== 'failed'; i++) await new Promise(resolve => setTimeout(resolve, 2))
      const state = remote.regionStyles(sheetId)
      expect(state.state).toBe('ready')
      expect(state.payload?.tables[0]?.def.elements.wholeTable?.dxf.font?.name).toBe(longFontName)
      expect(remote.stats.commands['read:regionStyles']).toBe(2)
    } finally { await remote.close(); await pagedSource.release(); await session.release() }
  })

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

  it('keeps a ready style layer across sort and font-generation changes without an unstyled transition', async () => {
    let owner: CacheIdentity = { docId: 'generation', sheetId: 's', valueGen: 4, fontGen: 0, viewGen: 0 }
    let requests = 0, changes = 0
    const tables: RemoteRegionStyles['tables'] = [{ table: { id: 't', displayName: 'T', sheetId: 's', extent: { sheetId: 's', firstRow: 0, firstCol: 0, rows: 5, cols: 2 }, headerRowCount: 1, totalsRowCount: 0, styleInfo: { name: 'TableStyleMedium2', showFirstColumn: false, showLastColumn: false, showRowStripes: true, showColumnStripes: false } }, def: builtinTableStyle('TableStyleMedium2')! }]
    const cache = new RemoteRegionStyleCache({
      identity: () => owner, hasRegionStyles: () => true, loadState: () => ({ completeness: 'complete' }), grant: () => 4096, releaseGrant: () => {},
      request: async command => {
        requests++
        const stamp = command.owner as CacheIdentity
        return { kind: 'complete', commandKind: 'read', requestId: 'generation', owner: stamp, diagnostics: [], replyBytes: 128,
          result: { kind: 'read', result: { op: 'regionStyles', result: { sheetId: 's', stamp, tables, cells: new Int32Array(), diagnostics: [] } } } } as unknown as WorkerReply<'read'>
      },
      nextRequestId: () => 'generation', release: () => {}, admitLayer: () => () => {}, changed: () => { changes++ },
    }, 8 * 1024 * 1024)
    try {
      cache.get('s')
      let ready = cache.get('s')
      for (let i = 0; i < 100 && ready.state !== 'ready'; i++) { await new Promise(resolve => setTimeout(resolve, 2)); ready = cache.get('s') }
      expect(ready.state).toBe('ready')
      const layer = ready.layer, generation = ready.gen, initialChanges = changes
      expect(layer?.rects).toHaveLength(1)

      owner = { ...owner, viewGen: owner.viewGen + 1 } // committed sort
      const sorted = cache.get('s')
      expect(sorted.state).toBe('ready')
      expect(sorted.layer).toBe(layer)
      expect(sorted.gen).toBe(generation)
      expect(changes).toBe(initialChanges)

      owner = { ...owner, fontGen: owner.fontGen + 1 }
      const refonted = cache.get('s')
      expect(refonted.state).toBe('ready')
      expect(refonted.layer).toBe(layer)
      expect(refonted.gen).toBe(generation)
      expect(changes).toBe(initialChanges)
      expect(requests).toBe(1)
    } finally { cache.close() }
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

  it('retries a cancellation reply emitted by ViewerKernel and then loads the style layer', async () => {
    const input = bytes('table-medium-2')
    const session = await openImportSession({ kind: 'blob', blob: new Blob([input.buffer as ArrayBuffer]) }, { docId: 'kernel-style-cancel', profile: 'D' })
    await session.done
    const pagedSource = createPagedWorkbookSource(session), sheetId = pagedSource.sheets[0]!.sheetId
    const store = pagedSource.store(sheetId), original = store.readRegionStyles!.bind(store)
    let reads = 0, entered!: () => void, kernelKind: string | undefined
    const firstRead = new Promise<void>(resolve => { entered = resolve })
    store.readRegionStyles = async (options, grant, memory) => {
      if (reads++ === 0) {
        entered()
        await new Promise<void>(resolve => options.signal.addEventListener('abort', () => resolve(), { once: true }))
      }
      return original(options, grant, memory)
    }
    const kernel = new ViewerKernel({ parse: async () => { throw new Error('expected supplied source') }, openSource: async () => pagedSource }, { yield: async () => {} })
    const client = new ViewerClient(connectInProcess(kernel), 'kernel-style-cancel')
    const opened = await openRemoteWorkbook(client, 'kernel-style-cancel', new Blob([input]), 'D')
    if (!('layouts' in opened)) throw new Error(`paged open failed: ${JSON.stringify(opened)}`)
    const remote = new RemoteWorkbook(client, opened)
    let requests = 0
    const cache = new RemoteRegionStyleCache({
      identity: id => remote.identity(id), hasRegionStyles: id => remote.layout(id).hasRegionStyles === true,
      loadState: id => remote.loadState(id), grant: () => 128 * 1024, releaseGrant: () => {},
      request: async command => {
        requests++
        const result = client.send(command)
        if (requests === 1) {
          await firstRead
          await client.send({ kind: 'cancel', requestId: client.nextRequestId(), targetRequestId: command.requestId })
        }
        const reply = await result
        if (requests === 1) kernelKind = reply.kind
        return reply
      }, nextRequestId: () => client.nextRequestId(), release: () => {}, admitLayer: () => () => {}, changed: () => {},
    }, 8 * 1024 * 1024)
    try {
      cache.get(sheetId)
      for (let i = 0; i < 200 && cache.get(sheetId).state !== 'ready'; i++) await new Promise(resolve => setTimeout(resolve, 5))
      expect(kernelKind).toBe('cancelled')
      expect(requests).toBe(2)
      expect(cache.get(sheetId).state).toBe('ready')
    } finally { cache.close(); await remote.close(); await pagedSource.release(); await session.release() }
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

  it('paints worker table header, stripe, and empty in-region cells from a paged viewport plan', async () => {
    const archive = await JSZip.loadAsync(bytes('table-medium-2'))
    const sheetXml = await archive.file('xl/worksheets/sheet1.xml')!.async('string')
    expect(sheetXml).toMatch(/<c r="C2"[^>]*>/)
    archive.file('xl/worksheets/sheet1.xml', sheetXml.replace(/<c r="C2"[^>]*>[\s\S]*?<\/c>/, ''))
    const input = new Uint8Array(await archive.generateAsync({ type: 'uint8array' }))
    const session = await openImportSession({ kind: 'blob', blob: new Blob([new Uint8Array(input).buffer as ArrayBuffer]) }, { docId: 'table-integration', profile: 'D' })
    await session.done
    const pagedSource = createPagedWorkbookSource(session)
    const kernel = new ViewerKernel({ parse: async () => { throw new Error('expected the supplied paged source') }, openSource: async () => pagedSource }, { yield: async () => {} })
    const port = connectInProcess(kernel)
    const client = new ViewerClient(port, 'table-integration')
    const opened = await openRemoteWorkbook(client, 'table-integration', new Blob([new Uint8Array(input).buffer as ArrayBuffer]), 'D')
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
      const owner = remote.identity(sheetId), axes = remote.geometry(sheetId)
      const geometry = pagedSheetGeometry(remote, sheetId, createIdentityOrder(), visibilityFromHidden(createIntervalSet([])))
      if (geometry.kind !== 'ready') throw new Error('paged geometry should be ready')
      const window = { box: { x: 0, y: 0, width: 480, height: 400 }, zoom: 1, deviceScale: 1 }
      const extent = { width: axes.cols.totalCss(), height: axes.rows.totalCss() }
      const frame = remote.currentPageFrame, tile = pageTileKey(sheetId, window.box), pageLimit = remote.pageLimits.tile
      expect(remote.admitPages(frame, tile, pageLimit, false, false)).toBe(true)
      expect(remote.admitPages(frame, tile, pageLimit, false)).toBe(true)
      const out = await remote.viewport(owner, window, new AbortController().signal, 'tile')
      expect(typeof out).toBe('object')
      if (typeof out !== 'object' || out.kind !== 'complete') throw new Error('worker viewport did not complete')
      const result = out.result
      if (result.content.kind !== 'detail') throw new Error('worker viewport did not return a detail plan')
      const proof = validateOverflowProof(result, owner, window, 'D', 'tile', remote.columns(sheetId))
      const cells = result.content.cells.flatMap(item => {
        if (item.kind !== 'cell' && item.kind !== 'unavailable') return []
        const overflow = proof.get(item.source.row * 16384 + item.source.col)!
        const style = remote.styleOf(item.cell.styleIndex)
        return [{ source: item.source, x: item.bounds.x, y: item.bounds.y, width: item.bounds.width, height: item.bounds.height, overflow: { left: overflow.left, right: overflow.right }, cell: { ...item.cell, ref: formatA1(item.source), ...(style ? { style } : {}) } }]
      })
      const plan: Extract<ViewportPlan, { kind: 'detail' }> = { kind: 'detail', cells, merges: (result.merges ?? []).map(merge => ({ range: merge.range, ...merge.bounds })), drawings: [], rowLines: result.rowLines ?? [], colLines: result.colLines ?? [], extent, diagnostics: [], coarse: result.coarse ?? [], work: result.work ?? emptyWork() }
      const canvas = createCanvas(window.box.width, window.box.height), ctx = canvas.getContext('2d') as unknown as CanvasRenderingContext2D
      const painted = renderSheetViewport(paged, ctx, geometry, window, undefined, { preparedPlan: plan })
      const at = (row: number, col: number) => painted.records.find(record => record.source.row === row && record.source.col === col)
      expect(at(0, 0)?.fill.color, 'table header fill').toBeTruthy()
      expect(at(1, 0)?.fill.color, 'table stripe fill').toBeTruthy()
      expect(at(1, 2)?.fill.color, 'synthetic empty in-region cell fill').toBeTruthy()
      expect(plan.cells.some(cell => cell.source.row === 1 && cell.source.col === 2)).toBe(false)
      remote.mainLedger.settle(out)
    } finally { await remote.close(); await pagedSource.release(); await session.release() }
  })

  it('retries an incomplete worker reply under the bounded backoff', async () => {
    vi.useFakeTimers()
    try {
      const owner: CacheIdentity = { docId: 'incomplete-style', sheetId: 's', valueGen: 0, fontGen: 0, viewGen: 0 }
      let requests = 0, released = 0
      const complete: WorkerReply<'read'> = { kind: 'complete', commandKind: 'read', requestId: 'r2', owner, diagnostics: [], replyBytes: 64,
        result: { kind: 'read', result: { op: 'regionStyles', result: { sheetId: 's', stamp: owner, tables: [], cells: new Int32Array(), diagnostics: [] } } } } as unknown as WorkerReply<'read'>
      const cache = new RemoteRegionStyleCache({ identity: () => owner, hasRegionStyles: () => true, loadState: () => ({ completeness: 'complete' }), grant: () => 4096, releaseGrant: () => {},
        request: async () => ++requests === 1
          ? ({ kind: 'incomplete', commandKind: 'read', requestId: 'r1', owner, missing: [], diagnostics: [] } as unknown as WorkerReply<'read'>)
          : complete,
        nextRequestId: () => `r${requests + 1}`, release: () => { released++ }, admitLayer: () => () => {}, changed: () => {} }, 4096)
      cache.get('s')
      await vi.advanceTimersByTimeAsync(250)
      await vi.runOnlyPendingTimersAsync()
      expect(requests).toBe(2)
      expect(cache.get('s').state).toBe('ready')
      expect(released).toBe(1)
      cache.close()
    } finally { vi.useRealTimers() }
  })
})
