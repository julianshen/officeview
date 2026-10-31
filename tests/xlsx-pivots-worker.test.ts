// @vitest-environment node
import { readFileSync } from 'node:fs'
import { fileURLToPath } from 'node:url'
import { resolve } from 'node:path'
import { describe, expect, it, vi } from 'vitest'
import { overlayUsedCounts, overlayUsedExtent, PIVOT_FIT_REPLY_BYTES, type WorkerCommand, type WorkerReply } from '../src/xlsx/view/contracts'
import { OfficePackage } from '../src/core/zip'
import { parseXlsx } from '../src/xlsx/parse'
import { ViewerKernel } from '../src/worker/viewerKernel'
import { createObjectWorkbookSource } from '../src/worker/viewerObjectSource'
import { ViewerClient } from '../src/worker/viewerClient'
import { connectInProcess } from '../src/worker/viewerKernel'
import { RemotePivotSessions } from '../src/worker/viewerRemotePivots'
import { pivotRuntimeFor } from '../src/xlsx/pivotLoader'
import { openImportSession } from '../src/xlsx/storage/pagedImport'
import { createPagedWorkbookSource } from '../src/xlsx/view/pagedSource'
import { sheetCellReader } from '../src/xlsx/pivots/engine/faithful'
import { openPivotSession } from '../src/xlsx/pivots/session/session'
import { ORACLE_VERIFIED_KINDS } from '../src/xlsx/pivots/session/session'
import { pivotChartGate } from '../src/xlsx/pivots/session/charts'
import type { PivotDelta } from '../src/xlsx/pivots/session/delta'
import { RECORD_LIMITS } from '../src/xlsx/pivots/engine/records'
import { pivotFitPayload } from '../src/xlsx/pivots/session/autofit'
import { decodeParts, loadPivotParts } from './helpers/xlsx-pivots-load'
import { RemoteWorkbook, openRemoteWorkbook } from '../src/worker/viewerRemoteSource'
import { pagedSheetGeometry } from '../src/components/xlsx/pagedGeometry'
import { createIdentityOrder } from '../src/xlsx/view/order'
import { createIntervalSet, visibilityFromHidden } from '../src/xlsx/view/visibility'
import { MemoryPool } from '../src/xlsx/view/tileCache'

const FIXTURE = resolve(fileURLToPath(new URL('./fixtures/xlsx-pivots/engine/', import.meta.url)), 'baseline.xlsx')
const FIXTURE_ROOT = resolve(fileURLToPath(new URL('./fixtures/xlsx-pivots/', import.meta.url)))
const TASK25_CASES = (JSON.parse(readFileSync(resolve(FIXTURE_ROOT, 'interaction/manifest.json'), 'utf8')) as Array<{
  case: string; file: string; task: number; status?: string; setup?: string; seed: string; base?: string; pivot: string; delta: PivotDelta[]
}>).filter(entry => entry.task === 25 && entry.status === 'built' && entry.setup === undefined && entry.delta.length > 0)
const SEEDS: Record<string, string> = {
  'engine-seed': resolve(FIXTURE_ROOT, 'recipes/engine-seed.xlsx'),
  'chart-seed': resolve(FIXTURE_ROOT, 'recipes/pivot-chart-seed.xlsx'),
  'getpivotdata-seed': resolve(FIXTURE_ROOT, 'interaction/getpivotdata-seed.xlsx'),
  'refresh-seed': resolve(FIXTURE_ROOT, 'interaction/refresh-seed.xlsx'),
}
const quiet = { yield: async () => {} }
let serial = 0
const id = () => `pivot-worker-${++serial}`
async function pivotAssets(source: ReturnType<typeof createPagedWorkbookSource>, sheetId: string) {
  if (!source.pivotAssets) throw new Error('Paged source does not expose pivot assets')
  return source.pivotAssets(sheetId, new AbortController().signal, 'D')
}
async function run(kernel: ViewerKernel, command: WorkerCommand): Promise<WorkerReply> {
  const replies: WorkerReply[] = []
  await kernel.handle(command, reply => replies.push(reply))
  expect(replies).toHaveLength(1)
  return replies[0]!
}
const completed = (reply: WorkerReply) => {
  if (reply.kind !== 'complete') throw new Error(`Expected complete, got ${JSON.stringify(reply)}`)
  return reply.result
}
async function openObjectPivot(options: { afterDispatch?: (command: WorkerCommand) => void | Promise<void>; duplicate?: boolean } = {}) {
  const doc = await parseXlsx(await OfficePackage.load(new Uint8Array(readFileSync(FIXTURE))))
  const sheet = doc.sheets.find(candidate => candidate.pivotTables?.length)!
  if (options.duplicate) sheet.pivotTables = [sheet.pivotTables![0]!, { ...sheet.pivotTables![0]!, name: 'Second Pivot Same Cache' }]
  const kernel = new ViewerKernel({ parse: async bytes => parseXlsx(await OfficePackage.load(bytes)), afterDispatch: options.afterDispatch }, quiet)
  const docId = id(), sheetId = sheet.sheetId ?? '1'
  await run(kernel, { kind: 'open', requestId: id(), docId, source: { kind: 'object', document: doc }, profile: 'D' })
  const session = kernel.session(docId)!
  const beforePivotOpen = session.memoryStats().sourceBytes
  const opened = await run(kernel, { kind: 'pivotOpen', requestId: id(), docId, sheetId, pivot: 0, grantBytes: 16 * 1024 * 1024 })
  expect(completed(opened)).toMatchObject({ kind: 'pivotOpen' })
  return { doc, sheet, kernel, docId, sheetId, session, beforePivotOpen }
}

describe('worker pivot protocol', () => {
  it('does not scan a no-pivot workbook during paged/open source initialization', async () => {
    const doc = await parseXlsx(await OfficePackage.load(new Uint8Array(readFileSync(FIXTURE))))
    for (const sheet of doc.sheets) { sheet.pivotTables = []; sheet.pivotRegions = [] }
    const docId = id(), source = createObjectWorkbookSource(doc, docId)
    const scan = vi.fn(source.scanFormulaCellsOutsidePivots?.bind(source))
    source.scanFormulaCellsOutsidePivots = scan
    const kernel = new ViewerKernel({ parse: async bytes => parseXlsx(await OfficePackage.load(bytes)), openSource: async () => source }, quiet)
    const opened = completed(await run(kernel, { kind: 'open', requestId: id(), docId, source: { kind: 'paged', handleId: docId }, profile: 'D' }))
    expect(opened.kind).toBe('open')
    expect(scan).not.toHaveBeenCalled()
    await run(kernel, { kind: 'close', requestId: id(), docId })
  })

  it('scans formulas only at the first valid pivot commit and reports the cached result', async () => {
    const doc = await parseXlsx(await OfficePackage.load(new Uint8Array(readFileSync(FIXTURE))))
    const sheet = doc.sheets.find(candidate => candidate.pivotTables?.length)!
    const docId = id(), source = createObjectWorkbookSource(doc, docId)
    const scan = vi.fn(source.scanFormulaCellsOutsidePivots?.bind(source))
    source.scanFormulaCellsOutsidePivots = scan
    const kernel = new ViewerKernel({ parse: async bytes => parseXlsx(await OfficePackage.load(bytes)), openSource: async () => source }, quiet)
    const sheetId = sheet.sheetId ?? '1'
    await run(kernel, { kind: 'open', requestId: id(), docId, source: { kind: 'paged', handleId: docId }, profile: 'D' })
    expect(scan).not.toHaveBeenCalled()
    await run(kernel, { kind: 'pivotOpen', requestId: id(), docId, sheetId, pivot: 0, grantBytes: 16 * 1024 * 1024 })
    const staged = completed(await run(kernel, { kind: 'pivotApply', requestId: id(), docId, sheetId, pivot: 0, baseRevision: 0, revision: 1,
      delta: { kind: 'collapse', field: 0, item: 0, collapsed: true }, grantBytes: 16 * 1024 * 1024 }))
    if (staged.kind !== 'pivotApply') throw new Error(JSON.stringify(staged))
    const committed = completed(await run(kernel, { kind: 'pivotCommit', requestId: id(), docId, sheetId, pivot: 0, revision: staged.revision }))
    expect(scan).toHaveBeenCalledTimes(1)
    expect(committed).toMatchObject({ kind: 'pivotCommit', hasFormulaCellsOutsidePivots: true })
    await run(kernel, { kind: 'close', requestId: id(), docId })
  })

  it('open streams records under the profile cap and charges the worker pool', async () => {
    const doc = await parseXlsx(await OfficePackage.load(new Uint8Array(readFileSync(FIXTURE))))
    const sheet = doc.sheets.find(candidate => candidate.pivotTables?.length)!
    const kernel = new ViewerKernel({ parse: async bytes => parseXlsx(await OfficePackage.load(bytes)) }, quiet)
    const docId = id(), sheetId = sheet.sheetId ?? '1'
    await run(kernel, { kind: 'open', requestId: id(), docId, source: { kind: 'object', document: doc }, profile: 'D' })
    const session = kernel.session(docId)!
    const runtime = pivotRuntimeFor(doc)!
    const originalRecords = runtime.records.bind(runtime), admissions: number[] = []
    const stream = vi.spyOn(runtime, 'records')
    stream.mockImplementation((cache, signal, limits) => originalRecords(cache, signal, {
      ...(limits ?? RECORD_LIMITS.D), admitBytes: bytes => { admissions.push(bytes); return limits?.admitBytes?.(bytes) ?? true },
    }))
    const before = session.memoryStats().sourceBytes
    const refused = await run(kernel, { kind: 'pivotOpen', requestId: id(), docId, sheetId, pivot: 0, grantBytes: 1 })
    expect(refused.kind).toBe('error')
    expect(session.memoryStats().sourceBytes).toBe(before)
    const admitted = completed(await run(kernel, { kind: 'pivotOpen', requestId: id(), docId, sheetId, pivot: 0, grantBytes: 16 * 1024 * 1024 }))
    expect(admitted.kind).toBe('pivotOpen')
    expect(stream).toHaveBeenCalledTimes(2)
    expect(stream.mock.calls[0]?.[2]?.maxBytes).toBe(1)
    expect(stream.mock.calls[1]?.[2]?.maxBytes).toBeLessThanOrEqual(RECORD_LIMITS.D.maxBytes)
    expect(admissions.length).toBeGreaterThan(0)
    expect(Math.max(...admissions)).toBeLessThan(RECORD_LIMITS.D.maxBytes)
    expect(session.memoryStats().sourceBytes).toBeGreaterThan(before)
  })

  it('caches paged pivot assets per sheet and value generation', async () => {
    const imported = await openImportSession({ kind: 'blob', blob: new Blob([new Uint8Array(readFileSync(FIXTURE))]) }, { docId: id(), profile: 'D' })
    await imported.done
    const source = createPagedWorkbookSource(imported), sheetId = source.sheets[0]!.sheetId
    const pool = new MemoryPool({ limitBytes: 32 * 1024 * 1024 })
    source.bindMemory?.(pool)
    const store = source.store(sheetId), scan = vi.spyOn(store, 'scanOccupied')
    try {
      const first = await pivotAssets(source, sheetId)
      const second = await pivotAssets(source, sheetId)
      expect(first).toBeDefined()
      expect(second).toBe(first)
      expect(scan).toHaveBeenCalledTimes(1)
      expect(pool.totalBytes).toBeGreaterThan(0)
      store.setValueOverlay?.({ rects: [], protectedRects: [], merges: [], cell: () => undefined })
      const nextGeneration = await pivotAssets(source, sheetId)
      expect(nextGeneration).toBeDefined()
      expect(nextGeneration).not.toBe(first)
      expect(scan).toHaveBeenCalledTimes(2)
    } finally {
      await source.release()
      expect(pool.totalBytes).toBe(0)
      pool.close()
    }
  })

  it('a chart-linked paged pivot refuses at open when the saved chart mapping is unsupported', async () => {
    const file = resolve(FIXTURE_ROOT, 'interaction/chart-hide-item.xlsx')
    const imported = await openImportSession({ kind: 'blob', blob: new Blob([new Uint8Array(readFileSync(file))]) }, { docId: id(), profile: 'D' })
    await imported.done
    const source = createPagedWorkbookSource(imported), sheetId = source.sheets[0]!.sheetId
    const docId = id(), kernel = new ViewerKernel({ parse: async bytes => parseXlsx(await OfficePackage.load(bytes)), openSource: async () => source }, quiet)
    await run(kernel, { kind: 'open', requestId: id(), docId, source: { kind: 'paged', handleId: docId }, profile: 'D' })
    const assets = await pivotAssets(source, sheetId)
    expect(assets?.linkedCharts.length).toBeGreaterThan(0)
    const linked = assets!.linkedCharts[0]!
    const unsupported = { ...linked, chart: { ...linked.chart, groups: linked.chart.groups.map(group => ({ ...group, family: 'line' as const })) } }
    source.pivotAssets = async () => ({ ...assets!, linkedCharts: [unsupported, ...assets!.linkedCharts.slice(1)] })
    const pivot = assets!.tables.findIndex(def => def.name === '樞紐分析表2')
    expect(pivot).toBeGreaterThanOrEqual(0)
    const refused = await run(kernel, { kind: 'pivotOpen', requestId: id(), docId, sheetId, pivot, grantBytes: 16 * 1024 * 1024 })
    expect(JSON.stringify(refused)).toContain('pivot-chart-mapping')
    await run(kernel, { kind: 'close', requestId: id(), docId })
  })

  it('R2-4: the worker collects only pivot charts and links them by sheet and name', async () => {
    const file = resolve(FIXTURE_ROOT, 'interaction/chart-hide-item.xlsx')
    const imported = await openImportSession({ kind: 'blob', blob: new Blob([new Uint8Array(readFileSync(file))]) }, { docId: id(), profile: 'D' })
    await imported.done
    const source = createPagedWorkbookSource(imported), sheetId = source.sheets[0]!.sheetId
    const assets = (await pivotAssets(source, sheetId))!
    expect(assets.linkedCharts.length).toBeGreaterThan(0)
    expect(assets.linkedCharts.every(item => Boolean(item.chart.pivot))).toBe(true)
    // The same pivot name on another sheet: the charts are not this pivot's, so the session opens without them.
    source.pivotAssets = async () => ({ ...assets, linkedCharts: assets.linkedCharts.map(item => ({ ...item, chart: { ...item.chart, pivot: { ...item.chart.pivot!, name: '[x.xlsx]Other!樞紐分析表2' } } })) })
    const docId = id(), kernel = new ViewerKernel({ parse: async bytes => parseXlsx(await OfficePackage.load(bytes)), openSource: async () => source }, quiet)
    await run(kernel, { kind: 'open', requestId: id(), docId, source: { kind: 'paged', handleId: docId }, profile: 'D' })
    const pivot = assets.tables.findIndex(def => def.name === '樞紐分析表2')
    expect(completed(await run(kernel, { kind: 'pivotOpen', requestId: id(), docId, sheetId, pivot, grantBytes: 16 * 1024 * 1024 }))).toMatchObject({ kind: 'pivotOpen' })
    const staged = completed(await run(kernel, { kind: 'pivotApply', requestId: id(), docId, sheetId, pivot, baseRevision: 0, revision: 1, delta: { kind: 'items', field: 0, hidden: [3] }, grantBytes: 16 * 1024 * 1024 }))
    if (staged.kind !== 'pivotApply') throw new Error(JSON.stringify(staged))
    expect(staged.charts).toEqual([])
    await run(kernel, { kind: 'close', requestId: id(), docId })
  })

  it('R2-4: concurrent pivot-asset misses for one sheet and generation build once', async () => {
    const imported = await openImportSession({ kind: 'blob', blob: new Blob([new Uint8Array(readFileSync(FIXTURE))]) }, { docId: id(), profile: 'D' })
    await imported.done
    const source = createPagedWorkbookSource(imported)
    const pool = new MemoryPool({ limitBytes: 64 * 1024 * 1024 })
    source.bindMemory?.(pool)
    const sheetId = (source.sheets.find(sheet => sheet.name === 'Pivot') ?? source.sheets[0]!).sheetId
    const [a, b] = await Promise.all([pivotAssets(source, sheetId), pivotAssets(source, sheetId)])
    expect(a).toBeDefined()
    expect(b).toBe(a)
    // The one entry is still charged and still serves a later caller (no retain-under-the-same-key release).
    expect(await pivotAssets(source, sheetId)).toBe(a)
    const release = a!.retain!()
    release()
    await source.release()
  })

  it('S3-Q: in the worker an unqualified chart source naming same-named pivots on two sheets refuses pivot-chart-mapping', async () => {
    const file = resolve(FIXTURE_ROOT, 'interaction/chart-hide-item.xlsx')
    const imported = await openImportSession({ kind: 'blob', blob: new Blob([new Uint8Array(readFileSync(file))]) }, { docId: id(), profile: 'D' })
    await imported.done
    const source = createPagedWorkbookSource(imported), sheetId = source.sheets[0]!.sheetId
    const assets = (await pivotAssets(source, sheetId))!
    // The workbook's pivots, by sheet and name, travel with the assets.
    expect(assets.pivots).toEqual([{ sheet: source.sheets[0]!.name, name: '樞紐分析表2' }])
    const unqualified = assets.linkedCharts.map(item => ({ ...item, chart: { ...item.chart, pivot: { ...item.chart.pivot!, name: '樞紐分析表2' } } }))
    source.pivotAssets = async () => ({ ...assets, linkedCharts: unqualified, pivots: [...assets.pivots!, { sheet: 'Other', name: '樞紐分析表2' }] })
    const docId = id(), kernel = new ViewerKernel({ parse: async bytes => parseXlsx(await OfficePackage.load(bytes)), openSource: async () => source }, quiet)
    await run(kernel, { kind: 'open', requestId: id(), docId, source: { kind: 'paged', handleId: docId }, profile: 'D' })
    const pivot = assets.tables.findIndex(def => def.name === '樞紐分析表2')
    const refused = await run(kernel, { kind: 'pivotOpen', requestId: id(), docId, sheetId, pivot, grantBytes: 16 * 1024 * 1024 })
    expect(JSON.stringify(refused)).toContain('pivot-chart-mapping')
    await run(kernel, { kind: 'close', requestId: id(), docId })
  })

  it('a paged pivot apply publishes rebuilt linked charts into the main drawing cache', async () => {
    const file = resolve(FIXTURE_ROOT, 'interaction/chart-hide-item.xlsx')
    const blob = new Blob([new Uint8Array(readFileSync(file))])
    const imported = await openImportSession({ kind: 'blob', blob }, { docId: id(), profile: 'D' })
    await imported.done
    const source = createPagedWorkbookSource(imported), pivotSheetId = source.sheets[0]!.sheetId
    const docId = id(), kernel = new ViewerKernel({ parse: async bytes => parseXlsx(await OfficePackage.load(bytes)), openSource: async () => source }, quiet)
    const port = connectInProcess(kernel), client = new ViewerClient(port)
    const openWorkerReply = await run(kernel, { kind: 'open', requestId: id(), docId, source: { kind: 'paged', handleId: docId }, profile: 'D' })
    const openReply = completed(openWorkerReply) as any
    const remote = new RemoteWorkbook(client, { ...openReply, docId, diagnostics: openWorkerReply.diagnostics, layouts: new Map(), profile: 'D' })
    try {
      const assets = await pivotAssets(source, pivotSheetId)
      const defIndex = assets!.tables.findIndex(def => def.name === '樞紐分析表2')
      const linked = assets!.linkedCharts.filter(item => (item.chart.pivot?.name ?? '').split('!').at(-1) === assets!.tables[defIndex]!.name)
      expect(linked.length).toBeGreaterThan(0)
      const residentDoc = await parseXlsx(await OfficePackage.load(new Uint8Array(readFileSync(file))))
      const residentSheet = residentDoc.sheets.find(candidate => candidate.pivotTables?.some(def => def.name === assets!.tables[defIndex]!.name))!
      const residentDef = residentSheet.pivotTables!.find(def => def.name === assets!.tables[defIndex]!.name)!
      const residentRuntime = pivotRuntimeFor(residentDoc)!
      const residentCache = residentRuntime.cache(residentDef)!
      const residentRecords = await residentRuntime.records(residentCache, new AbortController().signal, RECORD_LIMITS.D)
      if (!residentRecords.ok) throw new Error(`Resident chart case records refused: ${residentRecords.diagnostic.feature}`)
      const residentOpen = openPivotSession({ def: residentDef, cache: residentCache, store: residentRecords.store,
        date1904: residentRuntime.date1904, saved: sheetCellReader(residentSheet), occupied: (row, col) => assets!.occupied(row, col),
        linkedCharts: linked.map(item => item.chart), chartGate: pivotChartGate })
      if (!residentOpen.ok) throw new Error(`Resident chart case refused: ${residentOpen.diagnostic.feature}`)
      const delta = { kind: 'items' as const, field: 0, hidden: [3] }
      const residentApply = residentOpen.value.apply(delta)
      if (!residentApply.ok) throw new Error(`Resident chart apply refused: ${residentApply.diagnostic.feature}`)
      for (const sheetId of new Set(linked.map(item => item.sheetId))) {
        await vi.waitFor(() => expect(remote.drawings(sheetId).state).toBe('ready'), { timeout: 15_000 })
      }
      const beforeGen = new Map([...new Set(linked.map(item => item.sheetId))].map(sheetId => [sheetId, remote.drawings(sheetId).gen]))
      const beforeLists = new Map([...beforeGen.keys()].map(sheetId => [sheetId, remote.drawings(sheetId).drawings]))
      const mergeGen = remote.merges(pivotSheetId).gen
      const findIn = (nodes: readonly any[], item: { partPath: string; treePath: string }): any => nodes.flatMap(node => [node, ...(node.children ?? [])]).find(node => node.source?.partPath === item.partPath && node.source?.treePath === item.treePath)
      const savedCharts = linked.map(item => findIn(remote.drawings(item.sheetId).drawings, item)?.content?.chart)
      expect(await remote.pivots.open(Number(pivotSheetId), defIndex, new AbortController().signal)).toMatchObject({ kind: 'open' })
      const committed = await remote.pivots.apply(Number(pivotSheetId), defIndex, delta)
      expect(committed.kind).toBe('committed')
      if (committed.kind !== 'committed') throw new Error(JSON.stringify(committed))
      expect(committed.charts).toHaveLength(linked.length)
      expect(committed.charts.map(item => item.chart)).toEqual(residentApply.value.charts)
      for (const item of committed.charts) {
        const drawings = remote.drawings(item.sheetId)
        expect(drawings.gen).toBeGreaterThan(beforeGen.get(item.sheetId) ?? -1)
        const find = (nodes: readonly any[]): any => nodes.flatMap(node => [node, ...(node.children ?? [])]).find(node => node.source?.partPath === item.partPath && node.source?.treePath === item.treePath)
        expect(find(drawings.drawings)?.content?.chart).toEqual(item.chart)
        // R2-4: a new drawing list (copied along the changed path), never the old nodes mutated in place.
        expect(drawings.drawings).not.toBe(beforeLists.get(item.sheetId))
        expect(findIn(beforeLists.get(item.sheetId)!, item)?.content?.chart).toBe(savedCharts[committed.charts.indexOf(item)])
      }
      expect(remote.merges(pivotSheetId).gen).toBeGreaterThan(mergeGen)
      // R2-4: the replacement models are charged to the main ledger's drawings pool while they live.
      for (const item of committed.charts) {
        const installed = findIn(remote.drawings(item.sheetId).drawings, item)?.content?.chart
        expect(remote.mainLedger.hasLifetime(installed)).toBe(true)
      }
      // Release restores the saved models (the same objects) in yet another list, with a newer generation.
      const committedGen = new Map([...beforeGen.keys()].map(sheetId => [sheetId, remote.drawings(sheetId).gen]))
      remote.pivots.release(Number(pivotSheetId), defIndex)
      await port.idle()
      linked.forEach((item, k) => {
        expect(findIn(remote.drawings(item.sheetId).drawings, item)?.content?.chart).toBe(savedCharts[k])
        expect(remote.drawings(item.sheetId).gen).toBeGreaterThan(committedGen.get(item.sheetId)!)
      })
    } finally {
      await remote.close()
      await port.idle()
    }
  }, 60_000)

  it('Task 28: open and apply replies carry the menus\' model; a paged pivotDisplay read serves the committed definition', async () => {
    const blob = new Blob([new Uint8Array(readFileSync(FIXTURE))])
    const docId = id()
    const imported = await openImportSession({ kind: 'blob', blob }, { docId, profile: 'D' })
    await imported.done
    const source = createPagedWorkbookSource(imported)
    const kernel = new ViewerKernel({ parse: async bytes => parseXlsx(await OfficePackage.load(bytes)), openSource: async () => source }, quiet)
    await run(kernel, { kind: 'open', requestId: id(), docId, source: { kind: 'paged', handleId: docId }, profile: 'D' })
    const sheetId = source.sheets.find(sheet => sheet.name === 'Pivot')!.sheetId
    const pivot = (await pivotAssets(source, sheetId))!.tables.findIndex(def => def.name === '樞紐分析表2')
    const opened = completed(await run(kernel, { kind: 'pivotOpen', requestId: id(), docId, sheetId, pivot, grantBytes: 16 * 1024 * 1024 })) as Extract<WorkerReply<'pivotOpen'>, { kind: 'complete' }>['result']
    if (opened.kind !== 'pivotOpen') throw new Error(JSON.stringify(opened))
    const region = opened.menu.fields.find(field => field.name === 'Region')!
    expect(region.items.map(item => item.label)).toEqual(['East', 'North', 'South', 'West', 'Central'])
    expect(region.items[0]!.collapsed).toBe(false)
    expect(opened.menu.data.map(entry => entry.name)).toEqual(['Sum of Sales', 'Sum of Units'])
    const staged = completed(await run(kernel, { kind: 'pivotApply', requestId: id(), docId, sheetId, pivot, baseRevision: 0, revision: 1, delta: { kind: 'collapse', field: region.index, item: 0, collapsed: true }, grantBytes: 16 * 1024 * 1024 }))
    if (staged.kind !== 'pivotApply') throw new Error(JSON.stringify(staged))
    expect(staged.menu.fields.find(field => field.name === 'Region')!.items[0]!.collapsed).toBe(true)
    const commit = completed(await run(kernel, { kind: 'pivotCommit', requestId: id(), docId, sheetId, pivot, revision: staged.revision }))
    if (commit.kind !== 'pivotCommit' || 'refused' in commit) throw new Error(JSON.stringify(commit))
    const read = completed(await run(kernel, { kind: 'read', requestId: id(), owner: { docId, sheetId, valueGen: commit.valueGen, fontGen: 0, viewGen: 0 }, query: { op: 'pivotDisplay', grantBytes: 4 * 1024 * 1024 } }))
    if (read.kind !== 'read' || read.result.op !== 'pivotDisplay') throw new Error(JSON.stringify(read))
    expect(read.result.result.pivotTables.find(def => def.name === '樞紐分析表2')!.location.ref).toBe('A3:I21')
    await run(kernel, { kind: 'close', requestId: id(), docId })
  })

  it('R2-1 (S3-L): the first commit protects the saved area, end to end, and main publishes no stale saved merge', async () => {
    const file = resolve(FIXTURE_ROOT, 'engine/layout-merge-labels.xlsx')
    const blob = new Blob([new Uint8Array(readFileSync(file))])
    const docId = id()
    const imported = await openImportSession({ kind: 'blob', blob }, { docId, profile: 'D' })
    await imported.done
    const source = createPagedWorkbookSource(imported)
    const kernel = new ViewerKernel({ parse: async bytes => parseXlsx(await OfficePackage.load(bytes)), openSource: async () => source }, quiet)
    const port = connectInProcess(kernel), client = new ViewerClient(port)
    const opened = await openRemoteWorkbook(client, docId, blob, 'D')
    if ('kind' in opened) throw new Error(`Open refused: ${JSON.stringify(opened)}`)
    const remote = new RemoteWorkbook(client, opened)
    try {
      const sheet = source.sheets.find(candidate => candidate.name === 'Pivot') ?? source.sheets[0]!
      const assets = (await pivotAssets(source, sheet.sheetId))!
      const pivot = assets.tables.findIndex(def => def.name === '樞紐分析表2')
      const saved = assets.tables[pivot]!.location.ref
      expect(saved).toBe('A3:J24')
      remote.merges(sheet.sheetId)
      await vi.waitFor(() => expect(remote.merges(sheet.sheetId).merges.length).toBeGreaterThan(0))
      expect((await remote.pivots.open(Number(sheet.sheetId), pivot, new AbortController().signal)).kind).toBe('open')
      const result = await remote.pivots.apply(Number(sheet.sheetId), pivot, { kind: 'items', field: 0, hidden: [0] })
      if (result.kind !== 'committed') throw new Error(JSON.stringify(result))
      // Every saved row (2–23) and column (0–9) is inside the protected set the worker produced.
      for (let row = 2; row <= 23; row++) expect(result.protectedRects.some(rect => rect.r0 <= row && row <= rect.r1 && rect.c0 <= 0 && 9 <= rect.c1), `row ${row}`).toBe(true)
      // Main and worker agree: no saved label merge survives over the vacated rows.
      const fresh = new Set(result.merges.map(rect => JSON.stringify([rect.r0, rect.r1, rect.c0, rect.c1])))
      const stale = remote.merges(sheet.sheetId).merges.filter(merge => !fresh.has(JSON.stringify([merge.minRow, merge.maxRow, merge.minCol, merge.maxCol])) && merge.minRow >= 2 && merge.maxRow <= 23 && merge.maxCol <= 9)
      expect(stale).toEqual([])
    } finally {
      await remote.close()
      await port.idle()
    }
  }, 60_000)

  it('R2-2: a paged pool returns to its pre-open baseline after the last release, a repeated open included', async () => {
    const file = resolve(FIXTURE_ROOT, 'engine/layout-merge-labels.xlsx')
    const blob = new Blob([new Uint8Array(readFileSync(file))])
    const docId = id()
    const imported = await openImportSession({ kind: 'blob', blob }, { docId, profile: 'D' })
    await imported.done
    const source = createPagedWorkbookSource(imported)
    const kernel = new ViewerKernel({ parse: async bytes => parseXlsx(await OfficePackage.load(bytes)), openSource: async () => source }, quiet)
    await run(kernel, { kind: 'open', requestId: id(), docId, source: { kind: 'paged', handleId: docId }, profile: 'D' })
    const session = kernel.session(docId)!
    const sheetId = (source.sheets.find(candidate => candidate.name === 'Pivot') ?? source.sheets[0]!).sheetId
    const before = session.memoryStats().sourceBytes
    const pivotSlots = () => [...((session as unknown as { memory: { slots: Map<string, unknown> } }).memory.slots.keys())].filter(key => key.startsWith('pivot-'))
    const settled: number[] = []
    for (let k = 0; k < 2; k++) {
      expect(completed(await run(kernel, { kind: 'pivotOpen', requestId: id(), docId, sheetId, pivot: 0, grantBytes: 16 * 1024 * 1024 }))).toMatchObject({ kind: 'pivotOpen' })
      // A second open of the open pivot returns the existing handle (it reads the assets again; nothing may stay held).
      expect(completed(await run(kernel, { kind: 'pivotOpen', requestId: id(), docId, sheetId, pivot: 0, grantBytes: 16 * 1024 * 1024 }))).toMatchObject({ kind: 'pivotOpen' })
      expect(pivotSlots().length).toBeGreaterThan(0)
      expect(completed(await run(kernel, { kind: 'pivotRelease', requestId: id(), docId, sheetId, pivot: 0 }))).toMatchObject({ kind: 'pivotRelease', released: true })
      // Nothing of the pivot stays charged (assets, runtime, store, overlay); only evictable source pages the reads
      // brought in remain (page/sst cache), the same after every round.
      expect(pivotSlots()).toEqual([])
      settled.push(session.memoryStats().sourceBytes)
    }
    expect(settled[1]).toBe(settled[0])
    expect(settled[0]! - before).toBeLessThan(8192)
    await run(kernel, { kind: 'close', requestId: id(), docId })
  })

  it('R2-3: an aborted open that created a handle releases only that handle, never a newer reopen', async () => {
    const doc = await parseXlsx(await OfficePackage.load(new Uint8Array(readFileSync(FIXTURE))))
    const sheetId = doc.sheets.find(candidate => candidate.pivotTables?.length)!.sheetId ?? '1'
    const docId = id()
    let kernel!: ViewerKernel, first = true
    kernel = new ViewerKernel({ parse: async bytes => parseXlsx(await OfficePackage.load(bytes)), afterDispatch: async command => {
      if (command.kind !== 'pivotOpen' || !first) return
      first = false
      // The client cancels open1 (it created H1), releases the pivot and opens it again (H2) before open1's cleanup runs.
      await kernel.handle({ kind: 'cancel', requestId: id(), targetRequestId: command.requestId } as WorkerCommand, () => {})
      const session = kernel.session(docId)!
      session.pivotRelease(sheetId, 0)
      expect(await session.pivotOpen(sheetId, 0, 16 * 1024 * 1024, new AbortController().signal)).toMatchObject({ kind: 'open', created: true })
    } }, quiet)
    await run(kernel, { kind: 'open', requestId: id(), docId, source: { kind: 'object', document: doc }, profile: 'D' })
    await kernel.handle({ kind: 'pivotOpen', requestId: id(), docId, sheetId, pivot: 0, grantBytes: 16 * 1024 * 1024 }, () => {})
    // H2 survives open1's cleanup: an apply on it is staged, not refused as released.
    const staged = kernel.session(docId)!.pivotApply(sheetId, 0, 0, 1, { kind: 'collapse', field: 0, item: 0, collapsed: true }, 16 * 1024 * 1024)
    expect(staged.kind).toBe('apply')
    await run(kernel, { kind: 'close', requestId: id(), docId })
  })

  it('apply commits only the current revision', async () => {
    const doc = await parseXlsx(await OfficePackage.load(new Uint8Array(readFileSync(FIXTURE))))
    const sheet = doc.sheets.find(candidate => candidate.pivotTables?.length)!
    const def = sheet.pivotTables![0]!
    const kernel = new ViewerKernel({ parse: async bytes => parseXlsx(await OfficePackage.load(bytes)) }, quiet)
    const docId = id(), sheetId = sheet.sheetId ?? '1'
    await run(kernel, { kind: 'open', requestId: id(), docId, source: { kind: 'object', document: doc }, profile: 'D' })

    const open = await run(kernel, {
      kind: 'pivotOpen', requestId: id(), docId, sheetId, pivot: 0, grantBytes: 16 * 1024 * 1024,
    })
    const opened = completed(open) as unknown as { kind: string; revision: number; def: { name: string } }
    expect(opened.kind).toBe('pivotOpen')
    expect(opened.revision).toBe(0)
    expect(opened.def.name).toBe(def.name)

    const store = kernel.session(docId)!.source.store(sheetId)
    const before = store.stamp.valueGen
    const point = { row: 21, col: 4 }
    const base = await store.readCell(point, { signal: new AbortController().signal, stamp: store.stamp })
    const staged = await run(kernel, {
      kind: 'pivotApply', requestId: id(), docId, sheetId, pivot: 0, baseRevision: 0, revision: 1,
      delta: { kind: 'collapse', field: 0, item: 0, collapsed: true }, grantBytes: 16 * 1024 * 1024,
    })
    const apply = completed(staged) as unknown as { kind: string; revision: number }
    expect(apply.kind).toBe('pivotApply')
    expect(apply.revision).toBe(1)
    expect(store.stamp.valueGen).toBe(before)
    expect(await store.readCell(point, { signal: new AbortController().signal, stamp: store.stamp })).toEqual(base)

    const stale = await run(kernel, {
      kind: 'pivotCommit', requestId: id(), docId, sheetId, pivot: 0, revision: 2,
    })
    expect(completed(stale)).toMatchObject({ kind: 'pivotCommit', refused: 'stale' })
    expect(store.stamp.valueGen).toBe(before)

    const committed = await run(kernel, {
      kind: 'pivotCommit', requestId: id(), docId, sheetId, pivot: 0, revision: 1,
    })
    expect(completed(committed)).toMatchObject({ kind: 'pivotCommit', revision: 1, valueGen: before + 1 })
    expect(store.stamp.valueGen).toBe(before + 1)
    const after = await store.readCell(point, { signal: new AbortController().signal, stamp: store.stamp })
    expect(after).not.toEqual(base)
  })

  it('stale apply reply is dropped and refunded', async () => {
    const doc = await parseXlsx(await OfficePackage.load(new Uint8Array(readFileSync(FIXTURE))))
    const sheet = doc.sheets.find(candidate => candidate.pivotTables?.length)!
    const gate: { release?: () => void; entered?: () => void } = {}
    const waiting = new Promise<void>(resolve => { gate.entered = resolve })
    const kernel = new ViewerKernel({
      parse: async bytes => parseXlsx(await OfficePackage.load(bytes)),
      afterDispatch: command => command.kind === 'pivotApply' && command.revision === 1
        ? new Promise<void>(resolve => { gate.release = resolve; gate.entered?.() }) : undefined,
    }, quiet)
    const docId = id(), sheetId = sheet.sheetId ?? '1'
    await run(kernel, { kind: 'open', requestId: id(), docId, source: { kind: 'object', document: doc }, profile: 'D' })
    await run(kernel, { kind: 'pivotOpen', requestId: id(), docId, sheetId, pivot: 0, grantBytes: 16 * 1024 * 1024 })
    const before = kernel.session(docId)!.memoryStats().sourceBytes
    const first = run(kernel, { kind: 'pivotApply', requestId: id(), docId, sheetId, pivot: 0, baseRevision: 0, revision: 1,
      delta: { kind: 'collapse', field: 0, item: 0, collapsed: true }, grantBytes: 16 * 1024 * 1024 })
    await waiting
    const afterFirstStage = kernel.session(docId)!.memoryStats().sourceBytes
    expect(afterFirstStage).toBeGreaterThan(before)
    const second = await run(kernel, { kind: 'pivotApply', requestId: id(), docId, sheetId, pivot: 0, baseRevision: 0, revision: 2,
      delta: { kind: 'collapse', field: 0, item: 1, collapsed: true }, grantBytes: 16 * 1024 * 1024 })
    expect(completed(second)).toMatchObject({ kind: 'pivotApply', revision: 2 })
    const afterSecondStage = kernel.session(docId)!.memoryStats().sourceBytes
    const store = kernel.session(docId)!.source.store(sheetId)
    const beforeSupersedingCommit = await store.readCell({ row: 21, col: 4 }, { signal: new AbortController().signal, stamp: store.stamp })
    expect(store.stamp.valueGen).toBe(0)
    expect(beforeSupersedingCommit).toEqual(await store.readCell({ row: 21, col: 4 }, { signal: new AbortController().signal, stamp: store.stamp }))
    gate.release?.()
    expect((await first).kind).toBe('cancelled')
    const afterReplace = kernel.session(docId)!.memoryStats().sourceBytes
    expect(afterReplace).toBe(afterSecondStage)
    expect(completed(await run(kernel, { kind: 'pivotCommit', requestId: id(), docId, sheetId, pivot: 0, revision: 1 })))
      .toMatchObject({ kind: 'pivotCommit', refused: 'stale' })
    await run(kernel, { kind: 'pivotCommit', requestId: id(), docId, sheetId, pivot: 0, revision: 2 })
    await run(kernel, { kind: 'pivotRelease', requestId: id(), docId, sheetId, pivot: 0 })
    expect(kernel.session(docId)!.memoryStats().sourceBytes).toBeLessThan(before)
  })

  it('a staged candidate is never readable', async () => {
    const { kernel, docId, sheetId, session } = await openObjectPivot()
    const store = session.source.store(sheetId), initialGen = store.stamp.valueGen
    const point = { row: 21, col: 4 }
    const base = await store.readCell(point, { signal: new AbortController().signal, stamp: store.stamp })
    const a = completed(await run(kernel, { kind: 'pivotApply', requestId: id(), docId, sheetId, pivot: 0, baseRevision: 0, revision: 1,
      delta: { kind: 'collapse', field: 0, item: 0, collapsed: true }, grantBytes: 16 * 1024 * 1024 })) as { kind: string }
    expect(a.kind).toBe('pivotApply')
    const b = completed(await run(kernel, { kind: 'pivotApply', requestId: id(), docId, sheetId, pivot: 0, baseRevision: 0, revision: 2,
      delta: { kind: 'collapse', field: 0, item: 1, collapsed: true }, grantBytes: 16 * 1024 * 1024 })) as { kind: string }
    expect(b.kind).toBe('pivotApply')
    expect(store.stamp.valueGen).toBe(initialGen)
    expect(await store.readCell(point, { signal: new AbortController().signal, stamp: store.stamp })).toEqual(base)
    await run(kernel, { kind: 'pivotCommit', requestId: id(), docId, sheetId, pivot: 0, revision: 2 })
    expect(store.stamp.valueGen).toBe(initialGen + 1)
    expect(await store.readCell(point, { signal: new AbortController().signal, stamp: store.stamp })).not.toEqual(base)
  })

  it('pivotCommit with a revision that is not staged refuses as stale', async () => {
    const { kernel, docId, sheetId, session } = await openObjectPivot()
    const before = session.source.store(sheetId).stamp.valueGen
    const reply = await run(kernel, { kind: 'pivotCommit', requestId: id(), docId, sheetId, pivot: 0, revision: 99 })
    expect(completed(reply)).toMatchObject({ kind: 'pivotCommit', refused: 'stale' })
    expect(session.source.store(sheetId).stamp.valueGen).toBe(before)
  })

  it('release on every exit: close, sheet switch, abort, and owner change', async () => {
    let resume: (() => void) | undefined, entered!: () => void
    const applyEntered = new Promise<void>(resolve => { entered = resolve })
    const { doc, kernel, docId, sheetId, session } = await openObjectPivot({ afterDispatch: command => {
      if (command.kind === 'pivotApply' && command.revision === 1) return new Promise<void>(resolve => { resume = resolve; entered() })
    } })
    const otherSheet = doc.sheets.find(sheet => sheet.sheetId !== sheetId)!
    await session.pivotRelease(sheetId, 0)
    const baseline = session.memoryStats().sourceBytes
    await run(kernel, { kind: 'pivotOpen', requestId: id(), docId, sheetId, pivot: 0, grantBytes: 16 * 1024 * 1024 })
    const sourceBytes = session.memoryStats().sourceBytes
    expect(sourceBytes).toBeGreaterThan(baseline)
    await run(kernel, { kind: 'activateSheet', requestId: id(), owner: session.workbookOwner(), sheetId: otherSheet.sheetId ?? '2' })
    expect(session.memoryStats().sourceBytes).toBe(baseline)

    await run(kernel, { kind: 'pivotOpen', requestId: id(), docId, sheetId, pivot: 0, grantBytes: 16 * 1024 * 1024 })
    const openBytes = session.memoryStats().sourceBytes
    const requestId = id()
    const apply = run(kernel, { kind: 'pivotApply', requestId, docId, sheetId, pivot: 0, baseRevision: 0, revision: 1,
      delta: { kind: 'collapse', field: 0, item: 0, collapsed: true }, grantBytes: 16 * 1024 * 1024 })
    await applyEntered
    expect(session.memoryStats().sourceBytes).toBeGreaterThan(openBytes)
    await run(kernel, { kind: 'cancel', requestId: id(), targetRequestId: requestId })
    resume?.()
    expect((await apply).kind).toBe('cancelled')
    expect(session.memoryStats().sourceBytes).toBe(openBytes)
    await run(kernel, { kind: 'close', requestId: id(), docId })
    expect(kernel.lifecycleStats().retainedBytes).toBe(0)
  })

  it('M5: an owner change (the document closed and opened again under its id) ends the old sessions and their charges', async () => {
    const doc = await parseXlsx(await OfficePackage.load(new Uint8Array(readFileSync(FIXTURE))))
    const sheetId = doc.sheets.find(candidate => candidate.pivotTables?.length)!.sheetId ?? '1'
    const kernel = new ViewerKernel({ parse: async bytes => parseXlsx(await OfficePackage.load(bytes)) }, quiet)
    const docId = id()
    await run(kernel, { kind: 'open', requestId: id(), docId, source: { kind: 'object', document: doc }, profile: 'D' })
    const first = kernel.session(docId)!
    expect(completed(await run(kernel, { kind: 'pivotOpen', requestId: id(), docId, sheetId, pivot: 0, grantBytes: 16 * 1024 * 1024 }))).toMatchObject({ kind: 'pivotOpen', revision: 0 })
    const staged = completed(await run(kernel, { kind: 'pivotApply', requestId: id(), docId, sheetId, pivot: 0, baseRevision: 0, revision: 1, delta: { kind: 'collapse', field: 0, item: 0, collapsed: true }, grantBytes: 16 * 1024 * 1024 }))
    expect(staged).toMatchObject({ kind: 'pivotApply', revision: 1 })
    await run(kernel, { kind: 'close', requestId: id(), docId })
    expect(kernel.lifecycleStats().retainedBytes).toBe(0)
    // The old owner's session holds nothing: its staged candidate and handle went with the close.
    expect(first.pivotApply(sheetId, 0, 0, 2, { kind: 'collapse', field: 0, item: 0, collapsed: true }, 16 * 1024 * 1024)).toMatchObject({ kind: 'refused', feature: 'pivot-session-released' })
    // The same id, a new owner: the old revision is not committable and nothing of the old session is open here.
    await run(kernel, { kind: 'open', requestId: id(), docId, source: { kind: 'object', document: doc }, profile: 'D' })
    const second = kernel.session(docId)!
    expect(second).not.toBe(first)
    const stale = completed(await run(kernel, { kind: 'pivotCommit', requestId: id(), docId, sheetId, pivot: 0, revision: 1 }))
    expect(stale).toMatchObject({ kind: 'pivotCommit', refused: 'stale' })
    expect(second.pivotApply(sheetId, 0, 0, 1, { kind: 'collapse', field: 0, item: 0, collapsed: true }, 16 * 1024 * 1024)).toMatchObject({ kind: 'refused', feature: 'pivot-session-released' })
    expect(completed(await run(kernel, { kind: 'pivotOpen', requestId: id(), docId, sheetId, pivot: 0, grantBytes: 16 * 1024 * 1024 }))).toMatchObject({ kind: 'pivotOpen', revision: 0 })
    await run(kernel, { kind: 'close', requestId: id(), docId })
    expect(kernel.lifecycleStats().retainedBytes).toBe(0)
  })

  it('worker pool charges follow the overlay lifecycle', async () => {
    const { kernel, docId, sheetId, session, beforePivotOpen } = await openObjectPivot()
    const beforeOpen = beforePivotOpen
    const opened = session.memoryStats().sourceBytes
    expect(opened).toBeGreaterThan(beforeOpen)
    const sendApply = (revision: number, item: number, baseRevision: number) => run(kernel, { kind: 'pivotApply', requestId: id(), docId, sheetId, pivot: 0, baseRevision, revision,
      delta: { kind: 'collapse', field: 0, item, collapsed: true }, grantBytes: 16 * 1024 * 1024 })
    await sendApply(1, 0, 0)
    const stageA = session.memoryStats().sourceBytes
    expect(stageA).toBeGreaterThan(opened)
    const stagedB = completed(await sendApply(2, 1, 0))
    expect(stagedB.kind).toBe('pivotApply')
    const stageB = session.memoryStats().sourceBytes
    expect(stageB).toBe(stageA)
    const stale = completed(await run(kernel, { kind: 'pivotCommit', requestId: id(), docId, sheetId, pivot: 0, revision: 1 }))
    expect(stale).toMatchObject({ kind: 'pivotCommit', refused: 'stale' })
    expect(session.memoryStats().sourceBytes).toBe(stageB)
    await run(kernel, { kind: 'pivotCommit', requestId: id(), docId, sheetId, pivot: 0, revision: 2 })
    const committedB = session.memoryStats().sourceBytes
    expect(committedB).toBe(stageB)
    const stagedC = completed(await sendApply(3, 2, 2))
    expect(stagedC.kind).toBe('pivotApply')
    if (stagedB.kind === 'pivotApply' && stagedC.kind === 'pivotApply') {
      expect(stagedC.protectedRects).toEqual(expect.arrayContaining([...stagedB.protectedRects]))
      expect(stagedC.overlayExtent).toEqual(overlayUsedExtent(stagedC.protectedRects))
    }
    const stageC = session.memoryStats().sourceBytes
    expect(stageC).toBeGreaterThan(committedB)
    await run(kernel, { kind: 'pivotCommit', requestId: id(), docId, sheetId, pivot: 0, revision: 3 })
    const committedC = session.memoryStats().sourceBytes
    expect(committedC).toBe(stageC - (committedB - opened))
    await run(kernel, { kind: 'pivotRelease', requestId: id(), docId, sheetId, pivot: 0 })
    expect(session.memoryStats().sourceBytes).toBe(beforeOpen)
  })

  it('pivots on one cache share one store until the last release', async () => {
    const { kernel, docId, sheetId, session, beforePivotOpen } = await openObjectPivot({ duplicate: true })
    const baseline = beforePivotOpen
    const afterFirst = session.memoryStats().sourceBytes
    expect(afterFirst).toBeGreaterThan(baseline)
    const second = completed(await run(kernel, { kind: 'pivotOpen', requestId: id(), docId, sheetId, pivot: 1, grantBytes: 16 * 1024 * 1024 }))
    expect(second.kind).toBe('pivotOpen')
    expect(session.memoryStats().sourceBytes).toBe(afterFirst)
    await run(kernel, { kind: 'pivotRelease', requestId: id(), docId, sheetId, pivot: 0 })
    expect(session.memoryStats().sourceBytes).toBe(afterFirst)
    await run(kernel, { kind: 'pivotRelease', requestId: id(), docId, sheetId, pivot: 1 })
    expect(session.memoryStats().sourceBytes).toBe(baseline)
  })

  it.each(TASK25_CASES.map(entry => [entry.case, entry] as const))('paged layout equals resident layout for interaction case %s', async (_name, entry) => {
    const start = entry.base ? resolve(FIXTURE_ROOT, `${entry.base}.xlsx`) : SEEDS[entry.seed]
    if (!start) throw new Error(`Unknown interaction seed ${entry.seed}`)
    const parts = await loadPivotParts(start)
    const residentParts = parts.find(part => part.def.name === entry.pivot)
    if (!residentParts) throw new Error(`${entry.case}: missing pivot ${entry.pivot} in ${start}`)
    const resident = await decodeParts(residentParts)
    const saved = sheetCellReader(resident.sheet)
    const residentOpen = openPivotSession({ def: resident.def, cache: resident.cache, store: resident.store,
      date1904: resident.date1904, saved, occupied: (row, col) => saved(row, col) !== undefined,
      linkedCharts: [], verifiedKinds: ORACLE_VERIFIED_KINDS })
    if (!residentOpen.ok) throw new Error(`${entry.case}: resident open refused ${residentOpen.diagnostic.feature}`)

    const docId = id()
    const imported = await openImportSession({ kind: 'blob', blob: new Blob([new Uint8Array(readFileSync(start))]) }, { docId, profile: 'D' })
    await imported.done
    const source = createPagedWorkbookSource(imported)
    const sheet = source.sheets.find(candidate => candidate.name === resident.def.sheet)
    if (!sheet) throw new Error(`${entry.case}: paged sheet ${resident.def.sheet} not found`)
    const assets = await pivotAssets(source, sheet.sheetId)
    if (!assets) throw new Error(`${entry.case}: paged pivot assets unavailable`)
    const pivot = assets.tables.findIndex(def => def.name === entry.pivot)
    if (pivot < 0) throw new Error(`${entry.case}: paged pivot ${entry.pivot} not found`)
    const kernel = new ViewerKernel({ parse: async bytes => parseXlsx(await OfficePackage.load(bytes)), openSource: async () => source }, quiet)
    const openReply = await run(kernel, { kind: 'open', requestId: id(), docId, source: { kind: 'paged', handleId: docId }, profile: 'D' })
    expect(completed(openReply)).toMatchObject({ kind: 'open' })
    const session = kernel.session(docId)!
    const pivotOpen = await run(kernel, { kind: 'pivotOpen', requestId: id(), docId, sheetId: sheet.sheetId, pivot, grantBytes: 16 * 1024 * 1024 })
    expect(completed(pivotOpen)).toMatchObject({ kind: 'pivotOpen' })
    let revision = 0
    for (const delta of entry.delta) {
      revision++
      const residentApply = residentOpen.value.apply(delta)
      if (!residentApply.ok) throw new Error(`${entry.case}: resident apply refused ${residentApply.diagnostic.feature}`)
      const pagedReply = completed(await run(kernel, { kind: 'pivotApply', requestId: id(), docId, sheetId: sheet.sheetId, pivot, baseRevision: revision - 1, revision,
        delta, grantBytes: 16 * 1024 * 1024 })) as { kind: string; cells?: Array<{ row: number; col: number; value: unknown }> }
      expect(pagedReply.kind, entry.case).toBe('pivotApply')
      // The reply carries the value cells the main thread's autofit reads (Task 29 r1: bounded), as the resident layout has them.
      const placed = new Map(residentApply.value.layout.cells.map(cell => [`${cell.row},${cell.col}`, cell]))
      const expected = pivotFitPayload(residentApply.value.def, (row, col) => {
        const cell = placed.get(`${row},${col}`)
        return !cell || cell.value.kind === 'empty' ? null : {
          value: cell.value.value, ...(cell.value.kind === 'error' ? { valueIsError: true } : {}),
          ...(cell.numFmtId === undefined ? {} : { numFmtId: cell.numFmtId }),
        }
      })
      expect(expected?.length).toBeGreaterThan(0)
      expect(pagedReply.cells).toEqual(expected)
      const committed = await run(kernel, { kind: 'pivotCommit', requestId: id(), docId, sheetId: sheet.sheetId, pivot, revision })
      expect(completed(committed)).toMatchObject({ kind: 'pivotCommit', revision })
      expect(residentOpen.value.commit(residentApply.value)).toBe(true)
    }
    expect(session.source.store(sheet.sheetId).stamp.valueGen).toBe(revision)
    await run(kernel, { kind: 'close', requestId: id(), docId })
    residentOpen.value.release()
  }, 60_000)

  it('a paged apply that changes a merge and grows the extent updates main-thread geometry', async () => {
    const file = resolve(FIXTURE_ROOT, 'engine/layout-merge-labels.xlsx')
    const blob = new Blob([new Uint8Array(readFileSync(file))])
    const docId = id()
    const imported = await openImportSession({ kind: 'blob', blob }, { docId, profile: 'D' })
    await imported.done
    const source = createPagedWorkbookSource(imported)
    const kernel = new ViewerKernel({ parse: async bytes => parseXlsx(await OfficePackage.load(bytes)), openSource: async () => source }, quiet)
    const port = connectInProcess(kernel), client = new ViewerClient(port)
    const opened = await openRemoteWorkbook(client, docId, blob, 'D')
    if ('kind' in opened) throw new Error(`Open refused: ${JSON.stringify(opened)}`)
    const sheet = source.sheets.find(candidate => candidate.name === 'Pivot') ?? source.sheets[0]!
    const layouts = new Map(opened.layouts)
    layouts.set(sheet.sheetId, { ...opened.layouts.get(sheet.sheetId)!, used: { r0: 0, r1: 0, c0: 0, c1: 0 } })
    const remote = new RemoteWorkbook(client, { ...opened, layouts })
    try {
      const sheetId = Number(sheet.sheetId), before = remote.layout(sheet.sheetId).used, oldMerges = [...remote.merges(sheet.sheetId).merges], mergeGen = remote.merges(sheet.sheetId).gen
      const pivotIndex = (await pivotAssets(source, sheet.sheetId))!.tables.findIndex(def => def.name === '樞紐分析表2')
      expect(await remote.pivots.open(sheetId, pivotIndex, new AbortController().signal)).toMatchObject({ kind: 'open' })
      const result = await remote.pivots.apply(sheetId, pivotIndex, { kind: 'items', field: 0, hidden: [0] })
      expect(result.kind).toBe('committed')
      if (result.kind !== 'committed') throw new Error(JSON.stringify(result))
      expect(remote.merges(sheet.sheetId).gen).toBeGreaterThan(mergeGen)
      expect(result.merges.map(rect => [rect.r0, rect.r1, rect.c0, rect.c1])).not.toEqual(oldMerges.map(rect => [rect.minRow, rect.maxRow, rect.minCol, rect.maxCol]))
      expect(remote.merges(sheet.sheetId).merges).toEqual(expect.arrayContaining(result.merges.map(rect => ({ minRow: rect.r0, maxRow: rect.r1, minCol: rect.c0, maxCol: rect.c1 }))))
      expect(before).toBeDefined()
      // The published extent is the overlay's protected bounds, as the stores' extent reads it (Task 26 N6).
      expect({ usedRows: result.usedRows, usedCols: result.usedCols }).toEqual(overlayUsedCounts(result.protectedRects))
      expect(remote.layout(sheet.sheetId).used?.r1).toBeGreaterThan(before!.r1)
      expect(remote.layout(sheet.sheetId).used?.c1).toBeGreaterThanOrEqual(result.usedCols - 1)
      const geometry = pagedSheetGeometry(remote, sheet.sheetId, createIdentityOrder(), visibilityFromHidden(createIntervalSet([])))
      expect(geometry.metrics.sourceCellToLogical({ row: result.usedRows - 1, col: result.usedCols - 1 })).not.toBeNull()
    } finally {
      await remote.close()
      await port.idle()
    }
  }, 60_000)

  it('publishes true pivot bounds, suppresses covered source merges, and restores/shrinks on release', async () => {
    const blob = new Blob([new Uint8Array(readFileSync(FIXTURE))])
    const docId = id()
    const kernel = new ViewerKernel({ parse: async bytes => parseXlsx(await OfficePackage.load(bytes)) }, quiet)
    const port = connectInProcess(kernel), client = new ViewerClient(port)
    const opened = await openRemoteWorkbook(client, docId, blob, 'D')
    if ('kind' in opened) throw new Error(`Open refused: ${JSON.stringify(opened)}`)
    const sheetId = opened.sheets[0]!.sheetId
    const base = { r0: 10, r1: 15, c0: 12, c1: 18 }
    const layouts = new Map(opened.layouts)
    layouts.set(sheetId, { ...opened.layouts.get(sheetId)!, used: base })
    const remote = new RemoteWorkbook(client, { ...opened, layouts })
    const sourceMerges = [
      { minRow: 31, maxRow: 31, minCol: 41, maxCol: 42 }, // inside pivot 0's saved area
      { minRow: 33, maxRow: 33, minCol: 41, maxCol: 42 }, // in the gap between pivot 0's protected rectangles
      { minRow: 26, maxRow: 26, minCol: 46, maxCol: 47 }, // inside pivot 1
      { minRow: 50, maxRow: 51, minCol: 60, maxCol: 61 }, // outside both
    ]
    const structures = (remote as any).structures as Map<string, any>
    structures.set(sheetId, { gen: 0, merges: sourceMerges, sourceMerges, state: 'done' })
    const commit = (pivot: number, protectedRects: Array<{r0:number;r1:number;c0:number;c1:number}>, merges: Array<{r0:number;r1:number;c0:number;c1:number}>) =>
      (remote as any).publishPivotCommit(sheetId, pivot, { protectedRects, overlayExtent: overlayUsedExtent(protectedRects), merges, rects: protectedRects, charts: [] }, pivot + 1)
    try {
      const firstRects = [{ r0: 30, r1: 32, c0: 40, c1: 42 }, { r0: 35, r1: 36, c0: 41, c1: 43 }]
      commit(0, firstRects, [{ r0: 30, r1: 30, c0: 40, c1: 41 }])
      expect(remote.layout(sheetId).used).toEqual({ r0: 10, r1: 36, c0: 12, c1: 43 })
      expect(remote.merges(sheetId).merges).toEqual([
        { minRow: 33, maxRow: 33, minCol: 41, maxCol: 42 },
        { minRow: 26, maxRow: 26, minCol: 46, maxCol: 47 },
        { minRow: 50, maxRow: 51, minCol: 60, maxCol: 61 },
        { minRow: 30, maxRow: 30, minCol: 40, maxCol: 41 },
      ])
      commit(1, [{ r0: 25, r1: 27, c0: 45, c1: 47 }], [{ r0: 25, r1: 25, c0: 45, c1: 46 }])
      expect(remote.layout(sheetId).used).toEqual({ r0: 10, r1: 36, c0: 12, c1: 47 })
      ;(remote as any).withdrawPivotCommit(sheetId, 0)
      expect(remote.layout(sheetId).used).toEqual({ r0: 10, r1: 27, c0: 12, c1: 47 })
      expect(remote.merges(sheetId).merges).toEqual([
        { minRow: 31, maxRow: 31, minCol: 41, maxCol: 42 },
        { minRow: 33, maxRow: 33, minCol: 41, maxCol: 42 },
        { minRow: 50, maxRow: 51, minCol: 60, maxCol: 61 },
        { minRow: 25, maxRow: 25, minCol: 45, maxCol: 46 },
      ])
      ;(remote as any).withdrawPivotCommit(sheetId, 1)
      expect(remote.layout(sheetId).used).toEqual(base)
      expect(remote.merges(sheetId).merges).toEqual(sourceMerges)
    } finally {
      await remote.close()
      await port.idle()
    }
  })

  it('the worker-parsed object route reads committed pivot values remotely without a mirror', async () => {
    const blob = new Blob([new Uint8Array(readFileSync(FIXTURE))])
    const doc = await parseXlsx(await OfficePackage.load(new Uint8Array(readFileSync(FIXTURE))))
    const sheet = doc.sheets.find(candidate => candidate.pivotTables?.length)!
    const kernel = new ViewerKernel({ parse: async bytes => parseXlsx(await OfficePackage.load(bytes)) }, quiet)
    const port = connectInProcess(kernel), client = new ViewerClient(port)
    const docId = id(), sheetId = Number(sheet.sheetId ?? '1')
    const opened = await openRemoteWorkbook(client, docId, blob, 'D')
    if ('kind' in opened) throw new Error(`Open refused: ${JSON.stringify(opened)}`)
    const remote = new RemoteWorkbook(client, opened), remoteSheetId = String(sheetId)
    try {
      expect(await remote.pivots.open(sheetId, 0, new AbortController().signal)).toMatchObject({ kind: 'open' })
      const committed = await remote.pivots.apply(sheetId, 0, { kind: 'collapse', field: 0, item: 0, collapsed: true })
      expect(committed.kind).toBe('committed')
      if (committed.kind !== 'committed') throw new Error(JSON.stringify(committed))
      const changed = committed.cells?.find(cell => cell.value !== null)
      expect(changed).toBeDefined()
      if (!changed) throw new Error('Pivot commit contained no cells')
      const read = await remote.store(remoteSheetId).readCell({ row: changed.row, col: changed.col }, {
        signal: new AbortController().signal, stamp: remote.store(remoteSheetId).stamp,
      })
      expect(read.kind).toBe('cell')
      if (read.kind === 'cell') expect(read.cell.value).toEqual(changed.value?.value)
      const workerGen = kernel.session(docId)!.source.store(remoteSheetId).stamp.valueGen
      expect(remote.identity(remoteSheetId).valueGen).toBe(workerGen)
    } finally {
      await remote.close()
      await port.idle()
    }
  })

  it('Task 29 r2: a pivot reply\'s fit cells are granted on the main ledger before it is sent and charged from arrival', async () => {
    const blob = new Blob([new Uint8Array(readFileSync(FIXTURE))])
    const doc = await parseXlsx(await OfficePackage.load(new Uint8Array(readFileSync(FIXTURE))))
    const sheet = doc.sheets.find(candidate => candidate.pivotTables?.length)!
    const kernel = new ViewerKernel({ parse: async bytes => parseXlsx(await OfficePackage.load(bytes)) }, quiet)
    const port = connectInProcess(kernel), client = new ViewerClient(port)
    const docId = id(), sheetId = Number(sheet.sheetId ?? '1')
    const opened = await openRemoteWorkbook(client, docId, blob, 'D')
    if ('kind' in opened) throw new Error(`Open refused: ${JSON.stringify(opened)}`)
    const remote = new RemoteWorkbook(client, opened)
    const sent: Array<{ kind: string; fitGrantBytes?: number; reserved: number; requestId: string }> = []
    const send = client.send.bind(client)
    ;(client as any).send = (command: WorkerCommand) => {
      if (command.kind === 'pivotOpen' || command.kind === 'pivotApply') sent.push({ kind: command.kind, fitGrantBytes: command.fitGrantBytes, requestId: command.requestId, reserved: remote.mainLedger.ownerBytes('replies', `pivot-reply:${command.requestId}`) })
      return send(command)
    }
    try {
      const result = await remote.pivots.open(sheetId, 0, new AbortController().signal)
      if (result.kind !== 'open') throw new Error(JSON.stringify(result))
      expect(sent[0]).toMatchObject({ kind: 'pivotOpen', fitGrantBytes: PIVOT_FIT_REPLY_BYTES, reserved: PIVOT_FIT_REPLY_BYTES })
      // Arrived: the grant became the reply's charge, for as long as any of it (its cells) is kept.
      expect(remote.mainLedger.ownerBytes('replies', `pivot-reply:${sent[0]!.requestId}`)).toBe(0)
      expect(result.cells?.length).toBeGreaterThan(0)
      expect(remote.mainLedger.hasLifetime(result.cells!)).toBe(true)
      // No room for the grant: the change still applies, without cells (that change is not autofitted).
      remote.mainLedger.set('copy', 'filler', 64 * 1024 * 1024)
      const applied = await remote.pivots.apply(sheetId, 0, { kind: 'collapse', field: 0, item: 0, collapsed: true })
      expect(applied.kind).toBe('committed')
      expect(sent.at(-1)).toMatchObject({ kind: 'pivotApply', fitGrantBytes: 0, reserved: 0 })
      if (applied.kind === 'committed') expect(applied.cells).toBeUndefined()
      remote.mainLedger.set('copy', 'filler', 0)
    } finally {
      await remote.close()
      await port.idle()
    }
  })

  it('a refused or superseded stage does not block a later remote apply', async () => {
    const blob = new Blob([new Uint8Array(readFileSync(FIXTURE))])
    const doc = await parseXlsx(await OfficePackage.load(new Uint8Array(readFileSync(FIXTURE))))
    const sheet = doc.sheets.find(candidate => candidate.pivotTables?.length)!
    const kernel = new ViewerKernel({ parse: async bytes => parseXlsx(await OfficePackage.load(bytes)) }, quiet)
    const port = connectInProcess(kernel), client = new ViewerClient(port), docId = id(), sheetId = Number(sheet.sheetId ?? '1')
    await client.send({ kind: 'open', requestId: client.nextRequestId(), docId, source: { kind: 'blob', blob }, profile: 'D' })
    const pivots = new RemotePivotSessions({ docId, profile: 'D', client, isClosed: () => false })
    expect(await pivots.open(sheetId, 0, new AbortController().signal)).toMatchObject({ kind: 'open' })
    const send = client.send.bind(client)
    let releaseA!: () => void, enteredA!: () => void
    const aEntered = new Promise<void>(resolve => { enteredA = resolve })
    ;(client as any).send = (command: WorkerCommand) => {
      const result = send(command)
      if (command.kind === 'pivotApply' && !releaseA) {
        const wait = new Promise<void>(resolve => { releaseA = resolve; enteredA() })
        return result.then(async reply => { await wait; return reply })
      }
      return result
    }
    const a = pivots.apply(sheetId, 0, { kind: 'collapse', field: 0, item: 0, collapsed: true })
    await aEntered
    const b = await pivots.apply(sheetId, 0, { kind: 'collapse', field: 99, item: 0, collapsed: true })
    releaseA()
    expect((await a).kind).toBe('stale')
    expect(b.kind).toBe('refused')
    const c = await pivots.apply(sheetId, 0, { kind: 'collapse', field: 0, item: 1, collapsed: true })
    expect(c.kind).toBe('committed')
    expect(kernel.session(docId)!.source.store(String(sheetId)).stamp.valueGen).toBe(1)
    pivots.close()
    await client.send({ kind: 'close', requestId: client.nextRequestId(), docId })
  })

  it('an aborted shared records decode can be reopened with a fresh signal', async () => {
    const { doc, session, sheetId } = await openObjectPivot()
    await session.pivotRelease(sheetId, 0)
    const runtime = pivotRuntimeFor(doc)!
    const records = runtime.records.bind(runtime), controller = new AbortController()
    vi.spyOn(runtime, 'records').mockImplementationOnce(async (cache, signal, limits) => {
      controller.abort()
      return records(cache, signal, limits)
    })
    expect((await session.pivotOpen(sheetId, 0, 16 * 1024 * 1024, controller.signal)).kind).toBe('cancelled')
    const reopened = await session.pivotOpen(sheetId, 0, 16 * 1024 * 1024, new AbortController().signal)
    expect(reopened.kind).toBe('open')
  })

  it('record pool admission grows with the decoded store and reports quota distinctly', async () => {
    const { doc, kernel, docId, sheetId, session } = await openObjectPivot()
    await session.pivotRelease(sheetId, 0)
    const runtime = pivotRuntimeFor(doc)!, original = runtime.records.bind(runtime)
    const runtimeRecords = vi.spyOn(runtime, 'records')
    const requested: number[] = []
    runtimeRecords.mockImplementationOnce((cache, signal, limits) => original(cache, signal, {
      ...(limits ?? RECORD_LIMITS.D), admitBytes: bytes => { requested.push(bytes); return limits?.admitBytes?.(bytes) ?? true },
    }))
    const before = session.memoryStats().sourceBytes
    expect(completed(await run(kernel, { kind: 'pivotOpen', requestId: id(), docId, sheetId, pivot: 0, grantBytes: 16 * 1024 * 1024 })).kind).toBe('pivotOpen')
    expect(requested.length).toBeGreaterThan(0)
    expect(Math.max(...requested)).toBeLessThan(RECORD_LIMITS.D.maxBytes)
    await session.pivotRelease(sheetId, 0)
    expect(session.memoryStats().sourceBytes).toBe(before)

    runtimeRecords.mockImplementationOnce((cache, signal, limits) => original(cache, signal, { ...(limits ?? RECORD_LIMITS.D), admitBytes: () => false }))
    const refused = await run(kernel, { kind: 'pivotOpen', requestId: id(), docId, sheetId, pivot: 0, grantBytes: 16 * 1024 * 1024 })
    expect(refused.kind).toBe('quota')
    expect(session.memoryStats().sourceBytes).toBe(before)
    await run(kernel, { kind: 'close', requestId: id(), docId })
  })

  it('an open cancelled after lazy session setup releases its newly created handle', async () => {
    const doc = await parseXlsx(await OfficePackage.load(new Uint8Array(readFileSync(FIXTURE))))
    const sheet = doc.sheets.find(candidate => candidate.pivotTables?.length)!
    let unblock!: () => void, entered!: () => void
    const gate = new Promise<void>(resolve => { unblock = resolve })
    const didEnter = new Promise<void>(resolve => { entered = resolve })
    const kernel = new ViewerKernel({ parse: async bytes => parseXlsx(await OfficePackage.load(bytes)),
      afterDispatch: command => command.kind === 'pivotOpen' ? (entered(), gate) : undefined }, quiet)
    const docId = id(), sheetId = sheet.sheetId ?? '1'
    await run(kernel, { kind: 'open', requestId: id(), docId, source: { kind: 'object', document: doc }, profile: 'D' })
    const session = kernel.session(docId)!, baseline = session.memoryStats().sourceBytes
    const requestId = id()
    const opening = run(kernel, { kind: 'pivotOpen', requestId, docId, sheetId, pivot: 0, grantBytes: 16 * 1024 * 1024 })
    await didEnter
    await run(kernel, { kind: 'cancel', requestId: id(), targetRequestId: requestId })
    unblock()
    expect(await opening).toMatchObject({ kind: 'cancelled' })
    expect(session.memoryStats().sourceBytes).toBe(baseline)
    expect((session as any).pivotHandles.size).toBe(0)
    await run(kernel, { kind: 'close', requestId: id(), docId })
  })

  it('publishes a worker commit acknowledgement even after a newer apply is sent', async () => {
    const blob = new Blob([new Uint8Array(readFileSync(FIXTURE))])
    const doc = await parseXlsx(await OfficePackage.load(new Uint8Array(readFileSync(FIXTURE))))
    const sheet = doc.sheets.find(candidate => candidate.pivotTables?.length)!
    const kernel = new ViewerKernel({ parse: async bytes => parseXlsx(await OfficePackage.load(bytes)) }, quiet)
    const port = connectInProcess(kernel), client = new ViewerClient(port), docId = id(), sheetId = Number(sheet.sheetId ?? '1')
    await client.send({ kind: 'open', requestId: client.nextRequestId(), docId, source: { kind: 'blob', blob }, profile: 'D' })
    const committed: number[] = []
    const pivots = new RemotePivotSessions({ docId, profile: 'D', client, isClosed: () => false,
      committed: (_sheet, _pivot, _result, valueGen) => committed.push(valueGen) })
    expect(await pivots.open(sheetId, 0, new AbortController().signal)).toMatchObject({ kind: 'open' })
    const send = client.send.bind(client)
    let b: Promise<unknown> | undefined
    ;(client as any).send = (command: WorkerCommand) => {
      const result = send(command)
      if (command.kind === 'pivotCommit' && !b) b = pivots.apply(sheetId, 0, { kind: 'collapse', field: 0, item: 1, collapsed: true })
      return result
    }
    const a = await pivots.apply(sheetId, 0, { kind: 'collapse', field: 0, item: 0, collapsed: true })
    await b
    const workerGen = kernel.session(docId)!.source.store(String(sheetId)).stamp.valueGen
    expect(workerGen).toBeGreaterThan(0)
    expect(committed).toContain(workerGen)
    expect(a.kind).toBe('stale')
    pivots.close()
    await client.send({ kind: 'close', requestId: client.nextRequestId(), docId })
  })

  it('an aborted same-cache sibling cannot release the surviving open or its runtime charge', async () => {
    const { sheet, kernel, docId, sheetId, session } = await openObjectPivot({ duplicate: true })
    await session.pivotRelease(sheetId, 0)
    const source = session.source as any, original = source.pivotAssets.bind(source)
    let releaseFirst!: () => void
    const firstGate = new Promise<void>(resolve => { releaseFirst = resolve })
    const first = new AbortController(), second = new AbortController()
    vi.spyOn(source, 'pivotAssets').mockImplementation(async (...args: unknown[]) => {
      const assets = await original(...args)
      return { ...assets, saved: async (def: { name: string }, signal: AbortSignal) => {
        const saved = await assets.saved(def, signal)
        if (def.name === sheet.pivotTables![0]!.name) await firstGate
        else second.abort()
        return saved
      } }
    })
    const baseline = session.memoryStats().sourceBytes
    const p0 = session.pivotOpen(sheetId, 0, 16 * 1024 * 1024, first.signal)
    const p1 = session.pivotOpen(sheetId, 1, 16 * 1024 * 1024, second.signal)
    expect((await p1).kind).toBe('cancelled')
    releaseFirst()
    expect((await p0).kind).toBe('open')
    const survivorBytes = session.memoryStats().sourceBytes - baseline
    expect(survivorBytes).toBeGreaterThan(0)
    expect((await session.pivotOpen(sheetId, 1, 16 * 1024 * 1024, new AbortController().signal)).kind).toBe('open')
    expect(session.memoryStats().sourceBytes - baseline).toBe(survivorBytes)
    await session.pivotRelease(sheetId, 0)
    expect(session.memoryStats().sourceBytes - baseline).toBe(survivorBytes)
    await session.pivotRelease(sheetId, 1)
    expect(session.memoryStats().sourceBytes).toBe(baseline)
    await run(kernel, { kind: 'close', requestId: id(), docId })
  })
})

describe('worker refreshOnLoad route (Task 32)', () => {
  async function refreshKernel(route: 'object' | 'paged', name: string) {
    const bytes = new Uint8Array(readFileSync(resolve(FIXTURE_ROOT, `${name}.xlsx`)))
    const docId = id()
    if (route === 'object') {
      const doc = await parseXlsx(await OfficePackage.load(bytes))
      const sheet = doc.sheets.find(candidate => candidate.pivotTables?.length)!
      const kernel = new ViewerKernel({ parse: async input => parseXlsx(await OfficePackage.load(input)) }, quiet)
      await run(kernel, { kind: 'open', requestId: id(), docId, source: { kind: 'object', document: doc }, profile: 'D' })
      return { kernel, docId, sheetId: sheet.sheetId ?? '1' }
    }
    const imported = await openImportSession({ kind: 'blob', blob: new Blob([bytes]) }, { docId, profile: 'D' })
    await imported.done
    const source = createPagedWorkbookSource(imported)
    const kernel = new ViewerKernel({ parse: async input => parseXlsx(await OfficePackage.load(input)), openSource: async () => source }, quiet)
    await run(kernel, { kind: 'open', requestId: id(), docId, source: { kind: 'paged', handleId: docId }, profile: 'D' })
    return { kernel, docId, sheetId: source.sheets.find(candidate => candidate.name === 'Pivot')!.sheetId }
  }
  const caseCells = async (name: string) => sheetCellReader((await loadPivotParts(resolve(FIXTURE_ROOT, `${name}.xlsx`)))[0]!.sheet)

  it.each([
    ['object', 'recipes/refresh-seed'], ['object', 'recipes/refresh-seed-nodata'],
    ['paged', 'recipes/refresh-seed'], ['paged', 'recipes/refresh-seed-nodata'],
  ] as const)('%s %s: rebuilt from the source through the worker on the first change, with one notice', async (route, name) => {
    const { kernel, docId, sheetId } = await refreshKernel(route, name)
    const session = kernel.session(docId)!
    const pivotSlots = () => [...((session as unknown as { memory: { slots: Map<string, unknown> } }).memory.slots.keys())].filter(key => key.startsWith('pivot-'))
    expect(completed(await run(kernel, { kind: 'pivotOpen', requestId: id(), docId, sheetId, pivot: 0, grantBytes: 16 * 1024 * 1024 }))).toMatchObject({ kind: 'pivotOpen', revision: 0 })
    expect(pivotSlots().some(key => key.startsWith('pivot-refresh-store:'))).toBe(true)
    const staged = completed(await run(kernel, { kind: 'pivotApply', requestId: id(), docId, sheetId, pivot: 0, baseRevision: 0, revision: 1,
      delta: { kind: 'collapse', field: 0, item: 0, collapsed: true }, grantBytes: 16 * 1024 * 1024 })) as unknown as { revision: number; notice?: string; cells: Array<{ row: number; col: number; value: { value: unknown } | null }> }
    expect(staged.notice).toMatch(/refreshed from the workbook's data/)
    // Excel's refresh-then-collapse (interaction/refresh-on-load-collapse), cell by cell.
    const saved = await caseCells('interaction/refresh-on-load-collapse')
    for (const cell of staged.cells) expect(cell.value?.value ?? undefined, `R${cell.row + 1}C${cell.col + 1}`).toEqual(saved(cell.row, cell.col)?.value)
    expect(completed(await run(kernel, { kind: 'pivotCommit', requestId: id(), docId, sheetId, pivot: 0, revision: staged.revision }))).toMatchObject({ kind: 'pivotCommit' })
    const next = completed(await run(kernel, { kind: 'pivotApply', requestId: id(), docId, sheetId, pivot: 0, baseRevision: staged.revision, revision: 2,
      delta: { kind: 'collapse', field: 0, item: 0, collapsed: false }, grantBytes: 16 * 1024 * 1024 })) as unknown as { notice?: string }
    expect(next.notice).toBeUndefined()
    expect(completed(await run(kernel, { kind: 'pivotRelease', requestId: id(), docId, sheetId, pivot: 0 }))).toMatchObject({ kind: 'pivotRelease', released: true })
    expect(pivotSlots()).toEqual([])
    await run(kernel, { kind: 'close', requestId: id(), docId })
  })
})

describe('worker refresh construction is bounded (Task 32 fix round 1)', () => {
  it('a sparse tall refresh source is admitted to the pool as it is built, and nothing stays charged after release', async () => {
    const doc = await parseXlsx(await OfficePackage.load(new Uint8Array(readFileSync(resolve(FIXTURE_ROOT, 'recipes/refresh-seed.xlsx')))))
    const sheet = doc.sheets.find(candidate => candidate.pivotTables?.length)!
    const cache = pivotRuntimeFor(doc)!.cache(sheet.pivotTables![0]!)!
    cache.source = { ...cache.source, worksheet: { ...cache.source.worksheet!, ref: 'A1:L100001' } }
    const kernel = new ViewerKernel({ parse: async input => parseXlsx(await OfficePackage.load(input)) }, quiet)
    const docId = id(), sheetId = sheet.sheetId ?? '1'
    await run(kernel, { kind: 'open', requestId: id(), docId, source: { kind: 'object', document: doc }, profile: 'D' })
    const session = kernel.session(docId)!
    const pool = (session as unknown as { memory: { charge(): { set(bytes: number): boolean; release(): void } } }).memory
    const charge = pool.charge.bind(pool), peaks: number[] = []
    vi.spyOn(pool, 'charge').mockImplementation(() => {
      const inner = charge(); let peak = 0
      return { set: (bytes: number) => { peak = Math.max(peak, bytes); return inner.set(bytes) }, release: () => { peaks.push(peak); inner.release() } } as never
    })
    const reply = await run(kernel, { kind: 'pivotOpen', requestId: id(), docId, sheetId, pivot: 0, grantBytes: 16 * 1024 * 1024 })
    // A typed outcome either way; the build was admitted as it grew (a charge that reached megabytes).
    if (reply.kind !== 'complete') expect(reply.kind).toBe('error')
    expect(Math.max(...peaks)).toBeGreaterThan(1024 * 1024)
    vi.restoreAllMocks()
    await run(kernel, { kind: 'pivotRelease', requestId: id(), docId, sheetId, pivot: 0 })
    const slots = [...((session as unknown as { memory: { slots: Map<string, unknown> } }).memory.slots.keys())].filter(key => key.startsWith('pivot-refresh'))
    expect(slots).toEqual([])
    await run(kernel, { kind: 'close', requestId: id(), docId })
  })
})

describe('worker refreshed cache metadata stays charged (Task 32 fix round 2)', () => {
  it('the refresh store entry holds the store and the rebuilt cache metadata until release', async () => {
    const { recordsFromSource, refreshedCacheBytes, worksheetSourceRange } = await import('../src/xlsx/pivots/session/refresh')
    const doc = await parseXlsx(await OfficePackage.load(new Uint8Array(readFileSync(resolve(FIXTURE_ROOT, 'recipes/refresh-seed.xlsx')))))
    const sheet = doc.sheets.find(candidate => candidate.pivotTables?.length)!
    const cache = pivotRuntimeFor(doc)!.cache(sheet.pivotTables![0]!)!, range = worksheetSourceRange(cache)!
    const rebuilt = recordsFromSource(cache, sheetCellReader(doc.sheets.find(s => s.name === range.sheet)!), range.ref, false, RECORD_LIMITS.D)
    if (!rebuilt.ok) throw new Error(rebuilt.diagnostic.message)
    const kernel = new ViewerKernel({ parse: async input => parseXlsx(await OfficePackage.load(input)) }, quiet)
    const docId = id(), sheetId = sheet.sheetId ?? '1'
    await run(kernel, { kind: 'open', requestId: id(), docId, source: { kind: 'object', document: doc }, profile: 'D' })
    const session = kernel.session(docId)!
    const memory = (session as unknown as { memory: { retain(key: string, options: { bytes: number }): unknown } }).memory
    const retain = vi.spyOn(memory, 'retain')
    expect(completed(await run(kernel, { kind: 'pivotOpen', requestId: id(), docId, sheetId, pivot: 0, grantBytes: 16 * 1024 * 1024 }))).toMatchObject({ kind: 'pivotOpen' })
    const store = retain.mock.calls.find(([key]) => key.startsWith('pivot-refresh-store:'))!
    expect(store[1].bytes).toBe(rebuilt.value.store.bytes + refreshedCacheBytes(rebuilt.value.cache))
    retain.mockRestore()
    await run(kernel, { kind: 'pivotRelease', requestId: id(), docId, sheetId, pivot: 0 })
    expect([...((session as unknown as { memory: { slots: Map<string, unknown> } }).memory.slots.keys())].filter(key => key.startsWith('pivot-refresh'))).toEqual([])
    await run(kernel, { kind: 'close', requestId: id(), docId })
  })
})
