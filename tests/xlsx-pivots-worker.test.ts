// @vitest-environment node
import { readFileSync } from 'node:fs'
import { fileURLToPath } from 'node:url'
import { resolve } from 'node:path'
import { describe, expect, it, vi } from 'vitest'
import { overlayUsedCounts, type WorkerCommand, type WorkerReply } from '../src/xlsx/view/contracts'
import { OfficePackage } from '../src/core/zip'
import { parseXlsx } from '../src/xlsx/parse'
import { ViewerKernel } from '../src/worker/viewerKernel'
import { ViewerClient } from '../src/worker/viewerClient'
import { connectInProcess } from '../src/worker/viewerKernel'
import { createObjectWorkbookSource } from '../src/worker/viewerObjectSource'
import { RemotePivotSessions } from '../src/worker/viewerRemotePivots'
import { pivotRuntimeFor } from '../src/xlsx/pivotLoader'
import { openImportSession } from '../src/xlsx/storage/pagedImport'
import { createPagedWorkbookSource } from '../src/xlsx/view/pagedSource'
import { sheetCellReader } from '../src/xlsx/pivots/engine/faithful'
import { openPivotSession } from '../src/xlsx/pivots/session/session'
import { ORACLE_VERIFIED_KINDS } from '../src/xlsx/pivots/session/session'
import type { PivotDelta } from '../src/xlsx/pivots/session/delta'
import { RECORD_LIMITS } from '../src/xlsx/pivots/engine/records'
import { decodeParts, loadPivotParts } from './helpers/xlsx-pivots-load'
import { RemoteWorkbook, openRemoteWorkbook } from '../src/worker/viewerRemoteSource'
import { pagedSheetGeometry } from '../src/components/xlsx/pagedGeometry'
import { createIdentityOrder } from '../src/xlsx/view/order'
import { createIntervalSet, visibilityFromHidden } from '../src/xlsx/view/visibility'

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
  it('open streams records under the profile cap and charges the worker pool', async () => {
    const doc = await parseXlsx(await OfficePackage.load(new Uint8Array(readFileSync(FIXTURE))))
    const sheet = doc.sheets.find(candidate => candidate.pivotTables?.length)!
    const kernel = new ViewerKernel({ parse: async bytes => parseXlsx(await OfficePackage.load(bytes)) }, quiet)
    const docId = id(), sheetId = sheet.sheetId ?? '1'
    await run(kernel, { kind: 'open', requestId: id(), docId, source: { kind: 'object', document: doc }, profile: 'D' })
    const session = kernel.session(docId)!
    const runtime = pivotRuntimeFor(doc)!
    const stream = vi.spyOn(runtime, 'records')
    const before = session.memoryStats().sourceBytes
    const refused = await run(kernel, { kind: 'pivotOpen', requestId: id(), docId, sheetId, pivot: 0, grantBytes: 1 })
    expect(refused.kind).toBe('error')
    expect(session.memoryStats().sourceBytes).toBe(before)
    const admitted = completed(await run(kernel, { kind: 'pivotOpen', requestId: id(), docId, sheetId, pivot: 0, grantBytes: 16 * 1024 * 1024 }))
    expect(admitted.kind).toBe('pivotOpen')
    expect(stream).toHaveBeenCalledTimes(2)
    expect(stream.mock.calls[0]?.[2]?.maxBytes).toBe(1)
    expect(stream.mock.calls[1]?.[2]?.maxBytes).toBeLessThanOrEqual(RECORD_LIMITS.D.maxBytes)
    expect(session.memoryStats().sourceBytes).toBeGreaterThan(before)
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
      kind: 'pivotApply', requestId: id(), docId, sheetId, pivot: 0, revision: 1,
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
    const first = run(kernel, { kind: 'pivotApply', requestId: id(), docId, sheetId, pivot: 0, revision: 1,
      delta: { kind: 'collapse', field: 0, item: 0, collapsed: true }, grantBytes: 16 * 1024 * 1024 })
    await waiting
    const afterFirstStage = kernel.session(docId)!.memoryStats().sourceBytes
    expect(afterFirstStage).toBeGreaterThan(before)
    const second = await run(kernel, { kind: 'pivotApply', requestId: id(), docId, sheetId, pivot: 0, revision: 2,
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
    const a = completed(await run(kernel, { kind: 'pivotApply', requestId: id(), docId, sheetId, pivot: 0, revision: 1,
      delta: { kind: 'collapse', field: 0, item: 0, collapsed: true }, grantBytes: 16 * 1024 * 1024 })) as { kind: string }
    expect(a.kind).toBe('pivotApply')
    const b = completed(await run(kernel, { kind: 'pivotApply', requestId: id(), docId, sheetId, pivot: 0, revision: 2,
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
    const sourceBytes = session.memoryStats().sourceBytes
    const otherSheet = doc.sheets.find(sheet => sheet.sheetId !== sheetId)!
    await run(kernel, { kind: 'activateSheet', requestId: id(), owner: session.workbookOwner(), sheetId: otherSheet.sheetId ?? '2' })
    expect(session.memoryStats().sourceBytes).toBeLessThan(sourceBytes)

    await run(kernel, { kind: 'pivotOpen', requestId: id(), docId, sheetId, pivot: 0, grantBytes: 16 * 1024 * 1024 })
    const openBytes = session.memoryStats().sourceBytes
    const requestId = id()
    const apply = run(kernel, { kind: 'pivotApply', requestId, docId, sheetId, pivot: 0, revision: 1,
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

  it('worker pool charges follow the overlay lifecycle', async () => {
    const { kernel, docId, sheetId, session, beforePivotOpen } = await openObjectPivot()
    const beforeOpen = beforePivotOpen
    const opened = session.memoryStats().sourceBytes
    expect(opened).toBeGreaterThan(beforeOpen)
    const sendApply = (revision: number, item: number) => run(kernel, { kind: 'pivotApply', requestId: id(), docId, sheetId, pivot: 0, revision,
      delta: { kind: 'collapse', field: 0, item, collapsed: true }, grantBytes: 16 * 1024 * 1024 })
    await sendApply(1, 0)
    const stageA = session.memoryStats().sourceBytes
    expect(stageA).toBeGreaterThan(opened)
    await sendApply(2, 1)
    const stageB = session.memoryStats().sourceBytes
    expect(stageB).toBeLessThanOrEqual(stageA)
      const stale = completed(await run(kernel, { kind: 'pivotCommit', requestId: id(), docId, sheetId, pivot: 0, revision: 1 }))
      expect(stale).toMatchObject({ kind: 'pivotCommit', refused: 'stale' })
    expect(session.memoryStats().sourceBytes).toBe(stageB)
    await run(kernel, { kind: 'pivotCommit', requestId: id(), docId, sheetId, pivot: 0, revision: 2 })
    const committedB = session.memoryStats().sourceBytes
    await sendApply(3, 2)
    const stageC = session.memoryStats().sourceBytes
    expect(stageC).toBeGreaterThan(committedB)
    await run(kernel, { kind: 'pivotCommit', requestId: id(), docId, sheetId, pivot: 0, revision: 3 })
    expect(session.memoryStats().sourceBytes).toBeLessThan(stageC)
    await run(kernel, { kind: 'pivotRelease', requestId: id(), docId, sheetId, pivot: 0 })
    expect(session.memoryStats().sourceBytes).toBeLessThan(opened)
  })

  it('pivots on one cache share one store until the last release', async () => {
    const { kernel, docId, sheetId, session } = await openObjectPivot({ duplicate: true })
    const afterFirst = session.memoryStats().sourceBytes
    const second = completed(await run(kernel, { kind: 'pivotOpen', requestId: id(), docId, sheetId, pivot: 1, grantBytes: 16 * 1024 * 1024 }))
    expect(second.kind).toBe('pivotOpen')
    expect(session.memoryStats().sourceBytes).toBe(afterFirst)
    await run(kernel, { kind: 'pivotRelease', requestId: id(), docId, sheetId, pivot: 0 })
    expect(session.memoryStats().sourceBytes).toBe(afterFirst)
    await run(kernel, { kind: 'pivotRelease', requestId: id(), docId, sheetId, pivot: 1 })
    expect(session.memoryStats().sourceBytes).toBeLessThan(afterFirst)
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
      const pagedReply = completed(await run(kernel, { kind: 'pivotApply', requestId: id(), docId, sheetId: sheet.sheetId, pivot, revision,
        delta, grantBytes: 16 * 1024 * 1024 })) as { kind: string; cells?: Array<{ row: number; col: number; value: unknown }> }
      expect(pagedReply.kind, entry.case).toBe('pivotApply')
      const expected = residentApply.value.layout.cells.map(cell => ({
        row: cell.row, col: cell.col,
        value: cell.value.kind === 'empty' ? null : {
          value: cell.value.value, ...(cell.value.kind === 'error' ? { valueIsError: true } : {}),
          ...(cell.numFmtId === undefined ? {} : { numFmtId: cell.numFmtId }),
        },
      }))
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

  it('mirrors an accepted commit to a cloned nonpaged main-thread store', async () => {
    const workerDoc = await parseXlsx(await OfficePackage.load(new Uint8Array(readFileSync(FIXTURE))))
    const mainDoc = await parseXlsx(await OfficePackage.load(new Uint8Array(readFileSync(FIXTURE))))
    const sheet = workerDoc.sheets.find(candidate => candidate.pivotTables?.length)!
    const mainSheet = mainDoc.sheets.find(candidate => candidate.pivotTables?.length)!
    const workerSource = createObjectWorkbookSource(mainDoc, 'main-clone')
    const mainStore = workerSource.store(mainSheet.sheetId ?? '1')
    const kernel = new ViewerKernel({ parse: async bytes => parseXlsx(await OfficePackage.load(bytes)) }, quiet)
    const port = connectInProcess(kernel), client = new ViewerClient(port)
    const docId = id(), sheetId = Number(sheet.sheetId ?? '1')
    await client.send({ kind: 'open', requestId: client.nextRequestId(), docId, source: { kind: 'blob', blob: new Blob([new Uint8Array(readFileSync(FIXTURE))]) }, profile: 'D' })
    const pivots = new RemotePivotSessions({
      docId, profile: 'D', client, isClosed: () => false,
      mirror: (_id, overlay) => mainStore.setValueOverlay?.(overlay),
    })
    const opened = await pivots.open(sheetId, 0, new AbortController().signal)
    expect(opened, JSON.stringify(opened)).toMatchObject({ kind: 'open' })
    if (opened.kind !== 'open') throw new Error('pivot open refused')
    const committed = await pivots.apply(sheetId, 0, { kind: 'collapse', field: 0, item: 0, collapsed: true })
    expect(committed.kind).toBe('committed')
    const changed = committed.kind === 'committed' && committed.cells.find(cell => cell.value !== null)
    expect(changed).toBeDefined()
    if (changed) {
      const original = mainSheet.rows.flatMap(row => row.cells).find(cell => cell.row === changed.row && cell.col === changed.col)
      const read = await mainStore.readCell({ row: changed.row, col: changed.col }, { signal: new AbortController().signal, stamp: mainStore.stamp })
      expect(read.kind).toBe('cell')
      if (read.kind === 'cell') expect(read.cell.value).toEqual(changed.value?.value)
      pivots.release(sheetId, 0)
      await port.idle()
      const reverted = await mainStore.readCell({ row: changed.row, col: changed.col }, { signal: new AbortController().signal, stamp: mainStore.stamp })
      if (original) { expect(reverted.kind).toBe('cell'); if (reverted.kind === 'cell') expect(reverted.cell.value).toEqual(original.value) }
      else expect(reverted.kind).toBe('blank')
    } else {
      pivots.release(sheetId, 0)
    }
    pivots.close()
    await client.send({ kind: 'close', requestId: client.nextRequestId(), docId })
  })
})
