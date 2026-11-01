// Dependent values on the paged (worker) path (Task 31; rulings S3-S, S3-AA, S3-AD, S3-AG): formulas are never
// evaluated there, so a pivot commit shows a sheet notice when the workbook has formulas outside its pivots. The flag
// comes from one lazy formula scan per document that never sits between a commit and its acknowledgement.
import { expect, it, vi } from 'vitest'
import { readFileSync } from 'node:fs'
import { resolve } from 'node:path'
import { OfficePackage } from '../src/core/zip'
import { parseXlsx } from '../src/xlsx/parse'
import { openPagedWorkbook } from '../src/components/xlsx/pagedWorkbook'
import { FORCE_PAGED_LIMITS } from '../src/debug/forcePaged'
import { ViewerKernel, connectInProcess } from '../src/worker/viewerKernel'
import { ViewerClient } from '../src/worker/viewerClient'
import { createObjectWorkbookSource } from '../src/worker/viewerObjectSource'
import { RemotePivotSessions } from '../src/worker/viewerRemotePivots'
import { markPagedDependentsStale } from '../src/xlsx/formula/workbook'
import type { WorkerCommand, WorkerReply } from '../src/xlsx/view/contracts'
import { SEED, load } from './helpers/xlsx-pivots-dependent-fixtures'

const FIXTURE = resolve(__dirname, 'fixtures/xlsx-pivots/engine/baseline.xlsx')
let serial = 0
const id = () => `paged-dependents-${++serial}`
async function run(kernel: ViewerKernel, command: WorkerCommand): Promise<WorkerReply> {
  const replies: WorkerReply[] = []
  await kernel.handle(command, reply => replies.push(reply))
  return replies[0]!
}
const settle = () => new Promise(resolve => setTimeout(resolve, 5))

function workerHarness() {
  const kernel = new ViewerKernel({ parse: async (bytes: Uint8Array) => parseXlsx(await OfficePackage.load(bytes)), importLimits: FORCE_PAGED_LIMITS }, { yield: async () => {} })
  const listeners = new Set<(event: MessageEvent) => void>()
  return { port: {
    postMessage(message: unknown, transfer: Transferable[]) {
      const open = message as { kind: string; source?: { kind: string; blob?: Blob } }
      const blob = open.kind === 'open' && open.source?.kind === 'blob' ? open.source.blob : undefined
      const command = structuredClone(blob ? { ...open, source: { kind: 'blob' } } : message, { transfer: transfer as ArrayBuffer[] }) as WorkerCommand
      if (blob && command.kind === 'open') (command as { source: unknown }).source = { kind: 'blob', blob }
      void kernel.handle(command, (out: WorkerReply, replyTransfer) => {
        const data = structuredClone(out, { transfer: replyTransfer })
        for (const listener of [...listeners]) listener({ data } as MessageEvent)
      })
    },
    addEventListener(type: string, listener: (event: MessageEvent) => void) { if (type === 'message') listeners.add(listener) },
    removeEventListener(type: string, listener: (event: MessageEvent) => void) { if (type === 'message') listeners.delete(listener) },
    terminate: vi.fn(),
  } }
}

it('S3-S: the paged shell sheet the viewer hands markPagedDependentsStale shows the notice from the workbook flag', async () => {
  const resident = await load(SEED)
  const formulas = resident.sheets.findIndex(s => s.name === 'Formulas')
  // Resident sheet: the predicate fires (control).
  expect(markPagedDependentsStale(new Set(), formulas, resident.sheets[formulas]!, [])).toEqual(new Set([formulas]))
  const h = workerHarness()
  const doc = await openPagedWorkbook(new Blob([new Uint8Array(readFileSync(SEED))]), { port: h.port as never, profile: 'D', signal: new AbortController().signal })
  expect(doc).toBeTruthy()
  // The open no longer page-walks the workbook. Its first pivot commit supplies
  // the cached document-level scan result, which the viewer uses for this notice.
  expect(doc!.hasFormulaCellsOutsidePivots).toBeUndefined()
  expect(markPagedDependentsStale(new Set(), formulas, doc!.sheets[formulas]!, [], true)).toEqual(new Set([formulas]))
})

// ---- the lazy workbook formula scan (S3-AD, S3-AG) -------------------------------------------------------------------
it('S3-AD: a commit is acknowledged at once; cancelling a status request or closing never undoes the installed revision', async () => {
  const doc = await parseXlsx(await OfficePackage.load(new Uint8Array(readFileSync(FIXTURE))))
  const sheet = doc.sheets.find(candidate => candidate.pivotTables?.length)!
  const docId = id(), source = createObjectWorkbookSource(doc, docId)
  let scanStarted!: () => void
  const started = new Promise<void>(resolve => { scanStarted = resolve })
  const scanSignals: AbortSignal[] = []
  let finishScan!: (value: boolean) => void
  // A large paged workbook's scan remains pending across later commits (S3-AG: they reuse it, never restart it).
  source.scanFormulaCellsOutsidePivots = (signal?: AbortSignal) => new Promise<boolean | undefined>(resolve => {
    scanSignals.push(signal!)
    finishScan = resolve
    scanStarted(); signal?.addEventListener('abort', () => resolve(undefined))
  })
  const kernel = new ViewerKernel({ parse: async bytes => parseXlsx(await OfficePackage.load(bytes)), openSource: async () => source }, { yield: async () => {} })
  const sheetId = sheet.sheetId ?? '1'
  await run(kernel, { kind: 'open', requestId: id(), docId, source: { kind: 'paged', handleId: docId }, profile: 'D' })
  await run(kernel, { kind: 'pivotOpen', requestId: id(), docId, sheetId, pivot: 0, grantBytes: 16 * 1024 * 1024 })
  const staged = await run(kernel, { kind: 'pivotApply', requestId: id(), docId, sheetId, pivot: 0, baseRevision: 0, revision: 1,
    delta: { kind: 'collapse', field: 0, item: 0, collapsed: true }, grantBytes: 16 * 1024 * 1024 })
  if (staged.kind !== 'complete' || staged.result.kind !== 'pivotApply') throw new Error(JSON.stringify(staged))
  const store = (kernel.session(docId)! as unknown as { store(id: string): { stamp: { valueGen: number } } }).store(sheetId)
  const valueGenBefore = store.stamp.valueGen
  const commitId = id()
  const commit = run(kernel, { kind: 'pivotCommit', requestId: commitId, docId, sheetId, pivot: 0, revision: staged.result.revision })
  const reply = await Promise.race([commit, new Promise<never>((_, reject) => setTimeout(() => reject(new Error('commit ack waited for formula scan')), 100))])
  await started
  const statusId = id()
  const status = run(kernel, { kind: 'pivotFormulaStatus', requestId: statusId, docId, sheetId })
  await run(kernel, { kind: 'cancel', requestId: id(), targetRequestId: statusId } as WorkerCommand)
  expect((await status).kind).toBe('cancelled')
  const stagedAgain = await run(kernel, { kind: 'pivotApply', requestId: id(), docId, sheetId, pivot: 0, baseRevision: 1, revision: 2,
    delta: { kind: 'collapse', field: 0, item: 0, collapsed: false }, grantBytes: 16 * 1024 * 1024 })
  if (stagedAgain.kind !== 'complete' || stagedAgain.result.kind !== 'pivotApply') throw new Error(JSON.stringify(stagedAgain))
  const nextCommit = await run(kernel, { kind: 'pivotCommit', requestId: id(), docId, sheetId, pivot: 0, revision: stagedAgain.result.revision })
  expect(nextCommit.kind).toBe('complete')
  // Ruling S3-AG: the newer commit reuses the running scan; neither the cancelled status request nor the commit aborts it.
  expect(scanSignals.length).toBe(1)
  expect(scanSignals[0]!.aborted).toBe(false)
  const pendingFlag = run(kernel, { kind: 'pivotFormulaStatus', requestId: id(), docId, sheetId })
  finishScan(true)
  const flag = await pendingFlag
  expect(flag).toMatchObject({ kind: 'complete', result: { kind: 'pivotFormulaStatus', hasFormulaCellsOutsidePivots: true } })
  const workerCommitted = store.stamp.valueGen !== valueGenBefore
  // Ack is already complete before scan settles; closing cancels only the background scan.
  expect({ reply: reply.kind, workerCommitted }).toEqual({ reply: 'complete', workerCommitted: true })
  await run(kernel, { kind: 'close', requestId: id(), docId })
})

async function pagedKernel() {
  const doc = await parseXlsx(await OfficePackage.load(new Uint8Array(readFileSync(FIXTURE))))
  const sheet = doc.sheets.find(candidate => candidate.pivotTables?.length)!
  const docId = id(), source = createObjectWorkbookSource(doc, docId)
  const scan = { started: 0, aborts: 0, finish: undefined as undefined | ((value: boolean) => void) }
  source.scanFormulaCellsOutsidePivots = (signal?: AbortSignal) => new Promise<boolean | undefined>(done => {
    scan.started++; scan.finish = done
    signal?.addEventListener('abort', () => { scan.aborts++; done(undefined) })
  })
  const kernel = new ViewerKernel({ parse: async bytes => parseXlsx(await OfficePackage.load(bytes)), openSource: async () => source }, { yield: async () => {} })
  await run(kernel, { kind: 'open', requestId: id(), docId, source: { kind: 'paged', handleId: docId }, profile: 'D' })
  return { kernel, docId, sheetId: sheet.sheetId ?? '1', scan }
}

it('S3-AG: three quick commits share one scan, abort none, and the status arrives', async () => {
  const { kernel, docId, sheetId, scan } = await pagedKernel()
  await run(kernel, { kind: 'pivotOpen', requestId: id(), docId, sheetId, pivot: 0, grantBytes: 16 * 1024 * 1024 })
  let committed = 0
  for (let revision = 1; revision <= 3; revision++) {
    const staged = await run(kernel, { kind: 'pivotApply', requestId: id(), docId, sheetId, pivot: 0, baseRevision: committed, revision, delta: { kind: 'collapse', field: 0, item: 0, collapsed: revision % 2 === 1 }, grantBytes: 16 * 1024 * 1024 })
    if (staged.kind !== 'complete' || staged.result.kind !== 'pivotApply') throw new Error(JSON.stringify(staged))
    expect((await run(kernel, { kind: 'pivotCommit', requestId: id(), docId, sheetId, pivot: 0, revision: staged.result.revision })).kind).toBe('complete')
    committed = staged.result.revision
    await settle()
  }
  expect({ scans: scan.started, aborts: scan.aborts }).toEqual({ scans: 1, aborts: 0 })
  // Each commit's status request awaits the same scan; the first one asked still gets the answer.
  const first = run(kernel, { kind: 'pivotFormulaStatus', requestId: id(), docId, sheetId })
  const last = run(kernel, { kind: 'pivotFormulaStatus', requestId: id(), docId, sheetId })
  scan.finish!(true)
  for (const reply of [await first, await last]) expect(reply).toMatchObject({ kind: 'complete', result: { kind: 'pivotFormulaStatus', hasFormulaCellsOutsidePivots: true } })
  // Cached for the document: a later commit starts no scan.
  const staged = await run(kernel, { kind: 'pivotApply', requestId: id(), docId, sheetId, pivot: 0, baseRevision: committed, revision: 4, delta: { kind: 'collapse', field: 0, item: 0, collapsed: true }, grantBytes: 16 * 1024 * 1024 })
  if (staged.kind !== 'complete' || staged.result.kind !== 'pivotApply') throw new Error(JSON.stringify(staged))
  await run(kernel, { kind: 'pivotCommit', requestId: id(), docId, sheetId, pivot: 0, revision: staged.result.revision })
  await settle()
  expect({ scans: scan.started, aborts: scan.aborts }).toEqual({ scans: 1, aborts: 0 })
  await run(kernel, { kind: 'close', requestId: id(), docId })
})

it('S3-AG: only the document close aborts an in-flight scan', async () => {
  const { kernel, docId, sheetId, scan } = await pagedKernel()
  await run(kernel, { kind: 'pivotOpen', requestId: id(), docId, sheetId, pivot: 0, grantBytes: 16 * 1024 * 1024 })
  const staged = await run(kernel, { kind: 'pivotApply', requestId: id(), docId, sheetId, pivot: 0, baseRevision: 0, revision: 1, delta: { kind: 'collapse', field: 0, item: 0, collapsed: true }, grantBytes: 16 * 1024 * 1024 })
  if (staged.kind !== 'complete' || staged.result.kind !== 'pivotApply') throw new Error(JSON.stringify(staged))
  await run(kernel, { kind: 'pivotCommit', requestId: id(), docId, sheetId, pivot: 0, revision: staged.result.revision })
  await settle()
  // Releasing the pivot is not a close: the workbook-level scan keeps running.
  await run(kernel, { kind: 'pivotRelease', requestId: id(), docId, sheetId, pivot: 0 })
  await settle()
  expect({ scans: scan.started, aborts: scan.aborts }).toEqual({ scans: 1, aborts: 0 })
  await run(kernel, { kind: 'close', requestId: id(), docId })
  expect({ scans: scan.started, aborts: scan.aborts }).toEqual({ scans: 1, aborts: 1 })
})

it('a committed sheet gets its notice status after its pivot entry is released, and keeps it across a reopen', async () => {
  const { kernel, docId, sheetId, scan } = await pagedKernel()
  const client = new ViewerClient(connectInProcess(kernel))
  const sent: string[] = []
  const send = client.send.bind(client)
  ;(client as unknown as { send: typeof send }).send = ((command: WorkerCommand) => { sent.push(command.kind); return send(command as never) }) as typeof send
  const flags: Array<[number, boolean]> = []
  const pivots = new RemotePivotSessions({ docId, profile: 'D', client, isClosed: () => false, formulaFlagChanged: (sheet, flag) => flags.push([sheet, flag]) })
  const sheet = Number(sheetId)
  expect(await pivots.open(sheet, 0, new AbortController().signal)).toMatchObject({ kind: 'open' })
  expect((await pivots.apply(sheet, 0, { kind: 'collapse', field: 0, item: 0, collapsed: true })).kind).toBe('committed')
  // A sheet switch right after the commit releases the entry before the scan answers.
  pivots.release(sheet, 0)
  await settle()
  scan.finish!(true)
  await settle()
  expect(flags).toEqual([[sheet, true]])
  // Reopened and changed again: the document's flag is known, so no second status request is needed.
  const statusRequests = sent.filter(kind => kind === 'pivotFormulaStatus').length
  expect(await pivots.open(sheet, 0, new AbortController().signal)).toMatchObject({ kind: 'open' })
  expect((await pivots.apply(sheet, 0, { kind: 'collapse', field: 0, item: 1, collapsed: true })).kind).toBe('committed')
  expect(flags).toEqual([[sheet, true], [sheet, true]])
  expect(sent.filter(kind => kind === 'pivotFormulaStatus').length).toBe(statusRequests)
  expect(scan.started).toBe(1)
  pivots.close()
  await client.send({ kind: 'close', requestId: client.nextRequestId(), docId })
})

it('a cancel that lands after the commit was installed still acknowledges the commit', async () => {
  const { kernel, docId, sheetId } = await pagedKernel()
  await run(kernel, { kind: 'pivotOpen', requestId: id(), docId, sheetId, pivot: 0, grantBytes: 16 * 1024 * 1024 })
  const staged = await run(kernel, { kind: 'pivotApply', requestId: id(), docId, sheetId, pivot: 0, baseRevision: 0, revision: 1, delta: { kind: 'collapse', field: 0, item: 0, collapsed: true }, grantBytes: 16 * 1024 * 1024 })
  if (staged.kind !== 'complete' || staged.result.kind !== 'pivotApply') throw new Error(JSON.stringify(staged))
  const store = (kernel.session(docId)! as unknown as { store(id: string): { stamp: { valueGen: number } } }).store(sheetId)
  const before = store.stamp.valueGen, commitId = id()
  const commit = run(kernel, { kind: 'pivotCommit', requestId: commitId, docId, sheetId, pivot: 0, revision: staged.result.revision })
  // The microtask window between the install and the reply's emission.
  await Promise.resolve()
  await run(kernel, { kind: 'cancel', requestId: id(), targetRequestId: commitId } as WorkerCommand)
  const reply = await commit
  expect({ reply: reply.kind, installed: store.stamp.valueGen !== before }).toEqual({ reply: 'complete', installed: true })
  await run(kernel, { kind: 'close', requestId: id(), docId })
})
