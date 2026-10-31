// Round 4: the paged workbook formula scan is one per document (Ruling S3-AG), and a committed sheet's status survives
// its pivot entry's release and reopen (M-5).
import { readFileSync } from 'node:fs'
import { resolve } from 'node:path'
import { expect, it } from 'vitest'
import type { WorkerCommand, WorkerReply } from '../src/xlsx/view/contracts'
import { OfficePackage } from '../src/core/zip'
import { parseXlsx } from '../src/xlsx/parse'
import { ViewerKernel, connectInProcess } from '../src/worker/viewerKernel'
import { ViewerClient } from '../src/worker/viewerClient'
import { createObjectWorkbookSource } from '../src/worker/viewerObjectSource'
import { RemotePivotSessions } from '../src/worker/viewerRemotePivots'

const FIXTURE = resolve('tests/fixtures/xlsx-pivots/engine/baseline.xlsx')
let serial = 0
const id = () => `r4-${++serial}`
async function run(kernel: ViewerKernel, command: WorkerCommand): Promise<WorkerReply> {
  const replies: WorkerReply[] = []
  await kernel.handle(command, reply => replies.push(reply))
  return replies[0]!
}
const settle = () => new Promise(resolve => setTimeout(resolve, 5))

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

it('R4-3 / S3-AG: three quick commits share one scan, abort none, and the status arrives', async () => {
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

it('M-5: a committed sheet gets its status after its pivot entry is released, and keeps it across a reopen', async () => {
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

it('M-2: a cancel that lands after the commit was installed still acknowledges the commit', async () => {
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
