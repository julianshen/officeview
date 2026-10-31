// Round 3: the lazy paged formula scan runs after the worker has already committed (viewerSession.ts pivotCommit).
import { readFileSync } from 'node:fs'
import { resolve } from 'node:path'
import { expect, it } from 'vitest'
import type { WorkerCommand, WorkerReply } from '../src/xlsx/view/contracts'
import { OfficePackage } from '../src/core/zip'
import { parseXlsx } from '../src/xlsx/parse'
import { ViewerKernel } from '../src/worker/viewerKernel'
import { createObjectWorkbookSource } from '../src/worker/viewerObjectSource'

const FIXTURE = resolve('tests/fixtures/xlsx-pivots/engine/baseline.xlsx')
let serial = 0
const id = () => `r3-${++serial}`
async function run(kernel: ViewerKernel, command: WorkerCommand): Promise<WorkerReply> {
  const replies: WorkerReply[] = []
  await kernel.handle(command, reply => replies.push(reply))
  return replies[0]!
}

it('R3-1: commit acknowledgement is immediate and cancellation/close of the later scan does not undo the installed revision', async () => {
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
