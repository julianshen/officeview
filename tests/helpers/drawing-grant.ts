// FC-T2 fixtures: drawings reads against an in-process kernel and RemoteWorkbook.
import { expect } from 'vitest'
import { ViewerKernel, connectInProcess } from '../../src/worker/viewerKernel'
import { ViewerClient } from '../../src/worker/viewerClient'
import { RemoteWorkbook } from '../../src/worker/viewerRemoteSource'
import { MAIN_THREAD_BUDGETS } from '../../src/xlsx/view/tileCache'
import { createObjectWorkbookSource } from '../../src/worker/viewerObjectSource'
import { initialIdentity, type CommandOf } from '../../src/worker/viewerProtocol'
import type { WorkerReply } from '../../src/xlsx/view/contracts'
import type { XlsxDocument } from '../../src/xlsx/types'
import { workbookOf } from './xlsx-find'

export const E = 9525
/** A sheet with one grouped chart and, optionally, one picture of `imageBytes`. */
export function drawingDoc(sheetIds: readonly string[] = ['s'], imageBytes = 0): XlsxDocument {
  const doc = workbookOf(sheetIds.map(sheetId => ({ sheetId, cells: [{ row: 0, col: 0, value: 42 }] })))
  for (const sheet of doc.sheets) {
    sheet.columnMetrics = { maxDigitWidthPx: 7, defaultWidthPx: 80 }
    sheet.drawings = [{
      xEmu: 80 * E, yEmu: 0, widthEmu: 80 * E, heightEmu: 20 * E,
      source: { partPath: 'd', treePath: 'twoCellAnchor[0]/graphicFrame[0]', element: 'graphicFrame', representation: 'native' },
      content: { kind: 'xlsxChart', chart: { groups: [{ series: [{ values: Array.from({ length: 1000 }, (_, i) => ({ state: 'value', v: i + 42 })) }] }] } } as never,
      ...(imageBytes ? { image: { data: new Uint8Array(imageBytes), mime: 'image/png', partPath: 'p.png' } } : {}),
    }]
    sheet.drawingAnchors = { unresolved: 0, records: [{ id: 'd', kind: 'twoCell', from: { col: 1, row: 0, colOffEmu: 0, rowOffEmu: 0 }, to: { col: 2, row: 1, colOffEmu: 0, rowOffEmu: 0 } }] }
  }
  return doc
}
export async function kernelFixture(imageBytes = 4096) {
  const doc = drawingDoc(['s'], imageBytes), source = createObjectWorkbookSource(doc, 'k')
  const kernel = new ViewerKernel({ parse: async () => doc, openSource: async () => source }, { yield: async () => {} })
  await kernel.handle({ kind: 'open', requestId: 'open', docId: 'k', source: { kind: 'object', document: doc }, profile: 'D' }, () => {})
  // The raw sink: what the kernel would hand to postMessage / structuredClone.
  const read = async (requestId: string, replyByteLimit?: number) => {
    const posted: WorkerReply[] = []
    const command: CommandOf<'read'> = { kind: 'read', requestId, owner: initialIdentity('k', 's'), query: { op: 'drawings', ...(replyByteLimit === undefined ? {} : { replyByteLimit }) } }
    await kernel.handle(command, out => posted.push(out))
    return { posted, command }
  }
  return { kernel, source, read }
}
export const drawingsOf = (out: WorkerReply) => out.kind === 'complete' && out.result.kind === 'read' && out.result.result.op === 'drawings' ? out.result.result.result : undefined

/** A worker-backed workbook whose drawing provider can be held per call and
 * whose reply delivery can be queued (a reply posted but not yet consumed). */
export async function remoteFixture(options: { sheets?: string[]; imageBytes?: number; empty?: boolean; hold?: boolean; queue?: boolean } = {}) {
  const sheets = options.sheets ?? ['s']
  const doc = drawingDoc(sheets, options.imageBytes ?? 0)
  if (options.empty) for (const sheet of doc.sheets) { sheet.drawings = []; sheet.drawingAnchors = { unresolved: 0, records: [] } }
  const source = createObjectWorkbookSource(doc, 'r')
  const gates: (() => void)[] = []
  if (options.hold) for (const id of sheets) {
    const store = source.store(id), read = store.readDrawings!.bind(store)
    // Ignores its signal: resumes only when the test opens the gate.
    store.readDrawings = (...args) => new Promise<void>(resolve => gates.push(resolve)).then(() => read(...args))
  }
  const kernel = new ViewerKernel({ parse: async () => doc, openSource: async () => source }, { yield: async () => {} })
  const inner = connectInProcess(kernel)
  const queued: { data: unknown; deliver: () => void }[] = []
  let delay = false
  const sent: unknown[] = []
  const port = {
    ...inner,
    postMessage(message: unknown, transfer: Transferable[]) { sent.push(message); inner.postMessage(message, transfer) },
    addEventListener(type: 'message', listener: (event: { data: unknown }) => void) {
      inner.addEventListener(type, event => { if (delay) queued.push({ data: event.data, deliver: () => listener(event) }); else listener(event) })
    },
  }
  const client = new ViewerClient(port as never)
  expect((await client.send({ kind: 'open', requestId: client.nextRequestId(), docId: 'r', source: { kind: 'object', document: doc }, profile: 'D' })).kind).toBe('complete')
  const session = kernel.session('r')!
  const remote = new RemoteWorkbook(client, { docId: 'r', valuePolicy: source.valuePolicy, sheets: source.sheets, diagnostics: [], formatting: source.formatting?.(), layouts: new Map(sheets.map(id => [id, session.layout(id)])) })
  for (const id of sheets) {
    remote.merges(id)
    for (let i = 0; i < 200 && remote.loadState(id).completeness !== 'complete'; i++) await new Promise(r => setTimeout(r, 2))
  }
  const ledger = remote.mainLedger, cap = MAIN_THREAD_BUDGETS.D
  const command = (requestId: string, sheetId = sheets[0]): CommandOf<'read'> => ({ kind: 'read', requestId, owner: remote.identity(sheetId), query: { op: 'drawings' } })
  const drawingsSent = () => sent.filter(m => (m as { kind?: string; query?: { op?: string } }).query?.op === 'drawings') as CommandOf<'read'>[]
  const settle = async () => { await inner.idle(); for (let i = 0; i < 20; i++) await Promise.resolve() }
  return {
    remote, client, kernel, source, ledger, cap, command, drawingsSent, settle, gates,
    pressure: (headroom: number) => ledger.set('dom', 'pressure', Math.max(0, cap - (ledger.bytes - ledger.ownerBytes('dom', 'pressure')) - headroom)),
    queue: (on: boolean) => { delay = on; if (!on) for (const { deliver } of queued.splice(0)) deliver() },
    /** Delivers the queued replies `pick` selects, keeping the rest queued. */
    deliver: (pick: (data: { commandKind?: string; requestId?: string }) => boolean) => {
      for (const item of [...queued]) if (pick(item.data as never)) { queued.splice(queued.indexOf(item), 1); item.deliver() }
    },
    grant: (requestId: string) => ledger.ownerBytes('replies', `drawing-grant:${requestId}`),
  }
}
export const until = async (check: () => boolean, ms = 4000) => { const end = Date.now() + ms; while (!check() && Date.now() < end) await new Promise(r => setTimeout(r, 5)) }
