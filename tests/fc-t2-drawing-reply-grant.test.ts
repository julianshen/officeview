// @vitest-environment node
// FC-T2: a drawings read carries a receiving grant main reserved before it was
// sent; the worker never posts a reply over that grant.
import { describe, expect, it, vi } from 'vitest'
import { ViewerKernel } from '../src/worker/viewerKernel'
import { createObjectWorkbookSource } from '../src/worker/viewerObjectSource'
import { initialIdentity, transferablesOf, type CommandOf } from '../src/worker/viewerProtocol'
import { replyLimitEnvelope } from '../src/worker/viewerReplyLimit'
import { estimatePayloadBytes } from '../src/xlsx/view/payloadBytes'
import { estimatePayloadBytes as ledgerEstimate } from '../src/xlsx/view/mainLedger'
import type { WorkerReply } from '../src/xlsx/view/contracts'
import type { XlsxDocument } from '../src/xlsx/types'
import { workbookOf } from './helpers/xlsx-find'

const E = 9525
/** A sheet with one grouped chart and, optionally, one picture of `imageBytes`. */
export function drawingDoc(sheetIds: readonly string[] = ['s'], imageBytes = 0): XlsxDocument {
  const doc = workbookOf(sheetIds.map(sheetId => ({ sheetId, cells: [{ row: 0, col: 0, value: 42 }] })))
  for (const sheet of doc.sheets) {
    sheet.columnMetrics = { maxDigitWidthPx: 7, defaultWidthPx: 80 }
    sheet.drawings = [{
      xEmu: 80 * E, yEmu: 0, widthEmu: 80 * E, heightEmu: 20 * E,
      source: { partPath: 'd', treePath: 'twoCellAnchor[0]/graphicFrame[0]', element: 'graphicFrame', representation: 'native' },
      content: { kind: 'xlsxChart', chart: { groups: [{ series: [{ values: Array.from({ length: 1000 }, (_, i) => i + 42) }] }] } } as never,
      ...(imageBytes ? { image: { data: new Uint8Array(imageBytes), mime: 'image/png', partPath: 'p.png' } } : {}),
    }]
    sheet.drawingAnchors = { unresolved: 0, records: [{ id: 'd', kind: 'twoCell', from: { col: 1, row: 0, colOffEmu: 0, rowOffEmu: 0 }, to: { col: 2, row: 1, colOffEmu: 0, rowOffEmu: 0 } }] }
  }
  return doc
}
async function kernelFixture(imageBytes = 4096) {
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
const drawingsOf = (out: WorkerReply) => out.kind === 'complete' && out.result.kind === 'read' && out.result.result.op === 'drawings' ? out.result.result.result : undefined

describe('shared reply measure', () => {
  it('is equal on the sender object and on its transferred structured clone, counting whole backing buffers once', () => {
    const padded = new Uint8Array(64 * 1024).subarray(100, 200)
    const shared = new ArrayBuffer(8192)
    const out = {
      kind: 'complete', requestId: 'r-1', commandKind: 'read', owner: initialIdentity('d', 's'),
      result: { kind: 'read', result: { op: 'drawings', result: {
        completeness: 'complete', bytes: 0, diagnostics: [], drawings: [{ image: { data: padded } }, { a: new Uint8Array(shared, 0, 16) }, { b: new Float64Array(shared, 8, 4) }],
        images: [], theme: new Map([['accent1', 'ff0000']]),
      } } },
      // The kernel appends cached-value diagnostics to the final reply.
      diagnostics: [{ code: 'cached-value-policy', message: 'Values are the cached values saved with the file' }],
    }
    const before = estimatePayloadBytes(out)
    expect(before).toBeGreaterThanOrEqual(64 * 1024 + 8192)
    const clone = structuredClone(out, { transfer: transferablesOf(out) })
    expect(estimatePayloadBytes(clone)).toBe(before)
    expect(ledgerEstimate).toBe(estimatePayloadBytes)
  })
})

describe('kernel reply grant', () => {
  it('posts a final reply that measures exactly the grant and replaces one byte less with a payload-free quota reply', async () => {
    const { read } = await kernelFixture()
    const { posted: [full] } = await read('d0')
    const D = estimatePayloadBytes(full)
    expect(drawingsOf(full)?.completeness).toBe('complete')

    const exact = await read('d1', D)
    expect(exact.posted).toHaveLength(1)
    expect(drawingsOf(exact.posted[0])?.completeness).toBe('complete')
    expect(estimatePayloadBytes(exact.posted[0])).toBe(D)

    const under = await read('d2', D - 1)
    expect(under.posted).toHaveLength(1)
    expect(under.posted[0]).toMatchObject({ kind: 'quota', requestId: 'd2', resource: 'memory', requiredBytes: D, limitBytes: D - 1, diagnostics: [] })
    expect(transferablesOf(under.posted[0])).toEqual([])
    expect(estimatePayloadBytes(under.posted[0])).toBeLessThanOrEqual(replyLimitEnvelope(under.command))
  })

  it('checks the final reply, after the cached-value diagnostics were added', async () => {
    const { source, read } = await kernelFixture()
    const { posted: [plain] } = await read('d0')
    const D = estimatePayloadBytes(plain)
    vi.spyOn(source as unknown as { valueStatus(): unknown }, 'valueStatus').mockReturnValue({ policy: 'cached', message: 'Cached values are shown' })
    const { posted: [diagnosed] } = await read('d1')
    expect(diagnosed.diagnostics.some(d => d.code === 'cached-value-policy')).toBe(true)
    const final = estimatePayloadBytes(diagnosed)
    expect(final).toBeGreaterThan(D)
    // The provider's payload alone fits D; the diagnosed reply does not.
    const { posted: [out] } = await read('d2', final - 1)
    expect(out).toMatchObject({ kind: 'quota', requiredBytes: final })
  })

  it('refuses a grant below the payload-free envelope before it invokes the provider', async () => {
    const { source, read } = await kernelFixture()
    const provider = vi.spyOn(source.store('s'), 'readDrawings')
    const { posted, command } = await read('d0', 10)
    expect(provider).not.toHaveBeenCalled()
    expect(posted).toMatchObject([{ kind: 'quota', requestId: 'd0', diagnostics: [] }])
    const envelope = replyLimitEnvelope(command)
    await read('d1', envelope)
    expect(provider).toHaveBeenCalledTimes(1)
  })
})
