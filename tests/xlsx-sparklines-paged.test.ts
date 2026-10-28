import { describe, expect, it } from 'vitest'
import { readFileSync } from 'node:fs'
import { resolve } from 'node:path'
import { RemoteSparklineCache, sparklineNotices, type SparklineHost } from '../src/worker/viewerRemoteSparklines'
import { initialIdentity, transferablesOf } from '../src/worker/viewerProtocol'
import type { WorkerReply, SheetSparklines } from '../src/xlsx/view/contracts'
import { OfficePackage } from '../src/core/zip'
import { parseXlsx } from '../src/xlsx/parse'
import { openImportSession } from '../src/xlsx/storage/pagedImport'
import { createPagedWorkbookSource } from '../src/xlsx/view/pagedSource'
import { SPARKLINE_CASES } from '../scripts/sparkline-refs/cases'
import { rebuildResolvedSparklines } from '../src/worker/viewerRemoteSparklines'
import { ViewerKernel } from '../src/worker/viewerKernel'

const microtask = () => new Promise(resolve => setTimeout(resolve, 0))
const fixtureBytes = (name: string) => new Uint8Array(readFileSync(resolve(process.cwd(), `tests/fixtures/xlsx-sparklines/${name}.xlsx`)))
async function resolvePaged(bytes: Uint8Array, docId: string) {
  const session = await openImportSession({ kind: 'blob', blob: new Blob([new Uint8Array(bytes).buffer as ArrayBuffer]) }, { docId, profile: 'D' })
  try {
    await session.done
    const source = createPagedWorkbookSource(session)
    const output: Record<string, ReturnType<typeof rebuildResolvedSparklines>> = {}
    const charges: number[] = []
    for (const sheet of source.sheets) {
      const store = source.store(sheet.sheetId)
      const chargeStart = charges.length
      const result = await store.readSparklines!({ signal: new AbortController().signal, stamp: store.stamp }, 64 * 1024 * 1024, { set: bytes => { charges.push(bytes); return true } })
      output[sheet.name] = rebuildResolvedSparklines(result)
      if (charges.length > chargeStart) expect(charges.at(-1)!, `${sheet.name} preflight covers final reply`).toBeLessThanOrEqual(charges[chargeStart]!)
    }
    await source.release()
    return { output, charges }
  } finally { await session.release() }
}

describe('paged sparkline transfer cache', () => {
  it('does not request until a receiver grant is available', async () => {
    const requests: unknown[] = []
    const host: SparklineHost = {
      identity: () => initialIdentity('doc', 'sheet'), loadState: () => ({ completeness: 'complete' }),
      hasSparklines: () => true, grant: () => null, releaseGrant: () => {}, admitDecoded: () => null, release: () => {},
      request: async command => { requests.push(command); return Symbol('closed') },
      nextRequestId: () => '1', changed: () => {},
    }
    const cache = new RemoteSparklineCache(host)
    expect(cache.get('sheet').state).toBe('waiting')
    await microtask()
    expect(requests).toEqual([])
    cache.close()
  })

  it('adopts the worker result only for the current sheet owner and invalidates its host tiles', async () => {
    const identity = initialIdentity('doc', 'sheet')
    const requests: unknown[] = []
    let changed = 0
    const result: SheetSparklines = {
      sparklines: [{ host: { row: 4, col: 3 }, dataSheet: 'sheet', sourceRows: [0], groupIndex: 0, values: new Float64Array([1]), bridged: [], firstIndex: 0, lastIndex: 0, highIndices: [0], lowIndices: [0], min: 1, max: 1 }],
      groups: [{ type: 'line', lineWeightPt: 0.75, displayEmptyCellsAs: 'gap', markers: false, high: false, low: false, first: false, last: false, negative: false, displayXAxis: false, displayHidden: false, dateAxis: false, rightToLeft: false, minAxisType: 'individual', maxAxisType: 'individual', colors: {}, sparklines: [] }],
      diagnostics: [],
    }
    const host: SparklineHost = {
      identity: () => identity, loadState: () => ({ completeness: 'complete' }), hasSparklines: () => true, grant: () => 4096,
      releaseGrant: () => {}, admitDecoded: () => () => {}, release: () => {},
      request: async command => { requests.push(command); return { kind: 'complete', commandKind: 'read', requestId: command.requestId, owner: identity, diagnostics: [], result: { kind: 'read', result: { op: 'sparklines', result } } } as WorkerReply<'read'> },
      nextRequestId: () => '2', changed: () => { changed++ },
    }
    const cache = new RemoteSparklineCache(host)
    cache.get('sheet')
    expect(cache.get('sheet').resolved).toBeUndefined() // pending data paints nothing
    await microtask()
    expect((requests[0] as { query: { op: string; grantBytes: number } }).query).toMatchObject({ op: 'sparklines', grantBytes: 4096 })
    expect(cache.get('sheet').state).toBe('ready')
    expect(changed).toBe(1)
    cache.close()
  })

  it('worker snapshots resolve every reference fixture and the xlsxwriter fixture like resident parsing', async () => {
    const fixtures = [...SPARKLINE_CASES.map(item => item.id), 'xlsxwriter-written']
    for (const fixture of fixtures) {
      const bytes = fixtureBytes(fixture)
      const resident = await parseXlsx(await OfficePackage.load(bytes), { valuePolicy: 'cached' })
      const { output } = await resolvePaged(bytes, `spark-worker-parity-${fixture}`)
      expect(output, fixture).toEqual(Object.fromEntries(resident.sheets.map(sheet => [sheet.name, sheet.resolvedSparklines])))
    }
  })

  it('refuses an undersized grant before taking a worker preparation charge', async () => {
    const bytes = fixtureBytes('line-basic')
    const session = await openImportSession({ kind: 'blob', blob: new Blob([new Uint8Array(bytes).buffer as ArrayBuffer]) }, { docId: 'spark-worker-refuse', profile: 'D' })
    try {
      await session.done
      const source = createPagedWorkbookSource(session), sheet = source.sheets[0], store = source.store(sheet.sheetId)
      let charges = 0
      const result = await store.readSparklines!({ signal: new AbortController().signal, stamp: store.stamp }, 1, { set: () => { charges++; return true } })
      expect(result.sparklines).toEqual([])
      expect(result.diagnostics).toHaveLength(1)
      expect(result.diagnostics[0].feature).toBe('sparkline-resource-limit')
      expect(charges).toBe(0)
      await source.release()
    } finally { await session.release() }
  })

  it('emits transferable Float64Array buffers and releases the worker preparation share', async () => {
    const bytes = fixtureBytes('date-uneven')
    const session = await openImportSession({ kind: 'blob', blob: new Blob([new Uint8Array(bytes).buffer as ArrayBuffer]) }, { docId: 'spark-worker-transfer', profile: 'D' })
    try {
      await session.done
      const source = createPagedWorkbookSource(session), sheet = source.sheets[0], store = source.store(sheet.sheetId)
      let held = 0
      const result = await store.readSparklines!({ signal: new AbortController().signal, stamp: store.stamp }, 64 * 1024 * 1024, { set: bytes => { held = bytes; return true } })
      const originals = transferablesOf(result)
      expect(result.sparklines.length).toBeGreaterThan(0)
      expect(result.sparklines[0].values).toBeInstanceOf(Float64Array)
      expect(result.sparklines[0].axis?.positions).toBeInstanceOf(Float64Array)
      expect(held).toBeGreaterThan(0)
      structuredClone(result, { transfer: originals })
      expect(originals.every(buffer => buffer.byteLength === 0)).toBe(true)
      held = 0 // WorkerSession releases this share after the postMessage sink.
      expect(held).toBe(0)
      await source.release()
    } finally { await session.release() }
  })

  it('posts 50 detached worker replies and returns transient accounting to its open baseline each time', async () => {
    const bytes = fixtureBytes('date-uneven')
    const session = await openImportSession({ kind: 'blob', blob: new Blob([new Uint8Array(bytes).buffer as ArrayBuffer]) }, { docId: 'spark-worker-50', profile: 'D' })
    await session.done
    const source = createPagedWorkbookSource(session)
    const kernel = new ViewerKernel({ parse: async () => { throw new Error('the paged source must serve this fixture') }, openSource: async () => source }, { yield: async () => {} })
    const send = async (command: Parameters<ViewerKernel['handle']>[0]) => {
      let answer: unknown
      await kernel.handle(command, reply => {
        const transfers = transferablesOf(reply)
        answer = structuredClone(reply, { transfer: transfers })
        if (command.kind === 'read') {
          expect(transfers.length).toBeGreaterThan(0)
          expect(transfers.every(buffer => buffer.byteLength === 0)).toBe(true)
        }
      })
      return answer as WorkerReply<'read' | 'open'>
    }
    try {
      const opened = await send({ kind: 'open', requestId: 'spark-open-50', docId: 'spark-worker-50', source: { kind: 'blob', blob: new Blob([bytes]) }, profile: 'D' })
      expect(opened.kind).toBe('complete')
      const baseline = kernel.lifecycleStats().transientBytes
      const owner = initialIdentity('spark-worker-50', source.sheets[0].sheetId)
      for (let i = 0; i < 50; i++) {
        const reply = await send({ kind: 'read', requestId: `spark-read-${i}`, owner, query: { op: 'sparklines', grantBytes: 64 * 1024 * 1024 } })
        expect(reply.kind).toBe('complete')
        expect(kernel.lifecycleStats().transientBytes).toBe(baseline)
      }
    } finally { await send({ kind: 'close', requestId: 'spark-close-50', docId: 'spark-worker-50' }); await session.release() }
  })

  it('releases each retained reply exactly once across stale, refused, invalidated, evicted and closed entries', async () => {
    const identityA = initialIdentity('doc', 'a'), identityB = initialIdentity('doc', 'b')
    let current = identityA, serial = 0, releases = 0, decodedReleases = 0, changed = 0
    const result: SheetSparklines = {
      sparklines: [{ host: { row: 4, col: 3 }, dataSheet: 'a', sourceRows: [0], groupIndex: 0, values: new Float64Array([1]), bridged: [], firstIndex: 0, lastIndex: 0, highIndices: [0], lowIndices: [0], min: 1, max: 1 }],
      groups: [{ type: 'line', lineWeightPt: 0.75, displayEmptyCellsAs: 'gap', markers: false, high: false, low: false, first: false, last: false, negative: false, displayXAxis: false, displayHidden: false, dateAxis: false, rightToLeft: false, minAxisType: 'individual', maxAxisType: 'individual', colors: {}, sparklines: [] }],
      diagnostics: [],
    }
    const host: SparklineHost = {
      identity: sheet => sheet === 'a' ? current : identityB,
      loadState: () => ({ completeness: 'complete' }), hasSparklines: () => true, grant: () => 4096, releaseGrant: () => {},
      admitDecoded: () => () => { decodedReleases++ },
      release: () => { releases++ }, nextRequestId: () => `${++serial}`, changed: () => { changed++ },
      request: async command => ({ kind: 'complete', commandKind: 'read', requestId: command.requestId, owner: command.owner, diagnostics: [], replyBytes: 128,
        result: { kind: 'read', result: { op: 'sparklines', result } } } as WorkerReply<'read'>),
    }
    const cache = new RemoteSparklineCache(host, 1200)
    cache.get('a'); await microtask()
    expect(cache.get('a').state).toBe('ready')
    cache.invalidate('a')
    expect(releases).toBe(1)
    expect(decodedReleases).toBe(1)
    cache.get('a'); await microtask()
    cache.get('b'); await microtask() // second reply evicts the first under the shared cap
    expect(releases).toBe(2)
    expect(decodedReleases).toBe(2)
    cache.close(); cache.close()
    expect(releases).toBe(3)
    expect(decodedReleases).toBe(3)
    expect(changed).toBeGreaterThanOrEqual(1)

    // An answer for an owner superseded while the read was in flight is released
    // rather than adopted; the generation guard must not leak its reply lease.
    let finish!: (reply: WorkerReply<'read'>) => void
    let staleIdentity = identityA
    const staleHost: SparklineHost = { ...host, identity: () => staleIdentity, release: () => { releases++ },
      request: () => new Promise(resolve => { finish = resolve }) }
    const stale = new RemoteSparklineCache(staleHost)
    stale.get('a')
    staleIdentity = { ...identityA, valueGen: identityA.valueGen + 1 }
    finish({ kind: 'complete', commandKind: 'read', requestId: 'stale', owner: identityA, diagnostics: [], replyBytes: 128,
      result: { kind: 'read', result: { op: 'sparklines', result } } } as WorkerReply<'read'>)
    await microtask()
    expect(releases).toBe(4)
    stale.close()
    expect(releases).toBe(4)

    let refusedReleases = 0
    const refused = new RemoteSparklineCache({ ...host,
      release: () => { refusedReleases++ },
      request: async command => ({ kind: 'quota', commandKind: 'read', requestId: command.requestId, owner: command.owner,
        resource: 'memory', requiredBytes: 8192, limitBytes: 4096, diagnostics: [] }) as WorkerReply<'read'>,
    })
    refused.get('a'); await microtask()
    const refusal = refused.get('a')
    expect(refusal).toMatchObject({ state: 'failed', diagnostics: [{ feature: 'sparkline-resource-limit' }] })
    expect(sparklineNotices(refusal)).toEqual(['Sparkline reply exceeded its receiver grant'])
    refused.close()
    expect(refusedReleases).toBe(1)

    let cancelledGrantRelease = 0, cancelledReplyRelease = 0
    const cancelled = new RemoteSparklineCache({ ...host,
      grant: () => 4096,
      releaseGrant: () => { cancelledGrantRelease++ },
      release: () => { cancelledReplyRelease++ },
      // The request boundary owns and refunds a grant when cancellation settles
      // without a worker reply; the cache receives no payload to lease/release.
      request: async () => {
        cancelledGrantRelease++
        return Symbol('cancelled')
      },
    })
    cancelled.get('a'); await microtask()
    expect(cancelledGrantRelease).toBe(1)
    expect(cancelledReplyRelease).toBe(0)
    cancelled.close()
    expect(cancelledReplyRelease).toBe(0)
  })
})
