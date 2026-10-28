import { describe, expect, it, vi } from 'vitest'
import { readFileSync } from 'node:fs'
import { resolve } from 'node:path'
import JSZip from 'jszip'
import { RemoteSparklineCache, sheetWithSparklines, sheetWithoutResolvedSparklines, sparklineNotices, type SparklineHost } from '../src/worker/viewerRemoteSparklines'
import { initialIdentity, transferablesOf } from '../src/worker/viewerProtocol'
import type { WorkerCommand, WorkerReply, SheetSparklines } from '../src/xlsx/view/contracts'
import type { XlsxSparklineGroup } from '../src/xlsx/sparklines/model'
import { OfficePackage } from '../src/core/zip'
import { parseXlsx } from '../src/xlsx/parse'
import { openImportSession } from '../src/xlsx/storage/pagedImport'
import { createPagedWorkbookSource } from '../src/xlsx/view/pagedSource'
import { SPARKLINE_CASES } from '../scripts/sparkline-refs/cases'
import { rebuildResolvedSparklines } from '../src/worker/viewerRemoteSparklines'
import { ViewerKernel } from '../src/worker/viewerKernel'
import { ViewerClient } from '../src/worker/viewerClient'
import { RemoteWorkbook } from '../src/worker/viewerRemoteSource'
import { sparklineInvalidationBox } from '../src/xlsx/view/sparklineInvalidation'
import * as sparklineGroupsParser from '../src/xlsx/sparklines/parseGroups'
import type { XmlNode } from '../src/core/xml'

const microtask = () => new Promise(resolve => setTimeout(resolve, 0))
const fixtureBytes = (name: string) => new Uint8Array(readFileSync(resolve(process.cwd(), `tests/fixtures/xlsx-sparklines/${name}.xlsx`)))
async function datedEmptyMembersBytes() {
  const zip = await JSZip.loadAsync(fixtureBytes('date-uneven'))
  const path = 'xl/worksheets/sheet1.xml'
  let xml = await zip.file(path)!.async('string')
  for (const col of 'ABCDEFGH') {
    xml = xml.replace(new RegExp(`<c r="${col}1"[^>]*>.*?<v>.*?</v>.*?</c>`), `<c r="${col}1" t="e"><v>#N/A</v></c>`)
    xml = xml.replace(new RegExp(`<c r="${col}2"[^>]*>.*?</c>`), `<c r="${col}2"></c>`)
  }
  zip.file(path, xml)
  return zip.generateAsync({ type: 'uint8array' })
}
function excelCol(zeroBased: number): string {
  let value = zeroBased + 1, label = ''
  while (value > 0) { value--; label = String.fromCharCode(65 + value % 26) + label; value = Math.floor(value / 26) }
  return label
}
async function manyLongDateRangeSparklinesBytes(count: number) {
  const zip = await JSZip.loadAsync(fixtureBytes('date-uneven'))
  const path = 'xl/worksheets/sheet1.xml'
  let xml = await zip.file(path)!.async('string')
  const sparklines = Array.from({ length: count }, (_, index) =>
    `<x14:sparkline><xm:f>Sheet1!A1</xm:f><xm:sqref>${excelCol(9 + index)}1</xm:sqref></x14:sparkline>`,
  ).join('')
  const groups = `<x14:sparklineGroups xmlns:xm="http://schemas.microsoft.com/office/excel/2006/main"><x14:sparklineGroup dateAxis="1"><xm:f>Sheet1!A7:AZ7</xm:f><x14:sparklines>${sparklines}</x14:sparklines></x14:sparklineGroup></x14:sparklineGroups>`
  xml = xml.replace(/<x14:sparklineGroups[^>]*>.*?<\/x14:sparklineGroups>/, groups)
  zip.file(path, xml)
  return zip.generateAsync({ type: 'uint8array' })
}
async function resolvePaged(bytes: Uint8Array, docId: string) {
  const session = await openImportSession({ kind: 'blob', blob: new Blob([new Uint8Array(bytes).buffer as ArrayBuffer]) }, { docId, profile: 'D' })
  try {
    await session.done
    const source = createPagedWorkbookSource(session)
    const output: Record<string, { resolved: ReturnType<typeof rebuildResolvedSparklines>; diagnostics: unknown }> = {}
    const charges: number[] = []
    for (const sheet of source.sheets) {
      const store = source.store(sheet.sheetId)
      const chargeStart = charges.length
      const result = await store.readSparklines!({ signal: new AbortController().signal, stamp: store.stamp }, 64 * 1024 * 1024, { set: bytes => { charges.push(bytes); return true } })
      output[sheet.name] = { resolved: rebuildResolvedSparklines(result), diagnostics: result.diagnostics }
      if (charges.length > chargeStart) expect(charges.at(-1)!, `${sheet.name} preflight covers final reply`).toBeLessThanOrEqual(charges[chargeStart]!)
    }
    await source.release()
    return { output, charges }
  } finally { await session.release() }
}

describe('paged sparkline transfer cache', () => {
  it('never captures colour-like elements other than the eight sparkline colours', async () => {
    const zip = await JSZip.loadAsync(fixtureBytes('line-basic'))
    const path = 'xl/worksheets/sheet1.xml'
    const unknown = Array.from({ length: 2_000 }, (_, i) => `<x14:colorIgnored${i} rgb="FFFF0000"/>`).join('')
    zip.file(path, (await zip.file(path)!.async('string')).replace('<x14:sparklines>', `${unknown}<x14:sparklines>`))
    const bytes = await zip.generateAsync({ type: 'uint8array' })
    const resident = await parseXlsx(await OfficePackage.load(bytes))
    const captured: XmlNode[] = []
    const parseGroups = sparklineGroupsParser.parseSparklineGroups
    const spy = vi.spyOn(sparklineGroupsParser, 'parseSparklineGroups').mockImplementation((node, color, diagnostics) => {
      captured.push(node)
      return parseGroups(node, color, diagnostics)
    })
    const session = await openImportSession({ kind: 'blob', blob: new Blob([new Uint8Array(bytes).buffer as ArrayBuffer]) }, { docId: 'spark-unknown-colours', profile: 'D' })
    try {
      await session.done
      const source = createPagedWorkbookSource(session), sheet = source.sheets[0]
      const store = source.store(sheet.sheetId) as unknown as { sheet: { metadata: { sparklineGroups: XlsxSparklineGroup[] } } }
      expect(store.sheet.metadata.sparklineGroups).toEqual(resident.sheets[0].sparklineGroups)
      const root = captured.at(-1)!
      const group = (Array.isArray(root.sparklineGroup) ? root.sparklineGroup[0] : root.sparklineGroup) as XmlNode
      expect(Object.keys(group).filter(key => key.startsWith('colorIgnored'))).toEqual([])
      await source.release()
    } finally {
      spy.mockRestore()
      await session.release()
    }
  })

  it('bounds duplicate sparkline singleton children and matches the resident first-match parse', async () => {
    const zip = await JSZip.loadAsync(fixtureBytes('line-basic'))
    const path = 'xl/worksheets/sheet1.xml'
    let xml = await zip.file(path)!.async('string')
    const colorKinds = ['colorSeries', 'colorNegative', 'colorAxis', 'colorMarkers', 'colorFirst', 'colorLast', 'colorHigh', 'colorLow']
    const repeatedColors = colorKinds.map(kind => Array.from({ length: 1_024 }, () => `<x14:${kind} rgb="FFFF0000"/>`).join('')).join('')
    const repeatedGroupDates = '<xm:f>Sheet1!A1:H1</xm:f>' + Array.from({ length: 1_024 }, () => '<xm:f>Sheet1!B1:B1</xm:f>').join('')
    xml = xml.replace('<x14:sparklines>', `${repeatedGroupDates}${repeatedColors}<x14:sparklines>`)
    const repeatedFormula = Array.from({ length: 1_024 }, () => '<xm:f>Sheet1!B1:B1</xm:f>').join('')
    const repeatedHost = Array.from({ length: 1_024 }, () => '<xm:sqref>K1</xm:sqref>').join('')
    xml = xml.replace('<xm:sqref>J1</xm:sqref>', `<xm:sqref>J1</xm:sqref>${repeatedFormula}${repeatedHost}`)
    xml = xml.replace('</x14:sparklines>', `</x14:sparklines>${'<x14:sparklines/>'.repeat(1_024)}`)
    zip.file(path, xml)
    const bytes = await zip.generateAsync({ type: 'uint8array' })
    const resident = await parseXlsx(await OfficePackage.load(bytes))
    const captured: XmlNode[] = []
    const parseGroups = sparklineGroupsParser.parseSparklineGroups
    const spy = vi.spyOn(sparklineGroupsParser, 'parseSparklineGroups').mockImplementation((node, color, diagnostics) => {
      captured.push(node)
      return parseGroups(node, color, diagnostics)
    })
    const session = await openImportSession({ kind: 'blob', blob: new Blob([new Uint8Array(bytes).buffer as ArrayBuffer]) }, { docId: 'spark-duplicate-singletons', profile: 'D' })
    try {
      await session.done
      const source = createPagedWorkbookSource(session), sheet = source.sheets[0]
      const store = source.store(sheet.sheetId) as unknown as { sheet: { metadata: { sparklineGroups: XlsxSparklineGroup[] } } }
      expect(store.sheet.metadata.sparklineGroups).toEqual(resident.sheets[0].sparklineGroups)

      const capturedRoot = captured.at(-1)!
      const firstChild = (value: unknown): XmlNode => Array.isArray(value) ? value[0] as XmlNode : value as XmlNode
      const group = firstChild(capturedRoot.sparklineGroup)
      const sparklineContainer = firstChild(group.sparklines)
      const sparkline = firstChild(sparklineContainer.sparkline)
      for (const kind of colorKinds) expect(group[kind]).not.toBeInstanceOf(Array)
      expect(group.f).not.toBeInstanceOf(Array)
      expect(group.sparklines).not.toBeInstanceOf(Array)
      expect(sparkline.f).not.toBeInstanceOf(Array)
      expect(sparkline.sqref).not.toBeInstanceOf(Array)
      await source.release()
    } finally {
      spy.mockRestore()
      await session.release()
    }
  })

  it('does not materialize a large data range for a sparkline with no host', async () => {
    const session = await openImportSession({ kind: 'blob', blob: new Blob([new Uint8Array(fixtureBytes('line-basic')).buffer as ArrayBuffer]) }, { docId: 'spark-invalid-host', profile: 'D' })
    try {
      await session.done
      const source = createPagedWorkbookSource(session), sheet = source.sheets[0], store = source.store(sheet.sheetId)
      type InternalStore = {
        sheet: { metadata: { sparklineGroups: XlsxSparklineGroup[] } }
        readSparklineCell(at: { row: number; col: number }, options: { signal: AbortSignal; stamp: typeof store.stamp }): Promise<unknown>
      }
      const internal = store as unknown as InternalStore
      const group = internal.sheet.metadata.sparklineGroups[0]
      expect(group).toBeDefined()
      const sparkline = group.sparklines[0]
      internal.sheet.metadata.sparklineGroups = [{ ...group, sparklines: [{ ...sparkline, host: undefined, data: { ...sparkline.data!, r0: 0, r1: 4_999, c0: 0, c1: 0 } }] }]
      let reads = 0, reserved = 0
      internal.readSparklineCell = async () => { reads++; return { kind: 'unknown' } }

      const result = await store.readSparklines!({ signal: new AbortController().signal, stamp: store.stamp }, 64 * 1024 * 1024, { set: bytes => { reserved += bytes; return true } })

      expect(result.sparklines).toEqual([])
      expect(reads).toBe(0)
      expect(reserved).toBeGreaterThan(0)
      await source.release()
    } finally { await session.release() }
  })

  it('skips zero-area host boxes but reserves null for oversized merges', () => {
    expect(sparklineInvalidationBox(10, 10, [], 1)).toBeUndefined()
    expect(sparklineInvalidationBox(10, 20, [], 1)).toBeUndefined()
    expect(sparklineInvalidationBox(10, 20, [{ top: 5, bottom: 5 }], 1)).toBeUndefined()
    expect(sparklineInvalidationBox(10, 20, [{ top: 5, bottom: 15 }], 10_001)).toBeNull()
    expect(sparklineInvalidationBox(10, 20, [{ top: 5, bottom: 15 }, { top: 20, bottom: 30 }], 2)).toEqual({ x: 10, y: 5, width: 10, height: 25 })
  })

  it('uses the RemoteWorkbook ledger for grant adoption, abort settlement, and early refund', async () => {
    const listeners = new Set<(event: MessageEvent) => void>()
    let deliver: ((reply: WorkerReply<'read'>) => void) | undefined
    const port = {
      postMessage(command: { kind: string; requestId: string; owner: unknown }, _transfer: Transferable[]) {
        if (command.kind !== 'read') return
        deliver = reply => { for (const listener of listeners) listener({ data: reply } as MessageEvent) }
      },
      addEventListener(_type: string, listener: (event: MessageEvent) => void) { listeners.add(listener) },
      removeEventListener(_type: string, listener: (event: MessageEvent) => void) { listeners.delete(listener) },
    }
    const client = new ViewerClient(port as never, 'spark-ledger')
    const remote = new RemoteWorkbook(client, { docId: 'spark-ledger', valuePolicy: 'cached', sheets: [{ sheetId: 's', name: 'S' }], diagnostics: [], layouts: new Map([['s', { rowOverrides: [], colOverrides: [], hidden: [], hasSparklines: true }]]) })
    const owner = remote.identity('s')
    const grant = (requestId: string) => remote.mainLedger.set('replies', `sparkline-grant:${requestId}`, 4096)
    const command = (requestId: string) => ({ kind: 'read' as const, requestId, owner, query: { op: 'sparklines' as const, grantBytes: 4096 } })
    const abortedId = 'aborted-before-send', abortBeforeSend = new AbortController()
    abortBeforeSend.abort(); grant(abortedId)
    expect(await remote.request(command(abortedId), abortBeforeSend.signal)).toBeTypeOf('symbol')
    expect(remote.mainLedger.ownerBytes('replies', `sparkline-grant:${abortedId}`)).toBe(0)

    const adoptedId = 'adopt-and-settle'; grant(adoptedId)
    const adoptedPromise = remote.request(command(adoptedId))
    await microtask()
    const adopted = { kind: 'complete', commandKind: 'read', requestId: adoptedId, owner, diagnostics: [], replyBytes: 512,
      result: { kind: 'read', result: { op: 'sparklines', result: { sparklines: [], groups: [], diagnostics: [] } } } } as unknown as WorkerReply<'read'>
    deliver!(adopted)
    const received = await adoptedPromise
    expect(remote.mainLedger.snapshot().byPool.sparklines).toBeGreaterThan(0)
    remote.mainLedger.settle(received as WorkerReply)
    expect(remote.mainLedger.snapshot().byPool.sparklines).toBe(0)

    const cancelledId = 'cancel-after-send'; grant(cancelledId)
    const abortAfterSend = new AbortController(), cancelledPromise = remote.request(command(cancelledId), abortAfterSend.signal)
    await microtask(); abortAfterSend.abort()
    expect(await cancelledPromise).toBeTypeOf('symbol')
    const late = { kind: 'complete', commandKind: 'read', requestId: cancelledId, owner, diagnostics: [], replyBytes: 512,
      result: { kind: 'read', result: { op: 'sparklines', result: { sparklines: [], groups: [], diagnostics: [] } } } } as unknown as WorkerReply<'read'>
    deliver!(late)
    await microtask()
    await microtask()
    expect(remote.mainLedger.snapshot().byPool.sparklines).toBe(0)
    expect(remote.mainLedger.ownerBytes('replies', `sparkline-grant:${cancelledId}`)).toBe(0)
    await remote.close()
  })

  it('returns the real RemoteWorkbook sparkline pool to zero on stale, cancel, quota, eviction, invalidation, and close', async () => {
    const listeners = new Set<(event: MessageEvent) => void>(), held: Array<() => void> = []
    let blockNext = false, quotaNext = false, hostRow = 0
    const group = { type: 'line', lineWeightPt: 0.75, displayEmptyCellsAs: 'gap', markers: false, high: false, low: false, first: false, last: false, negative: false, displayXAxis: false, displayHidden: false, dateAxis: false, rightToLeft: false, minAxisType: 'individual', maxAxisType: 'individual', colors: {}, sparklines: [] }
    const data: SheetSparklines = { sparklines: [{ host: { row: 0, col: 9 }, dataSheet: 'S', sourceRows: [0], groupIndex: 0, values: new Float64Array([7]), bridged: [], firstIndex: 0, lastIndex: 0, highIndices: [], lowIndices: [], min: 7, max: 7 }], groups: [group as SheetSparklines['groups'][number]], diagnostics: [] }
    let port!: { postMessage(command: WorkerCommand): void; addEventListener(_type: string, listener: (event: MessageEvent) => void): void; removeEventListener(_type: string, listener: (event: MessageEvent) => void): void }
    port = {
      postMessage(command) {
        const respond = () => {
          let out: unknown
          if (command.kind === 'close') out = { kind: 'complete', commandKind: 'close', requestId: command.requestId, owner: { docId: command.docId }, diagnostics: [], result: { kind: 'close' } }
          else if (command.kind === 'read' && command.query.op === 'sparklines') out = quotaNext
            ? { kind: 'quota', commandKind: 'read', requestId: command.requestId, owner: command.owner, resource: 'memory', requiredBytes: 8192, limitBytes: 4096, diagnostics: [] }
            : { kind: 'complete', commandKind: 'read', requestId: command.requestId, owner: command.owner, replyBytes: 128, diagnostics: [], result: { kind: 'read', result: { op: 'sparklines', result: { ...data, sparklines: data.sparklines.map(spark => ({ ...spark, host: { ...spark.host, row: hostRow } })) } } } }
          else out = { kind: 'error', commandKind: command.kind, requestId: command.requestId, owner: 'owner' in command ? command.owner : { docId: 'remote-cache' }, code: 'unsupported', retryable: false, diagnostics: [] }
          quotaNext = false
          for (const listener of listeners) listener({ data: out } as MessageEvent)
        }
        if (blockNext) { blockNext = false; held.push(respond) }
        else setTimeout(respond, 0)
      },
      addEventListener(_type, listener) { listeners.add(listener) },
      removeEventListener(_type, listener) { listeners.delete(listener) },
    }
    const client = new ViewerClient(port as never, 'spark-cache')
    const open = { docId: 'remote-cache', valuePolicy: 'cached' as const, sheets: [{ sheetId: 'a', name: 'S' }, { sheetId: 'b', name: 'S2' }, { sheetId: 'c', name: 'S3' }], diagnostics: [], layouts: new Map([
      ['a', { rowOverrides: [], colOverrides: [], hidden: [], hasSparklines: true }], ['b', { rowOverrides: [], colOverrides: [], hidden: [], hasSparklines: true }], ['c', { rowOverrides: [], colOverrides: [], hidden: [], hasSparklines: true }],
    ]) }
    const remote = new RemoteWorkbook(client, open)
    const internals = remote as unknown as { completed: Set<string>; structures: Map<string, { gen: number; merges: never[]; state: 'done' }>; identities: Map<string, ReturnType<RemoteWorkbook['identity']>>; sparklineCache: RemoteSparklineCache }
    for (const id of ['a', 'b', 'c']) { internals.completed.add(id); internals.structures.set(id, { gen: 0, merges: [], state: 'done' }) }
    Object.defineProperty(internals.sparklineCache, 'budgetBytes', { value: 1000 })
    const retained = new Map<string, ReturnType<typeof sheetWithSparklines>>()
    const unsubscribeEvictions = remote.subscribeSparklineEvictions(sheetId => {
      const sheet = retained.get(sheetId)
      if (sheet) retained.set(sheetId, sheetWithoutResolvedSparklines(sheet))
    })
    const bytes = () => remote.mainLedger.snapshot().byPool.sparklines
    const waitReady = async (id: string) => { remote.sparklines(id); await new Promise(resolve => setTimeout(resolve, 5)); return remote.sparklines(id) }

    expect((await waitReady('a')).state).toBe('ready')
    expect(bytes()).toBeGreaterThan(0)
    const beforeOwner = remote.identity('a')
    hostRow = 1
    internals.identities.set('a', { ...beforeOwner, valueGen: beforeOwner.valueGen + 1 })
    expect((await waitReady('a')).state).toBe('ready')
    expect(remote.takeSparklineChanges('a')).toEqual(expect.arrayContaining([{ row: 0, col: 9 }, { row: 1, col: 9 }]))
    internals.sparklineCache.invalidate('a')
    expect(bytes()).toBe(0)

    quotaNext = true
    expect((await waitReady('b')).state).toBe('failed')
    expect(bytes()).toBe(0)

    blockNext = true
    remote.sparklines('a')
    await microtask()
    const old = remote.identity('a')
    internals.identities.set('a', { ...old, valueGen: old.valueGen + 1 })
    held.splice(0).forEach(respond => respond())
    await new Promise(resolve => setTimeout(resolve, 5))
    expect(bytes()).toBe(0)

    // A cancelled in-flight request may still reply; the request boundary
    // settles that late reply without retaining its main-ledger lease.
    blockNext = true
    remote.sparklines('a')
    await microtask()
    internals.sparklineCache.invalidate('a')
    held.splice(0).forEach(respond => respond())
    await new Promise(resolve => setTimeout(resolve, 5))
    expect(bytes()).toBe(0)

    internals.identities.set('a', old)
    expect((await waitReady('a')).state).toBe('ready')
    retained.set('a', sheetWithSparklines({ name: 'A', rows: [], cols: [], merges: [], mergeRanges: [] }, remote.sparklines('a')))
    expect(retained.get('a')?.resolvedSparklines).toBeDefined()
    const oneEntry = bytes()
    expect((await waitReady('c')).state).toBe('ready') // C evicts A under the 1,000-byte cache cap.
    expect(bytes()).toBe(oneEntry)
    expect(retained.get('a')?.resolvedSparklines).toBeUndefined()
    unsubscribeEvictions()
    await remote.close()
    expect(bytes()).toBe(0)
  })

  it('drops an evicted sheet entry map when its LRU charge is released', async () => {
    const group = { type: 'line' as const, lineWeightPt: 0.75, displayEmptyCellsAs: 'gap' as const, markers: false, high: false, low: false, first: false, last: false, negative: false, displayXAxis: false, displayHidden: false, dateAxis: false, rightToLeft: false, minAxisType: 'individual' as const, maxAxisType: 'individual' as const, colors: {}, sparklines: [] }
    const result: SheetSparklines = {
      sparklines: [{ host: { row: 0, col: 9 }, dataSheet: 'sheet', sourceRows: [0], groupIndex: 0, values: new Float64Array([7]), bridged: [], firstIndex: 0, lastIndex: 0, highIndices: [], lowIndices: [], min: 7, max: 7 }],
      groups: [group], diagnostics: [],
    }
    let retainedSheet: ReturnType<typeof sheetWithSparklines> | undefined
    const evictions: string[] = []
    const identity = (id: string) => initialIdentity('spark-lru', id)
    const cache = new RemoteSparklineCache({
      identity, loadState: () => ({ completeness: 'complete' }), hasSparklines: () => true, grant: () => 4096,
      releaseGrant: () => {}, admitDecoded: () => () => {}, release: () => {}, nextRequestId: (() => { let id = 0; return () => `${++id}` })(),
      changed: (id, _hosts, reason?: string) => {
        if (reason !== 'evicted') return
        evictions.push(id)
        if (id === 'a' && retainedSheet) retainedSheet = sheetWithoutResolvedSparklines(retainedSheet)
      },
      request: async command => ({ kind: 'complete', commandKind: 'read', requestId: command.requestId, owner: command.owner, replyBytes: 128, diagnostics: [], result: { kind: 'read', result: { op: 'sparklines', result } } } as WorkerReply<'read'>),
    }, 1200)
    const ready = async (id: string) => { const pending = cache.get(id); await new Promise(resolve => setTimeout(resolve, 0)); return pending }
    const a = await ready('a')
    retainedSheet = sheetWithSparklines({ name: 'A', rows: [], cols: [], merges: [], mergeRanges: [] }, a)
    expect(retainedSheet.resolvedSparklines).toBeDefined()
    await ready('b')

    expect(evictions).toContain('a')
    expect(retainedSheet.resolvedSparklines).toBeUndefined()
    expect(cache.retainedEntries).toBe(1)
    expect(cache.retainedBytes).toBe(904)
    cache.close()
  })

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

  it('notifies subscribers when a completed refusal has diagnostics but no hosts', async () => {
    const identity = initialIdentity('doc', 'refused')
    const changed: Array<{ sheetId: string; hosts: readonly { row: number; col: number }[] }> = []
    const message = 'Sparkline data exceeds the receiver memory grant'
    const host: SparklineHost = {
      identity: () => identity, loadState: () => ({ completeness: 'complete' }), hasSparklines: () => true,
      grant: () => 4096, releaseGrant: () => {}, admitDecoded: () => () => {}, release: () => {}, nextRequestId: () => 'refused',
      changed: (sheetId, hosts) => changed.push({ sheetId, hosts: [...hosts] }),
      request: async command => ({ kind: 'complete', commandKind: 'read', requestId: command.requestId, owner: identity, replyBytes: 128, diagnostics: [], result: { kind: 'read', result: { op: 'sparklines', result: {
        sparklines: [], groups: [], diagnostics: [{ kind: 'unsupported', feature: 'sparkline-resource-limit', message }],
      } } } } as WorkerReply<'read'>),
    }
    const cache = new RemoteSparklineCache(host)
    cache.get('refused')
    await microtask()

    expect(sparklineNotices(cache.get('refused'))).toEqual([message])
    expect(changed).toEqual([{ sheetId: 'refused', hosts: [] }])
    cache.close()
  })

  it('retries refused grants and makes non-retryable request errors terminal once', async () => {
    let grants = 0, requests = 0
    const identity = initialIdentity('doc', 'sheet')
    const host: SparklineHost = {
      identity: () => identity, loadState: () => ({ completeness: 'complete' }), hasSparklines: () => true,
      grant: () => ++grants === 1 ? null : 4096, releaseGrant: () => {}, admitDecoded: () => null, release: () => {},
      request: async () => { requests++; return { kind: 'error', commandKind: 'read', requestId: '1', owner: identity, code: 'unsupported', retryable: false, diagnostics: [] } as WorkerReply<'read'> },
      nextRequestId: () => `${grants}`, changed: () => {},
    }
    const cache = new RemoteSparklineCache(host)
    cache.get('sheet')
    await new Promise(resolve => setTimeout(resolve, 300))
    expect(grants).toBeGreaterThan(1)
    expect(requests).toBe(1)
    expect(cache.get('sheet')).toMatchObject({ state: 'failed', diagnostics: [{ feature: 'sparkline-read-failed' }] })
    cache.close()
  })

  it('clears stale resolved data and notifies host tiles when get observes a new value owner', async () => {
    const first = initialIdentity('doc', 'sheet')
    let owner = first, serial = 0
    const hosts: { row: number; col: number }[][] = []
    const result: SheetSparklines = { sparklines: [{ host: { row: 4, col: 3 }, dataSheet: 'sheet', sourceRows: [0], groupIndex: 0, values: new Float64Array([1]), bridged: [], firstIndex: 0, lastIndex: 0, highIndices: [], lowIndices: [], min: 1, max: 1 }], groups: [{ type: 'line', lineWeightPt: 0.75, displayEmptyCellsAs: 'gap', markers: false, high: false, low: false, first: false, last: false, negative: false, displayXAxis: false, displayHidden: false, dateAxis: false, rightToLeft: false, minAxisType: 'individual', maxAxisType: 'individual', colors: {}, sparklines: [] }], diagnostics: [] }
    const host: SparklineHost = {
      identity: () => owner, loadState: () => ({ completeness: 'complete' }), hasSparklines: () => true, grant: () => 4096,
      releaseGrant: () => {}, admitDecoded: () => () => {}, release: () => {}, nextRequestId: () => `${++serial}`,
      changed: (_id, changed) => hosts.push([...changed]),
      request: async command => ({ kind: 'complete', commandKind: 'read', requestId: command.requestId, owner: command.owner, replyBytes: 128, diagnostics: [], result: { kind: 'read', result: { op: 'sparklines', result } } } as WorkerReply<'read'>),
    }
    const cache = new RemoteSparklineCache(host)
    cache.get('sheet'); await microtask()
    expect(cache.get('sheet').resolved).toBeTruthy()
    hosts.length = 0
    owner = { ...first, valueGen: first.valueGen + 1 }
    expect(cache.get('sheet')).toMatchObject({ state: 'loading', resolved: undefined, gen: 2 })
    expect(hosts).toEqual([])
    await Promise.resolve()
    expect(hosts.some(list => list.length === 1 && list[0].row === 4 && list[0].col === 3)).toBe(true)
    cache.close()
  })

  it('waits through a referenced sheet import and retries the sparkline snapshot after it completes', async () => {
    const identity = initialIdentity('doc', 'host')
    let dataState: 'loading' | 'complete' = 'loading', requests = 0, ids = 0
    const requestTimes: number[] = []
    const group = { type: 'line' as const, lineWeightPt: 0.75, displayEmptyCellsAs: 'gap' as const, markers: false, high: false, low: false, first: false, last: false, negative: false, displayXAxis: false, displayHidden: false, dateAxis: false, rightToLeft: false, minAxisType: 'individual' as const, maxAxisType: 'individual' as const, colors: {}, sparklines: [] }
    const resolved: SheetSparklines = { sparklines: [{ host: { row: 0, col: 9 }, dataSheet: 'Data', sourceRows: [0], groupIndex: 0, values: new Float64Array([7]), bridged: [], firstIndex: 0, lastIndex: 0, highIndices: [], lowIndices: [], min: 7, max: 7 }], groups: [group], diagnostics: [] }
    const host: SparklineHost = {
      identity: () => identity, loadState: id => ({ completeness: id === 'data' ? dataState : 'complete' }), hasSparklines: () => true,
      grant: () => 4096, releaseGrant: () => {}, admitDecoded: () => () => {}, release: () => {}, nextRequestId: () => `${++ids}`, changed: () => {},
      request: async command => {
        requests++
        requestTimes.push(Date.now())
        const result = dataState === 'loading' ? { sparklines: [], groups: [], pendingSheetIds: ['data'], diagnostics: [{ kind: 'unsupported' as const, feature: 'sparkline-data-unavailable', message: 'waiting for data sheet' }] } : resolved
        return { kind: 'complete', commandKind: 'read', requestId: command.requestId, owner: identity, replyBytes: 128, diagnostics: [], result: { kind: 'read', result: { op: 'sparklines', result } } } as WorkerReply<'read'>
      },
    }
    const cache = new RemoteSparklineCache(host)
    expect(cache.get('host').state).toBe('loading')
    await microtask()
    expect(cache.get('host').state).toBe('waiting')
    expect(cache.get('host').diagnostics).toEqual([])
    // Stay in loading beyond the former 20 × 250 ms retry window: it must not
    // become a terminal unavailable result while this referenced sheet imports.
    await new Promise(resolve => setTimeout(resolve, 5200))
    expect(cache.get('host').state).toBe('waiting')
    expect(requests).toBeLessThanOrEqual(6)
    const pendingIntervals = requestTimes.slice(1).map((time, index) => time - requestTimes[index])
    expect(pendingIntervals.slice(0, 4).every((interval, index) => interval >= [200, 400, 800, 1600][index])).toBe(true)
    dataState = 'complete'
    await new Promise(resolve => setTimeout(resolve, 3200))
    expect(cache.get('host')).toMatchObject({ state: 'ready', resolved: expect.any(Map) })
    cache.close()
  }, 14_000)

  it('ends a referenced sheet that failed import with one bounded diagnostic', async () => {
    const identity = initialIdentity('doc', 'host')
    const host: SparklineHost = {
      identity: () => identity, loadState: id => ({ completeness: id === 'data' ? 'incomplete' : 'complete' }), hasSparklines: () => true,
      grant: () => 4096, releaseGrant: () => {}, admitDecoded: () => null, release: () => {}, nextRequestId: () => 'failed-import', changed: () => {},
      request: async command => ({ kind: 'complete', commandKind: 'read', requestId: command.requestId, owner: identity, diagnostics: [], result: { kind: 'read', result: { op: 'sparklines', result: {
        sparklines: [], groups: [], pendingSheetIds: ['data'], diagnostics: [
          { kind: 'unsupported', feature: 'other', message: 'other source warning' },
          { kind: 'unsupported', feature: 'sparkline-data-unavailable', message: 'source sheet failed' },
        ],
      } } } } as WorkerReply<'read'>),
    }
    const cache = new RemoteSparklineCache(host)
    cache.get('host'); await microtask()
    expect(cache.get('host')).toMatchObject({ state: 'failed', diagnostics: [{ feature: 'sparkline-data-unavailable', message: 'source sheet failed' }] })
    expect(cache.get('host').diagnostics).toHaveLength(1)
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
      expect(output, fixture).toEqual(Object.fromEntries(resident.sheets.map(sheet => [sheet.name, { resolved: sheet.resolvedSparklines, diagnostics: sheet.diagnostics ?? [] }])))
    }
  })

  it('resident and paged paths omit date-axis members with only errors or gaps while preserving valid siblings', async () => {
    const bytes = await datedEmptyMembersBytes()
    const resident = await parseXlsx(await OfficePackage.load(bytes), { valuePolicy: 'cached' })
    const { output } = await resolvePaged(bytes, 'spark-dated-empty-members')
    expect(output).toEqual(Object.fromEntries(resident.sheets.map(sheet => [sheet.name, { resolved: sheet.resolvedSparklines, diagnostics: sheet.diagnostics ?? [] }])))
    const sheet = resident.sheets[0]
    expect(sheet.resolvedSparklines?.get(0)?.get(9)).toBeUndefined()
    expect(sheet.resolvedSparklines?.get(1)?.get(9)).toBeUndefined()
    expect(sheet.resolvedSparklines?.get(2)?.get(9)).toBeDefined()
    expect(sheet.resolvedSparklines?.get(3)?.get(9)).toBeDefined()
  })

  it('preflights and reads only the paired date prefix when a paged group has many longer date ranges', async () => {
    const bytes = await manyLongDateRangeSparklinesBytes(500)
    const resident = await parseXlsx(await OfficePackage.load(bytes), { valuePolicy: 'cached' })
    expect(resident.sheets[0].sparklineGroups?.[0].dateAxis).toBe(true)
    expect(resident.sheets[0].sparklineGroups?.[0].dateRange).toMatchObject({ r0: 6, c0: 0, r1: 6, c1: 51 })
    expect(resident.sheets[0].resolvedSparklines?.size).toBeGreaterThan(0)
    const session = await openImportSession({ kind: 'blob', blob: new Blob([new Uint8Array(bytes).buffer as ArrayBuffer]) }, { docId: 'spark-long-dates', profile: 'D' })
    try {
      await session.done
      const source = createPagedWorkbookSource(session), sheet = source.sheets[0], store = source.store(sheet.sheetId)
      const readCells = new Set<string>()
      type ReadSparklineCell = (cell: { row: number; col: number }, options: { signal: AbortSignal; stamp: typeof store.stamp }) => Promise<unknown>
      const instrumentedStore = store as unknown as { readSparklineCell: ReadSparklineCell }
      const readSparklineCell = instrumentedStore.readSparklineCell.bind(store)
      instrumentedStore.readSparklineCell = async (cell, options) => {
        readCells.add(`${cell.row}:${cell.col}`)
        return readSparklineCell(cell, options)
      }
      let admitted = 0
      const limit = 2 * 1024 * 1024
      const result = await store.readSparklines!({ signal: new AbortController().signal, stamp: store.stamp }, limit, { set: bytes => { admitted = Math.max(admitted, bytes); return bytes <= limit } })
      expect(result.diagnostics.some(item => item.feature === 'sparkline-resource-limit')).toBe(false)
      expect(result.sparklines).toHaveLength(500)
      expect(result.sparklines[0].axis).toBeDefined()
      expect(readCells).toEqual(new Set(['0:0', '6:0']))
      expect(admitted).toBeGreaterThan(0)
      expect(admitted).toBeLessThanOrEqual(limit)
      expect(rebuildResolvedSparklines(result)).toEqual(resident.sheets[0].resolvedSparklines)
      await source.release()
    } finally { await session.release() }
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
