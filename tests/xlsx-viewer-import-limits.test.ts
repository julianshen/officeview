/** S1.1b import limits: the quota ledger (design note §4.3; ledger parts of
 * RED tests 4 and 7) and the admission gates G1/G2 (RED test 19). */
import { deflateRawSync } from 'node:zlib'
import { describe, expect, it } from 'vitest'
import { bytesRangeSource, openPart, readZipInventory, type RangeSource, type ZipLimits } from '../src/core/zip-stream'
import { PackageError } from '../src/core/package-error'
import { createQuotaLedger, POOLS, quotaLimitsFor, type QuotaLimits } from '../src/xlsx/storage/quota'
import { buildZip, type BuildEntry } from './helpers/zip-builder'

const MiB = 1024 * 1024
const LIMITS: ZipLimits = {
  maxEntries: 4096, maxCentralDirectoryBytes: 4 * MiB, maxNameBytes: 1024,
  maxPartExpandedBytes: 2048 * MiB, maxTotalExpandedBytes: 4096 * MiB, maxRatio: 2000,
}
const LEDGER_LIMITS: QuotaLimits = { ...quotaLimitsFor('D'), expandedQueue: 8 * MiB }

function allZero(current: Record<string, number>): boolean {
  return POOLS.every((pool) => current[pool] === 0)
}

async function single(entry: BuildEntry): Promise<{ src: RangeSource & { reads: number }; entry: Awaited<ReturnType<typeof readZipInventory>>['entries'][number] }> {
  const bytes = buildZip([entry]).bytes
  const inner = bytesRangeSource(bytes)
  const src = { size: inner.size, reads: 0, read(a: number, b: number, s?: AbortSignal) { src.reads++; return inner.read(a, b, s) } }
  const inv = await readZipInventory(src, LIMITS)
  return { src, entry: inv.entries[0] }
}

describe('quota ledger', () => {
  it('tracks current and peak per pool, releases idempotently and types the overflow', () => {
    const ledger = createQuotaLedger({ ...LEDGER_LIMITS, decodedBuffers: 100, pages: 50 }, { maxTotalExpandedBytes: 1000 })
    const a = ledger.reserve('decodedBuffers', 60)
    const b = ledger.reserve('decodedBuffers', 40)
    expect(ledger.snapshot().current.decodedBuffers).toBe(100)
    let err: unknown
    try { ledger.reserve('decodedBuffers', 1) } catch (e) { err = e }
    expect((err as PackageError).code).toBe('resource-limit')
    try { ledger.reserve('pages', 51) } catch (e) { err = e }
    expect((err as PackageError).code).toBe('quota')
    a.release()
    a.release()
    expect(ledger.snapshot().current.decodedBuffers).toBe(40)
    expect(ledger.snapshot().peak.decodedBuffers).toBe(100)
    b.release()
    expect(allZero(ledger.snapshot().current)).toBe(true)
  })

  it('charges expanded bytes against the total and holds only the latest step in the queue', () => {
    const ledger = createQuotaLedger(LEDGER_LIMITS, { maxTotalExpandedBytes: 200_000 })
    ledger.charge(65_536)
    ledger.charge(65_536)
    expect(ledger.snapshot().current.expandedQueue).toBe(65_536)
    expect(ledger.snapshot().totalExpanded).toBe(131_072)
    const part = ledger.partBudget('xl/a.xml', 10)
    part.charge(10)
    expect(() => part.charge(1)).toThrow(PackageError)
    expect(() => ledger.charge(200_000)).toThrow(/expanded total/)
    ledger.drainQueue()
    expect(ledger.snapshot().current.expandedQueue).toBe(0)
  })

  it('releaseAll zeroes every pool and fails later charges with cancelled (AC-STORE-02)', () => {
    const ledger = createQuotaLedger(LEDGER_LIMITS, { maxTotalExpandedBytes: 1e9 })
    ledger.reserve('pages', 1000)
    ledger.reserve('sharedStrings', 10)
    ledger.charge(5)
    ledger.releaseAll()
    const snap = ledger.snapshot()
    expect(allZero(snap.current)).toBe(true)
    expect(snap.released).toBe(true)
    let err: unknown
    try { ledger.charge(1) } catch (e) { err = e }
    expect((err as PackageError).code).toBe('cancelled')
  })
})

describe('ledger through openPart (ledger parts of RED tests 4 and 7)', () => {
  it('a 64 MiB zero bomb under the 8 MiB ledger stops at limit + one output buffer; the ledger returns to zero', async () => {
    const size = 64 * MiB
    const raw = new Uint8Array(deflateRawSync(new Uint8Array(size)))
    const { src, entry } = await single({ name: 'bomb.xml', raw, uncompressedSize: size, crc: 0 })
    const ledger = createQuotaLedger(LEDGER_LIMITS, { maxTotalExpandedBytes: 8 * MiB })
    let yielded = 0
    let err: unknown
    try {
      for await (const chunk of openPart(src, entry, { signal: new AbortController().signal, budget: ledger })) yielded += chunk.length
    } catch (e) { err = e }
    expect((err as PackageError).code).toBe('resource-limit')
    expect(yielded).toBeLessThanOrEqual(8 * MiB + 65_536)
    expect(ledger.snapshot().peak.expandedQueue).toBeLessThanOrEqual(65_536)
    expect(ledger.snapshot().totalExpanded).toBeLessThanOrEqual(8 * MiB)
    ledger.drainQueue()
    expect(allZero(ledger.snapshot().current)).toBe(true)
  })

  it('cancellation mid-part surfaces cancelled, stops reads and releaseAll leaves every pool at zero', async () => {
    const data = new Uint8Array(3 * MiB)
    for (let i = 0, x = 7; i < data.length; i++) { x = (x * 1103515245 + 12345) >>> 0; data[i] = x >>> 24 }
    const { src, entry } = await single({ name: 'p.xml', data })
    const ledger = createQuotaLedger(LEDGER_LIMITS, { maxTotalExpandedBytes: 1e9 })
    const held = ledger.reserve('pages', 4 * MiB)
    void held
    const ac = new AbortController()
    let pulls = 0
    let readsAtAbort = 0
    let err: unknown
    try {
      for await (const chunk of openPart(src, entry, { signal: ac.signal, budget: ledger, rangeReadBytes: 65_536 })) {
        void chunk
        if (++pulls === 3) { ac.abort(); readsAtAbort = src.reads }
      }
    } catch (e) { err = e }
    expect((err as PackageError).code).toBe('cancelled')
    expect(src.reads).toBe(readsAtAbort)
    expect(ledger.snapshot().current.pages).toBe(4 * MiB)
    ledger.releaseAll()
    expect(allZero(ledger.snapshot().current)).toBe(true)
  })
})
