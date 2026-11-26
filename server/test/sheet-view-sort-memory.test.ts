/**
 * Stage 4c Task 8 unblock — a sort job leaves nothing on the `sheet-view` port's main ledger when it ends, so
 * repeated sorts of a 1,048,576-row sheet never exhaust the port's controls budget, with no help from the garbage
 * collector. The committed view's order is charged by the server's P11 `DocumentLedger` hold, never twice. Before the
 * fix each sort left ~8.8 MiB behind (its order, plus the key pages it read), and the 64 MiB ledger refused the 6th
 * sort unless a GC had run.
 *
 * What a finished sort may still leave is the local engine's per-operation records, which local mode charges until
 * they are collected by design (`xlsx-viewer-evaluator-lifetime*.test.ts` pin it: an escaped matcher keeps them): the
 * session metadata (4,096 + 512 B) and the operation's structural index (1,696 B on this sheet).
 */
import { afterAll, describe, expect, test } from 'bun:test'
import type { WireRange } from '../../src/stream/sheet-wire'
import { DocumentHost } from '../src/render'
import type { SheetClock } from '../src/sheet/clock'
import type { SheetViewPort } from '../src/sheet/view-ops'
import { ONE_MILLION_ROWS, oneMillionSort } from './sheet-fixtures'

const hosts: DocumentHost[] = []
afterAll(() => { for (const h of hosts) for (const k of [...(h as unknown as { docs: Map<string, unknown> }).docs.keys()]) h.drop(k) })

const CLOCK: SheetClock = { day: '2026-11-27', timeZone: 'UTC' }
const IDENTITY = { sort: null, sortFilter: null, filter: null }
/** The per-operation records above, on `oneMillionSort`. */
const OPERATION_RECORD_BYTES = 4096 + 512 + 1696

/** The document's one `sheet-view` port (a test reaches into the host's entry). */
async function portOf(host: DocumentHost, docKey: string): Promise<SheetViewPort> {
  const doc = (host as unknown as { docs: Map<string, { viewPorts?: Map<string, { port: Promise<{ port: SheetViewPort }> }> }> }).docs.get(docKey)!
  const entries = [...doc.viewPorts!.values()]
  expect(entries.length).toBe(1)
  return (await entries[0]!.port).port
}

describe('sort jobs release their charges', () => {
  test('14 sequential million-row sorts, no forced GC', async () => {
    const host = new DocumentHost(); hosts.push(host)
    const { docKey } = await host.open('doc', await oneMillionSort(), 'full', { sheetMode: 'tiles' })
    const range: WireRange = { r0: 0, r1: ONE_MILLION_ROWS - 1, c0: 0, c1: 0, header: 0, totals: 0 }
    const first = await host.sheetView(docKey, CLOCK, 'full', { op: 'derive', sheetIndex: 0, view: IDENTITY })
    expect(first.op).toBe('view')
    const port = await portOf(host, docKey)
    const baseline = port.ledgerBytes()
    const seen: Array<[number, string, number]> = []
    for (let k = 0; k < 14; k++) {
      const order = k % 3 === 1 ? 'desc' : 'asc'
      const r = await host.sheetView(docKey, CLOCK, 'full', { op: 'sort', sheetIndex: 0, base: IDENTITY, sort: { range, keys: [{ col: 0, order }], caseSensitive: false } })
      const left = port.ledgerBytes() - baseline
      seen.push([k, r.op === 'view' && 'order' in r ? 'ok' : JSON.stringify(r), left <= (k + 1) * OPERATION_RECORD_BYTES ? 0 : left])
    }
    expect(seen).toEqual(seen.map(([k]) => [k, 'ok', 0]))
  }, 60_000)
})
