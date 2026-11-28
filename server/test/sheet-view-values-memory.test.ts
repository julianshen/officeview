/**
 * Stage 4c Task 7 unblock — a `values` job releases every main-ledger charge it took when it ends, so a whole list
 * fetched page by page (and its searches) never exhausts the port's controls budget, with no help from the garbage
 * collector. Before the fix each 10,000-entry job left ~9.6 MiB of reachability-leased page charges behind, and the
 * 64 MiB port ledger refused within ~21 jobs unless a GC had run.
 */
import { afterAll, describe, expect, test } from 'bun:test'
import { buildXlsx, type XlsxCellSpec } from '../../src/testdata/ooxml-builders'
import { DocumentHost } from '../src/render'
import type { SheetClock } from '../src/sheet/clock'

const hosts: DocumentHost[] = []
afterAll(() => { for (const h of hosts) for (const k of [...(h as unknown as { docs: Map<string, unknown> }).docs.keys()]) h.drop(k) })

const CLOCK: SheetClock = { day: '2026-11-27', timeZone: 'UTC' }
const IDENTITY = { sort: null, sortFilter: null, filter: null }

describe('values jobs release their charges', () => {
  test('60 sequential values jobs over a 10,000-entry list, no forced GC', async () => {
    const strings = ['H', ...Array.from({ length: 12_000 }, (_, i) => `W-${i}`)]
    const rows = strings.map((_, i) => ({ r: i + 1, cells: [{ ref: `A${i + 1}`, t: 's', v: i } as XlsxCellSpec] }))
    const host = new DocumentHost(); hosts.push(host)
    const { docKey } = await host.open('doc', await buildXlsx([{ name: 'S', rows }], strings), 'visible-only', { sheetMode: 'tiles' })
    const range = { r0: 0, r1: 12_000, c0: 0, c1: 0, header: 1 as const, totals: 0 as const }
    const failures: Array<[number, string]> = []
    for (let k = 0; k < 60; k++) {
      const search = k % 4 === 3 ? { search: { text: String(k % 10), wildcards: false } } : {}
      try {
        const r = await host.sheetView(docKey, CLOCK, 'visible-only', { op: 'values', sheetIndex: 0, view: IDENTITY, range, col: 0, offset: (k % 20) * 500, ...search })
        if (r.op !== 'values' || r.count !== 10_000) failures.push([k, JSON.stringify(r).slice(0, 80)])
      } catch (e) { failures.push([k, String(e)]) }
    }
    expect(failures).toEqual([])
  }, 30_000)
})
