/**
 * C3 audit (bench-fix): control allocations whose charge could only end when a
 * never-collected or cache-owned object was collected. Collection callbacks are
 * simulated deterministically: every target an operation owns is "collected",
 * shared or cache-owned objects stay alive. Each repeated operation must then
 * leave the controls pool where it started.
 */
import { describe, expect, it } from 'vitest'
import { MainMemoryLedger } from '../src/xlsx/view/mainLedger'
import { MainControlMemory, controlSourceStore } from '../src/xlsx/view/mainControlMemory'
import { sortRowOrder } from '../src/xlsx/view/sort'
import { ObjectXlsxSheetAdapter } from '../src/xlsx/view/source'
import { createIdentityOrder } from '../src/xlsx/view/order'
import { valueSourceFromStore } from '../src/xlsx/view/valueIndex'
import type { XlsxSheet } from '../src/xlsx/types'

function tracked() {
  const registrations: { target: object; token: string }[] = []
  const ledger = new MainMemoryLedger({ register(target, token) { registrations.push({ target, token }) } })
  const memory = new MainControlMemory(ledger)
  /** Collects every registered target except those still alive. */
  const collectExcept = (alive: ReadonlySet<object>) => { for (const entry of registrations.splice(0)) if (!alive.has(entry.target)) ledger.collected(entry.token) }
  return { ledger, memory, collectExcept }
}

describe('C3 control-allocation tails', () => {
  it('a clear sort charges no bookkeeping onto the shared identity order', async () => {
    const { ledger, memory, collectExcept } = tracked()
    const store = new ObjectXlsxSheetAdapter({ docId: 'd', sheetId: 's', valueGen: 0 }, { name: 'S', sheetId: 's', rows: [], cols: [], merges: [] } as unknown as XlsxSheet, [], 100)
    const samples: number[] = []
    for (let i = 0; i < 3; i++) {
      const result = await sortRowOrder(store, null, { allocation: memory, now: () => 0, yield: async () => {} })
      expect(result.kind).toBe('complete')
      collectExcept(new Set([createIdentityOrder()]))
      samples.push(ledger.snapshot().byPool.controls)
    }
    expect(samples, 'controls after each collected clear sort').toEqual([0, 0, 0])
    expect(memory.pendingBytes).toBe(0)
  })

  it('structural reads do not re-charge the adapter-owned drawing anchors on every result', async () => {
    const { ledger, memory, collectExcept } = tracked()
    const sheet = { name: 'S', sheetId: 's', rows: [], cols: [], merges: [], mergeRanges: [], drawingAnchors: { records: [{ kind: 'absolute', id: 'd', x: 0, y: 0, cx: 9525, cy: 9525 }], unresolved: 0 } } as unknown as XlsxSheet
    const store = controlSourceStore(new ObjectXlsxSheetAdapter({ docId: 'd', sheetId: 's', valueGen: 0 }, sheet, [], 100), memory)
    const rect = { r0: 0, r1: 1048575, c0: 0, c1: 16383 }
    const samples: number[] = []
    let cache: Set<object> | undefined
    for (let i = 0; i < 3; i++) {
      const out = await store.queryStructural(rect, { signal: new AbortController().signal, stamp: store.stamp })
      expect(out.drawings).toHaveLength(1)
      // The adapter's resolved anchors (and every object they own) outlive each result.
      cache ??= new Set(reachableFrom(out.drawings[0]))
      collectExcept(cache)
      samples.push(ledger.snapshot().byPool.controls)
    }
    expect(samples[0], 'the resident anchor cache is accounted').toBeGreaterThan(0)
    expect(samples, 'controls after each collected structural result').toEqual([samples[0], samples[0], samples[0]])
  })

  it('a consumed and settled value page ends its decoded lanes and the source pages it came from', async () => {
    const { ledger, memory } = tracked()
    const rows = Array.from({ length: 4096 }, (_, row) => ({ index: row, cells: [{ row, col: 0, ref: '', value: row, styleIndex: 0 }] }))
    const sheet = { name: 'S', sheetId: 's', rows, cols: [], merges: [], mergeRanges: [] } as unknown as XlsxSheet
    const store = controlSourceStore(new ObjectXlsxSheetAdapter({ docId: 'd', sheetId: 's', valueGen: 0 }, sheet, [], 10000), memory)
    const source = valueSourceFromStore(store, 0, { r0: 0, r1: 4095 }, 4096, memory)
    const samples: number[] = []
    for (let i = 0; i < 3; i++) {
      const page = await source.readPage(0, new AbortController().signal)
      expect(page.kind).toBe('values')
      if (page.kind === 'values') page.settle?.()
      samples.push(ledger.snapshot().byPool.controls)
    }
    expect(samples, 'no collection ran: only settlement can end these charges').toEqual([0, 0, 0])
    expect(memory.pendingBytes).toBe(0)
  })

  it('an unsettled (still held) value page stays charged', async () => {
    const { ledger, memory } = tracked()
    const rows = Array.from({ length: 64 }, (_, row) => ({ index: row, cells: [{ row, col: 0, ref: '', value: row, styleIndex: 0 }] }))
    const sheet = { name: 'S', sheetId: 's', rows, cols: [], merges: [], mergeRanges: [] } as unknown as XlsxSheet
    const store = controlSourceStore(new ObjectXlsxSheetAdapter({ docId: 'd', sheetId: 's', valueGen: 0 }, sheet, [], 10000), memory)
    const page = await valueSourceFromStore(store, 0, { r0: 0, r1: 63 }, 64, memory).readPage(0, new AbortController().signal)
    expect(page.kind).toBe('values')
    expect(ledger.snapshot().byPool.controls).toBeGreaterThan(0)
  })
})

function reachableFrom(root: object): object[] {
  const out: object[] = [], pending: unknown[] = [root], seen = new Set<object>()
  while (pending.length) {
    const value = pending.pop()
    if (!value || typeof value !== 'object' || seen.has(value)) continue
    seen.add(value); out.push(value)
    if (ArrayBuffer.isView(value)) { pending.push(value.buffer); continue }
    for (const child of Object.values(value)) pending.push(child)
  }
  return out
}
