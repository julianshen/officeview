// @vitest-environment node
// r14 invariant: over any sequence of retain / release / trim / close / abandon on
// decoded records that alias one RAM backing, the pool counts at least every
// distinct live backing once plus every live record's own heap once, never a
// backing twice, and nothing once the last lease is released.
import { expect, it } from 'vitest'
import { SourcePageStore, memoryPageBackend, PageBuilder, K_NUMBER } from '../src/xlsx/storage/sourcePages'
import { MemoryPool, type PoolCharge } from '../src/xlsx/view/tileCache'

interface Held { record: object; buffer: ArrayBufferLike; own: number; caller?: PoolCharge; release(): void }

function rng(seed: number): () => number {
  let s = seed >>> 0
  return () => { s = (s * 1664525 + 1013904223) >>> 0; return s / 2 ** 32 }
}

for (let seed = 1; seed <= 64; seed++) it(`aliasing leases keep the pool >= distinct live backings + record heaps, seed ${seed}`, async () => {
  const random = rng(seed)
  const store = new SourcePageStore({ docId: `alias-invariant-${seed}`, backend: memoryPageBackend() })
  await store.commitStrings(0, ['漢'.repeat(20000)], 1)
  const b = new PageBuilder()
  for (let c = 0; c < 2000; c++) b.add(0, c, K_NUMBER, 0, c, '', false, undefined)
  await store.commit({ sheetId: '1', pageId: 'p0', r0: 0, r1: 0, cellCount: 2000, bytes: b.bytes, payload: b.build() })
  const pool = new MemoryPool({ limitBytes: 1_000_000 })
  store.bindMemory(pool)
  const held: Held[] = []
  let closed = false

  const pin = async (string: boolean, unretained: boolean): Promise<void> => {
    const pending = string ? store.pinnedString('sst', 0) : store.pinnedPage('1', 'p0')
    if (unretained) pool.trim(0)
    const lease = (await pending)!
    const record = 'record' in lease ? lease.record : lease.page
    const buffer = ('record' in lease ? lease.record.text : lease.page.rows).buffer
    const entry: Held = { record, buffer, own: lease.bytes, release: lease.release }
    if (!lease.retained) { entry.caller = pool.charge(); expect(entry.caller.set(lease.bytes)).toBe(true) }
    held.push(entry)
  }
  const drop = (i: number): void => {
    const [entry] = held.splice(i, 1)
    entry.release()
    entry.caller?.release()
  }
  const check = (): void => {
    const buffers = new Map<ArrayBufferLike, number>(), records = new Map<object, number>()
    let callers = 0
    for (const h of held) { buffers.set(h.buffer, h.buffer.byteLength); records.set(h.record, h.own); callers += h.caller?.bytes ?? 0 }
    const backings = [...buffers.values()].reduce((a, n) => a + n, 0)
    const own = [...records.values()].reduce((a, n) => a + n, 0)
    const total = pool.stats().totalBytes
    expect(total).toBeGreaterThanOrEqual(backings + own)
    // Exactly: each backing once, the heap of each record held by retained leases
    // once (its kept entry), and what unretained holders charged themselves.
    if (closed) {
      const kept = [...records].filter(([record]) => !held.some(h => h.record === record && h.caller)).reduce((a, [, n]) => a + n, 0)
      expect(total).toBe(backings + kept + callers)
    }
  }

  try {
    for (let step = 0; step < 14; step++) {
      const r = random()
      if (!closed && step >= 6 && r < 0.2) {
        if (random() < 0.5) await store.close(); else await store.abandon(store.docId)
        closed = true
      } else if (!closed && r < 0.6) await pin(random() < 0.5, random() < 0.5)
      else if (r < 0.75) pool.trim(0)
      else if (held.length) drop(Math.floor(random() * held.length))
      check()
    }
    if (!closed) { await store.close(); closed = true; check() }
    while (held.length) { drop(Math.floor(random() * held.length)); check() }
    expect(pool.stats().totalBytes).toBe(0)
  } finally {
    for (const h of held) { h.release(); h.caller?.release() }
    await store.close()
    pool.close()
  }
})
