import { expect, it, vi } from 'vitest'
import { LinkedControlTransaction, type LinkedControlOperation, type LinkedControlTarget } from '../src/xlsx/pivots/slicers/transaction'
import { SlicerController, sameSlicerHiddenItems } from '../src/xlsx/pivots/slicers/controller'
import { timelineLayout } from '../src/xlsx/pivots/slicers/paint'
import type { SlicerModel, TimelineModel } from '../src/xlsx/pivots/slicers/model'

const operation: LinkedControlOperation = { kind: 'slicerClear', name: 'Product' }
function pendingTransaction() {
  let bytes = 0, peak = 0, finish!: (result: { kind: 'staged'; value: number; bytes: number }) => void
  const publish = vi.fn(), discard = vi.fn()
  const target: LinkedControlTarget<LinkedControlOperation, number> = {
    id: 'one', estimate: () => ({ candidateBytes: 80, replyBytes: 20 }),
    stage: () => new Promise(resolve => { finish = resolve }), current: () => true, publish, discard,
  }
  const tx = new LinkedControlTransaction([target], { reserve(n) { bytes += n; peak = Math.max(peak, bytes); return { release() { bytes -= n } } } })
  return { tx, publish, discard, get bytes() { return bytes }, get peak() { return peak }, finish: () => finish({ kind: 'staged', value: 1, bytes: 100 }) }
}
it.each(['cancel', 'close', 'abort'] as const)('retains the 100-byte construction grant until a signal-ignoring producer drains: %s', async exit => {
  const h = pendingTransaction(), controller = new AbortController()
  const pending = h.tx.stage(operation, { revision: 1, grantBytes: 100, signal: controller.signal })
  if (exit === 'cancel') h.tx.cancel(1)
  else if (exit === 'close') h.tx.release()
  else controller.abort()
  expect(h.bytes).toBe(100)
  h.finish()
  expect(await pending).toMatchObject({ kind: 'cancelled' })
  expect(h.bytes).toBe(0)
  expect(h.peak).toBe(100)
  expect(h.discard).toHaveBeenCalledTimes(1)
  expect(h.publish).not.toHaveBeenCalled()
})
it('cannot commit an incomplete target set while a producer is pending', async () => {
  const h = pendingTransaction()
  const pending = h.tx.stage(operation, { revision: 1, grantBytes: 100, signal: new AbortController().signal })
  expect(h.tx.commit(1).kind).toBe('stale')
  expect(h.bytes).toBe(100)
  h.finish()
  expect((await pending).kind).toBe('staged')
  expect(h.tx.commit(1).kind).toBe('committed')
  h.tx.release()
  expect(h.bytes).toBe(0)
})
it('refuses before visiting 10,000 item identities when candidate memory is unavailable', async () => {
  let visits = 0
  const items = Array.from({ length: 10_000 }, (_, id) => ({ get id() { visits++; return id }, label: String(id), selected: true, noData: false }))
  const model = { name: 'Product', items, allItemIds: Array.from({ length: 10_000 }, (_, i) => i) } as SlicerModel
  const apply = vi.fn()
  const controller = new SlicerController([model], { apply, announce() {}, admit: () => undefined })
  visits = 0
  expect(await controller.clearSlicer('Product', new AbortController().signal)).toBe(false)
  expect(visits).toBe(0)
  expect(apply).not.toHaveBeenCalled()
  controller.dispose()
})
it('constructs only visible periods for a 100-year daily timeline', () => {
  const model = { name: 'Date', level: 'days', range: { start: Date.UTC(2000, 0, 1), end: Date.UTC(2100, 0, 1) } } as TimelineModel
  const result = timelineLayout(model, { x: 0, y: 0, width: 600, height: 120 }, 1)
  expect(result.periods.length).toBeGreaterThan(0)
  expect(result.periods.length).toBeLessThan(100)
})

it('compares 10,000 hidden items with linear source reads', () => {
  let reads = 0
  const actual = new Proxy(Array.from({ length: 10_000 }, (_, i) => i), { get(target, key, receiver) {
    if (typeof key === 'string' && /^\d+$/.test(key)) reads++
    return Reflect.get(target, key, receiver)
  } })
  expect(sameSlicerHiddenItems(Array.from({ length: 10_000 }, (_, i) => 9_999 - i), actual)).toBe(true)
  expect(reads).toBeLessThanOrEqual(10_000)
  expect(sameSlicerHiddenItems([1, 2], [1, 3])).toBe(false)
})
it('bounds four-target staging and returns every byte over 20 close cycles', async () => {
  let used = 0, peak = 0, constructions = 0
  for (let cycle = 0; cycle < 20; cycle++) {
    const targets = Array.from({ length: 4 }, (_, id): LinkedControlTarget<LinkedControlOperation, number> => ({
      id: String(id), estimate: () => ({ candidateBytes: 80, replyBytes: 20 }),
      stage() { expect(used).toBe(400); constructions++; return { kind: 'staged', value: id, bytes: 100 } },
      current: () => true, publish() {}, discard() {},
    }))
    const tx = new LinkedControlTransaction(targets, { reserve(bytes) { used += bytes; peak = Math.max(peak, used); return { release() { used -= bytes } } } })
    expect((await tx.stage(operation, { revision: 1, grantBytes: 400, signal: new AbortController().signal })).kind).toBe('staged')
    if (cycle % 2) tx.commit(1); else tx.cancel(1)
    tx.release()
    expect(used).toBe(0)
  }
  expect(peak).toBe(400)
  expect(constructions).toBe(80)
})
it('refuses an oversized worker composite before copying any rectangle elements', async () => {
  const { combinePivotOverlays } = await import('../src/worker/viewerSession')
  let reads = 0
  const rects = new Proxy(Array.from({ length: 10_000 }, (_, i) => ({ r0: i, r1: i, c0: 0, c1: 0 })), { get(target, key, receiver) {
    if (typeof key === 'string' && /^\d+$/.test(key)) reads++
    return Reflect.get(target, key, receiver)
  } })
  const overlay = { rects, protectedRects: rects, merges: [], cell: () => null }
  expect(() => combinePivotOverlays([overlay], 512)).toThrow('memory budget')
  expect(reads).toBe(0)
})
