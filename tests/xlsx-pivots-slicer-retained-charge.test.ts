import { describe, expect, it, vi } from 'vitest'
import { SlicerController, SlicerRetainedCharges } from '../src/xlsx/pivots/slicers/controller'
import type { SlicerModel } from '../src/xlsx/pivots/slicers/model'
import { MainMemoryLedger } from '../src/xlsx/view/mainLedger'
import { MAIN_THREAD_BUDGETS } from '../src/xlsx/view/tileCache'

function model(name = 'Slicer_Product', revision = 0, count = 4): SlicerModel {
  const ids = Array.from({ length: count }, (_, i) => i)
  return {
    name, cacheName: name, cacheId: 1, sourceFieldIndex: 2, revision,
    allItemIds: ids, selectedItemIds: ids,
    caption: 'Product', showCaption: true, columnCount: 1, rowHeightEmu: 209550,
    style: 'SlicerStyleLight1', startItem: 0,
    items: ids.map(id => ({ id, label: `Item${id}`, selected: true, noData: false })),
    drives: [{ pivotTable: 'Pivot1' }],
  }
}

const retainedBytes = (m: SlicerModel): number =>
  256 + m.items.length * 128 + (m.allItemIds?.length ?? m.items.length) * 128

function ledgerWithNoGC() {
  return new MainMemoryLedger({ register() {} })
}

describe('slicer retained memory charge (option A)', () => {
  it('repeated cross-sheet filtering with disposal keeps accounted bytes at or above live retained bytes', async () => {
    const ledger = ledgerWithNoGC()
    // The host owns each transferred hold exactly like InteractiveSheetDoc
    // does: one SlicerRetainedCharges entry per slicer cache.
    const retained = new SlicerRetainedCharges()
    const kept = new Map<string, SlicerModel>()
    const liveRetainedBytes = () => [...kept.values()].reduce((sum, m) => sum + retainedBytes(m), 0)
    const accounted = () => ledger.snapshot().byPool.controls
    const makeHost = () => ({
      apply: async () => ({ kind: 'committed' as const, revision: 1 }),
      announce: vi.fn(),
      onChange: (m: SlicerModel, hold?: { release(): void }) => {
        kept.set(m.name, m)
        retained.retain(m.cacheName ?? m.name, hold)
      },
      admit: (bytes: number) => {
        if (ledger.bytes + bytes > MAIN_THREAD_BUDGETS.D) return undefined
        const owner = `slicer-selection:test:${Math.random()}`
        ledger.set('controls', owner, bytes)
        let held = true
        return { release: () => { if (!held) return; held = false; ledger.release('controls', owner) } }
      },
    })
    // Filter on sheet A, dispose (navigation), filter on sheet B, dispose, ...
    for (let sheet = 0; sheet < 3; sheet++) {
      const name = `Slicer_Sheet${sheet}`
      const controller = new SlicerController([model(name)], makeHost())
      expect(await controller.applySlicer(name, [0], new AbortController().signal)).toBe(true)
      controller.dispose()
      expect(accounted(), `after disposing sheet ${sheet}`).toBeGreaterThanOrEqual(liveRetainedBytes())
    }
    expect(liveRetainedBytes()).toBeGreaterThan(0)
    expect(accounted()).toBeGreaterThanOrEqual(liveRetainedBytes())
    expect(accounted()).toBe(liveRetainedBytes())
    retained.clear()
    expect(accounted()).toBe(0)
  })

  it('a partial sibling publish never refunds a hold the host still retains', async () => {
    const ledger = ledgerWithNoGC()
    const retained = new SlicerRetainedCharges()
    const accounted = () => ledger.snapshot().byPool.controls
    const admit = (bytes: number) => {
      if (ledger.bytes + bytes > MAIN_THREAD_BUDGETS.D) return undefined
      const owner = `slicer-selection:sibling:${Math.random()}`
      ledger.set('controls', owner, bytes)
      let held = true
      return { release: () => { if (!held) return; held = false; ledger.release('controls', owner) } }
    }
    // Two drawings share one slicer cache: one admitted hold covers both copies.
    const first = model('Product'), second = { ...model('Product 2'), cacheName: 'Product' }
    const announce = vi.fn()
    let calls = 0
    const controller = new SlicerController([first, second], {
      apply: async () => ({ kind: 'committed' as const, revision: 1 }),
      announce,
      admit,
      onChange: (m: SlicerModel, hold?: { release(): void }) => {
        calls++
        // The host retains the first sibling, then fails closed on the second.
        retained.retain(m.cacheName ?? m.name, hold)
        if (calls === 2) throw new Error('paint failed')
      },
    })
    const result = await controller.applySlicer('Product', [0], new AbortController().signal)
    expect(result).toBe(false)
    expect(calls).toBe(2)
    // The hold was transferred to the host on the first publish, so the failed
    // second publish must not refund it: the retained copy stays charged.
    expect(retained.size).toBe(1)
    expect(accounted()).toBeGreaterThan(0)
    const charged = accounted()
    retained.clear()
    expect(accounted()).toBe(0)
    expect(charged).toBeGreaterThan(0)
    controller.dispose()
  })

  it('replacing a model, evicting it, and unmounting the viewer each bring the slicer-model pool back to 0', async () => {
    const ledger = ledgerWithNoGC()
    const retained = new SlicerRetainedCharges()
    const pool = () => ledger.snapshot().byPool.controls
    const admit = (bytes: number) => {
      if (ledger.bytes + bytes > MAIN_THREAD_BUDGETS.D) return undefined
      const owner = `slicer-selection:retained:${Math.random()}`
      ledger.set('controls', owner, bytes)
      let held = true
      return { release: () => { if (!held) return; held = false; ledger.release('controls', owner) } }
    }
    const cacheKeyOf = (m: SlicerModel) => m.cacheName ?? m.name
    const retain = (m: SlicerModel) => {
      const bytes = retainedBytes(m)
      const hold = admit(bytes)
      if (!hold) throw new Error('admit failed')
      retained.retain(cacheKeyOf(m), hold)
    }
    // Replace: retaining a new revision releases the old charge; the pool holds only the new copy.
    retain(model('S', 0))
    const afterFirst = pool()
    expect(afterFirst).toBeGreaterThan(0)
    expect(retained.size).toBe(1)
    retain(model('S', 1))
    expect(pool()).toBe(afterFirst)
    expect(retained.size).toBe(1)
    // Evict: dropping the retained copy releases its charge (idempotent).
    retained.evict('S')
    expect(pool()).toBe(0)
    expect(retained.size).toBe(0)
    retained.evict('S')
    expect(pool()).toBe(0)
    // Unmount: clearing all retained copies releases everything (resident + paged share this path).
    retain(model('A', 0))
    retain(model('B', 0))
    expect(pool()).toBeGreaterThan(0)
    expect(retained.size).toBe(2)
    retained.clear()
    expect(pool()).toBe(0)
    expect(retained.size).toBe(0)
    retained.clear()
    expect(pool()).toBe(0)
  })

  it('a listener throwing on the first publish never refunds a hold the host already retained', async () => {
    const ledger = ledgerWithNoGC()
    const retained = new SlicerRetainedCharges()
    const accounted = () => ledger.snapshot().byPool.controls
    const admit = (bytes: number) => {
      if (ledger.bytes + bytes > MAIN_THREAD_BUDGETS.D) return undefined
      const owner = `slicer-selection:listener-throw:${Math.random()}`
      ledger.set('controls', owner, bytes)
      let held = true
      return { release: () => { if (!held) return; held = false; ledger.release('controls', owner) } }
    }
    const announce = vi.fn()
    const controller = new SlicerController([model('Product')], {
      apply: async () => ({ kind: 'committed' as const, revision: 1 }),
      announce,
      admit,
      onChange: (m: SlicerModel, hold?: { release(): void }) => {
        // The host retains the hold before any listener runs.
        retained.retain(m.cacheName ?? m.name, hold)
      },
    })
    // A subscriber throwing on the first publish must not trigger an early
    // refund: the hold was already transferred when onChange returned.
    controller.subscribe(() => { throw new Error('subscriber failed') })
    const result = await controller.applySlicer('Product', [0], new AbortController().signal)
    expect(result).toBe(false)
    expect(announce).toHaveBeenCalledWith('Slicer change failed.')
    // The retained copy stays charged.
    expect(retained.size).toBe(1)
    expect(accounted()).toBeGreaterThan(0)
    const charged = accounted()
    retained.clear()
    expect(accounted()).toBe(0)
    expect(charged).toBeGreaterThan(0)
    controller.dispose()
  })

  it('releasing twice does not refund twice', () => {
    const ledger = ledgerWithNoGC()
    const owner = 'slicer-selection:idempotent'
    ledger.set('controls', owner, 1024)
    expect(ledger.snapshot().byPool.controls).toBe(1024)
    let held = true
    const release = () => { if (!held) return; held = false; ledger.release('controls', owner) }
    release()
    expect(ledger.snapshot().byPool.controls).toBe(0)
    release()
    expect(ledger.snapshot().byPool.controls).toBe(0)
    expect(ledger.bytes).toBe(0)
  })

  it('refusal under pressure gives the notice, not a throw (resident path)', async () => {
    const ledger = ledgerWithNoGC()
    // Saturate the budget so the slicer admission cannot fit.
    ledger.set('copy', 'saturate', MAIN_THREAD_BUDGETS.D)
    const announce = vi.fn()
    const controller = new SlicerController([model()], {
      apply: vi.fn(async () => ({ kind: 'committed' as const, revision: 1 })),
      announce,
      admit: (bytes: number) => {
        if (ledger.bytes + bytes > MAIN_THREAD_BUDGETS.D) return undefined
        const owner = 'slicer-selection:refused'
        ledger.set('controls', owner, bytes)
        let held = true
        return { release: () => { if (!held) return; held = false; ledger.release('controls', owner) } }
      },
    })
    let threw: unknown
    let result = false
    try {
      result = await controller.applySlicer('Slicer_Product', [0], new AbortController().signal)
    } catch (error) { threw = error }
    expect(threw).toBeUndefined()
    expect(result).toBe(false)
    expect(announce).toHaveBeenCalledWith('The slicer selection exceeds the viewer memory budget.')
    controller.dispose()
  })

  it('refusal under pressure gives the notice, not a throw (worker/paged path)', async () => {
    const ledger = ledgerWithNoGC()
    ledger.set('copy', 'saturate', MAIN_THREAD_BUDGETS.D)
    const announce = vi.fn()
    // The paged path resolves its pivot target through the worker; the memory
    // admission is the same host admit, so pressure refuses before any worker call.
    const apply = vi.fn(async () => ({ kind: 'committed' as const, revision: 1 }))
    const controller = new SlicerController([model()], {
      apply, announce,
      admit: (bytes: number) => {
        if (ledger.bytes + bytes > MAIN_THREAD_BUDGETS.D) return undefined
        const owner = 'slicer-selection:paged-refused'
        ledger.set('controls', owner, bytes)
        let held = true
        return { release: () => { if (!held) return; held = false; ledger.release('controls', owner) } }
      },
    })
    let threw: unknown
    let result = false
    try {
      result = await controller.applySlicer('Slicer_Product', [0], new AbortController().signal)
    } catch (error) { threw = error }
    expect(threw).toBeUndefined()
    expect(result).toBe(false)
    expect(apply).not.toHaveBeenCalled()
    expect(announce).toHaveBeenCalledWith('The slicer selection exceeds the viewer memory budget.')
    controller.dispose()
  })
})
