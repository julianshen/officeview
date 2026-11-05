import { describe, expect, it, vi } from 'vitest'
import { LinkedControlTransaction, measuredDeselectFromAll, resolveSlicerPivotTargets, resolveWorkerControlLink, SharedControlIndexCache, validateLinkedControlOperation, type LinkedControlOperation, type LinkedControlTarget } from '../src/xlsx/pivots/slicers/transaction'

const operation: LinkedControlOperation = { kind: 'slicerSelection', name: 'Slicer_Product', selectedItemIds: [0] }

function harness(limit = 1_000) {
  let used = 0
  const holds: Array<{ bytes: number; released: boolean }> = []
  const memory = {
    reserve(bytes: number) {
      if (used + bytes > limit) return undefined
      used += bytes
      const hold = { bytes, released: false }
      holds.push(hold)
      return { release() { if (!hold.released) { hold.released = true; used -= bytes } } }
    },
    get used() { return used },
    holds,
  }
  return memory
}

type Candidate = { value: number; bytes: number }
function target(id: string, state: { value: number }, opts: { fail?: boolean; order?: string[] } = {}): LinkedControlTarget<LinkedControlOperation, Candidate> {
  return {
    id,
    estimate: () => ({ candidateBytes: 40, replyBytes: 10 }),
    stage: (_operation, context) => {
      opts.order?.push(`stage:${id}:${context.grantBytes}`)
      return opts.fail ? { kind: 'refused', message: `${id} refused` } : { kind: 'staged', value: { value: state.value + 1, bytes: 50 }, bytes: 50 }
    },
    current: candidate => candidate.bytes === 50,
    publish: candidate => { opts.order?.push(`publish:${id}`); state.value = candidate.value },
    discard: () => { opts.order?.push(`discard:${id}`) },
  }
}

describe('LinkedControlTransaction', () => {
  it('admits all candidate and reply bytes before constructing any target', async () => {
    const memory = harness(99), stage = vi.fn()
    const t: LinkedControlTarget<LinkedControlOperation, Candidate> = { ...target('a', { value: 0 }), estimate: () => ({ candidateBytes: 80, replyBytes: 20 }), stage }
    const tx = new LinkedControlTransaction([t], memory)
    const result = await tx.stage(operation, { revision: 7, grantBytes: 100, signal: new AbortController().signal })
    expect(result).toMatchObject({ kind: 'refused', feature: 'pivot-control-memory-limit', requiredBytes: 100 })
    expect(stage).not.toHaveBeenCalled()
    expect(memory.used).toBe(0)
  })

  it('stages every linked target before publishing and leaves all unchanged when one refuses', async () => {
    const states = [{ value: 0 }, { value: 0 }], order: string[] = []
    const tx = new LinkedControlTransaction([target('a', states[0]!, { order }), target('b', states[1]!, { fail: true, order })], harness())
    const result = await tx.stage(operation, { revision: 7, grantBytes: 100, signal: new AbortController().signal })
    expect(result).toMatchObject({ kind: 'refused', feature: 'pivot-control-target-refused' })
    expect(order).toEqual(['stage:a:50', 'stage:b:50', 'discard:a'])
    expect(states).toEqual([{ value: 0 }, { value: 0 }])
  })

  it('drops a stale commit and refunds the staged hold exactly once', async () => {
    const memory = harness(), state = { value: 0 }, discard = vi.fn()
    const t = { ...target('a', state), discard }
    const tx = new LinkedControlTransaction([t], memory)
    await tx.stage(operation, { revision: 7, grantBytes: 100, signal: new AbortController().signal })
    // A candidate that became obsolete at one target is dropped as a whole.
    t.current = () => false
    expect(tx.commit(7)).toMatchObject({ kind: 'stale' })
    expect(state.value).toBe(0)
    expect(discard).toHaveBeenCalledTimes(1)
    expect(memory.used).toBe(0)
    tx.cancel(7)
    expect(discard).toHaveBeenCalledTimes(1)
  })

  it('drops and refunds a candidate when commit carries the wrong revision', async () => {
    const memory = harness(), state = { value: 0 }, discard = vi.fn()
    const tx = new LinkedControlTransaction([{ ...target('a', state), discard }], memory)
    await tx.stage(operation, { revision: 7, grantBytes: 100, signal: new AbortController().signal })
    expect(tx.commit(8)).toMatchObject({ kind: 'stale' })
    expect(state.value).toBe(0)
    expect(discard).toHaveBeenCalledTimes(1)
    expect(memory.used).toBe(0)
  })

  it('commits the whole target set, replaces its hold, and releases once on close', async () => {
    const memory = harness(), states = [{ value: 0 }, { value: 0 }]
    const tx = new LinkedControlTransaction(states.map((state, i) => target(String(i), state)), memory)
    await tx.stage(operation, { revision: 1, grantBytes: 100, signal: new AbortController().signal })
    expect(tx.commit(1)).toEqual({ kind: 'committed', revision: 1, targetIds: ['0', '1'] })
    expect(states.map(s => s.value)).toEqual([1, 1])
    expect(memory.used).toBe(100)
    await tx.stage(operation, { revision: 2, grantBytes: 100, signal: new AbortController().signal })
    expect(tx.commit(2).kind).toBe('committed')
    expect(states.map(s => s.value)).toEqual([2, 2])
    expect(memory.used).toBe(100)
    tx.release()
    tx.release()
    expect(memory.used).toBe(0)
    expect(memory.holds.filter(h => h.released)).toHaveLength(2)
  })

  it('cancellation and release discard a pending stage once', async () => {
    const memory = harness(), state = { value: 0 }, discard = vi.fn()
    const tx = new LinkedControlTransaction([{ ...target('a', state), discard }], memory)
    await tx.stage(operation, { revision: 7, grantBytes: 100, signal: new AbortController().signal })
    expect(tx.cancel(7)).toBe(true)
    tx.release()
    expect(discard).toHaveBeenCalledTimes(1)
    expect(memory.used).toBe(0)
    expect(state.value).toBe(0)
  })
})

describe('SharedControlIndexCache', () => {
  it('builds and charges one source index per shared cache until the last control releases', () => {
    const memory = harness(), build = vi.fn(() => ({ rows: 100 }))
    const cache = new SharedControlIndexCache<string, { rows: number }>(memory)
    const slicer = cache.acquire('cache:7', 80, build)
    const timeline = cache.acquire('cache:7', 80, build)
    expect(slicer.kind).toBe('acquired')
    expect(timeline.kind).toBe('acquired')
    if (slicer.kind !== 'acquired' || timeline.kind !== 'acquired') throw new Error('expected leases')
    expect(slicer.value).toBe(timeline.value)
    expect(build).toHaveBeenCalledTimes(1)
    expect(memory.used).toBe(80)
    slicer.lease.release()
    expect(memory.used).toBe(80)
    timeline.lease.release()
    expect(memory.used).toBe(0)
  })

  it('refuses before building when the source index cannot be admitted', () => {
    const memory = harness(79), build = vi.fn(() => ({ rows: 100 }))
    const cache = new SharedControlIndexCache<string, { rows: number }>(memory)
    expect(cache.acquire('cache:7', 80, build)).toMatchObject({ kind: 'refused', requiredBytes: 80 })
    expect(build).not.toHaveBeenCalled()
  })

  it('shares one admitted 100,000-record index across eight linked pivots instead of rescanning per target', () => {
    const memory = harness(2_000_000), cache = new SharedControlIndexCache<string, { ids: Set<number> }>(memory)
    let visits = 0, builds = 0
    const build = () => {
      builds++
      const ids = new Set<number>()
      for (let row = 0; row < 100_000; row++) { visits++; ids.add(row) }
      return { ids }
    }
    const leases = Array.from({ length: 8 }, () => cache.acquire('pivot-cache:7:Product', 1_500_000, build))
    expect(leases.every(item => item.kind === 'acquired')).toBe(true)
    expect(builds).toBe(1)
    expect(visits).toBe(100_000)
    expect(memory.used).toBe(1_500_000)
    for (const lease of leases) if (lease.kind === 'acquired') lease.lease.release()
    expect(memory.used).toBe(0)
  })
})

describe('linked-control bounds and resolution', () => {
  it('represents measured clears explicitly and fails closed on unmeasured click intent', () => {
    expect(validateLinkedControlOperation({ kind: 'slicerClear', name: 'Slicer_Product' })).toBe(true)
    expect(validateLinkedControlOperation({ kind: 'timelineClear', name: 'NativeTimeline_Date' })).toBe(true)
    expect(validateLinkedControlOperation({ kind: 'slicerSelection', name: 'Slicer_Product', selectedItemIds: [1], modifier: true })).toBe(false)
    expect(validateLinkedControlOperation({ kind: 'timelineSelection', name: 'NativeTimeline_Date', start: 10, end: 20, action: 'drag' })).toBe(false)
    expect(measuredDeselectFromAll('Slicer_Product', [0, 1, 2, 3], [0, 1, 2, 3], 0)).toEqual({ kind: 'slicerSelection', name: 'Slicer_Product', selectedItemIds: [1, 2, 3] })
    // S5-C's measured fact is only the all-selected starting state. No helper normalizes later toggle states.
    expect(measuredDeselectFromAll('Slicer_Product', [0, 1, 2, 3], [1, 2, 3], 1)).toBeUndefined()
    expect(measuredDeselectFromAll('Slicer_Product', [0, 1, 2, 3], [0, 1, 2, 3], 4)).toBeUndefined()
  })

  it('resolves only the saved cache linkage and refuses missing, duplicate or over-limit targets', () => {
    const pivots = [
      { sheetId: 's1', pivot: 0, name: 'Sales', cacheId: 7 },
      { sheetId: 's2', pivot: 1, name: 'Returns', cacheId: 7 },
      { sheetId: 's3', pivot: 2, name: 'OtherCache', cacheId: 8 },
    ]
    expect(resolveSlicerPivotTargets({ name: 'Slicer_Product', kind: 'tabular', pivotTables: ['Returns'], pivotCacheId: 7 }, pivots))
      .toMatchObject({ ok: true, targets: [pivots[1]] })
    expect(resolveSlicerPivotTargets({ name: 'Slicer_Product', kind: 'tabular', pivotTables: ['Sales'], pivotCacheId: 8 }, pivots))
      .toMatchObject({ ok: false, feature: 'pivot-control-target-missing' })
    expect(resolveSlicerPivotTargets({ name: 'Slicer_Product', kind: 'tabular', pivotTables: ['Sales', 'Returns'], pivotCacheId: 7 }, [
      { ...pivots[0]!, layoutKey: 'Region×Product' }, { ...pivots[1]!, layoutKey: 'Date×Product' },
    ])).toMatchObject({ ok: false, feature: 'pivot-control-linked-layout-unverified' })
    const sameLayout = resolveSlicerPivotTargets({ name: 'Slicer_Product', kind: 'tabular', pivotTables: ['Sales', 'Returns'], pivotCacheId: 7 }, [
      { ...pivots[0]!, layoutKey: 'Region×Product' }, { ...pivots[1]!, layoutKey: 'Region×Product' },
    ])
    expect(sameLayout).toMatchObject({ ok: true })
    expect(resolveSlicerPivotTargets({ name: 'Slicer_Product', kind: 'tabular', pivotTables: ['Sales', 'Sales'], pivotCacheId: 7 }, pivots))
      .toMatchObject({ ok: false, feature: 'pivot-control-target-duplicate' })
    expect(resolveSlicerPivotTargets({ name: 'Slicer_Product', kind: 'tabular', pivotTables: Array(9).fill('Sales'), pivotCacheId: 7 }, pivots))
      .toMatchObject({ ok: false, feature: 'pivot-control-target-limit' })
  })

  it('refuses unknown controls and any target from a mixed pivot cache at the worker boundary', () => {
    const link = { name: 'Slicer_Product', kind: 'slicer' as const, cacheId: 7, pivotTables: ['Sales', 'Returns'] }
    expect(resolveWorkerControlLink({ kind: 'slicerSelection', name: 'Other', selectedItemIds: [0] }, [link], [
      { name: 'Sales', cacheId: 7 }, { name: 'Returns', cacheId: 7 },
    ])).toMatchObject({ ok: false, feature: 'pivot-control-linkage-unknown' })
    expect(resolveWorkerControlLink({ kind: 'slicerSelection', name: 'Slicer_Product', selectedItemIds: [0] }, [link], [
      { name: 'Sales', cacheId: 7 }, { name: 'Returns', cacheId: 8 },
    ])).toMatchObject({ ok: false, feature: 'pivot-control-linkage-cache-mismatch' })
  })

  it('rejects oversized, duplicate and malformed item selections before target estimates', async () => {
    const estimate = vi.fn(() => ({ candidateBytes: 1, replyBytes: 1 }))
    const t = { ...target('a', { value: 0 }), estimate }
    const tx = new LinkedControlTransaction([t], harness())
    const oversized = { kind: 'slicerSelection', name: 'Slicer_Product', selectedItemIds: Array.from({ length: 10_001 }, (_, i) => i) } as const
    expect(await tx.stage(oversized, { revision: 1, grantBytes: 1_000_000, signal: new AbortController().signal }))
      .toMatchObject({ kind: 'refused', feature: 'pivot-control-operation-invalid' })
    expect(await tx.stage({ ...operation, selectedItemIds: [0, 0] }, { revision: 2, grantBytes: 100, signal: new AbortController().signal }))
      .toMatchObject({ kind: 'refused', feature: 'pivot-control-operation-invalid' })
    expect(estimate).not.toHaveBeenCalled()
  })

  it('discards a late target result after cancellation and releases its admission once', async () => {
    const memory = harness(), state = { value: 0 }, discard = vi.fn()
    let finish!: (value: { kind: 'staged'; value: Candidate; bytes: number }) => void
    const t: LinkedControlTarget<LinkedControlOperation, Candidate> = {
      ...target('a', state),
      stage: () => new Promise(resolve => { finish = resolve }),
      discard,
    }
    const tx = new LinkedControlTransaction([t], memory)
    const pending = tx.stage(operation, { revision: 1, grantBytes: 100, signal: new AbortController().signal })
    await Promise.resolve()
    expect(memory.used).toBe(50)
    expect(tx.cancel(1)).toBe(true)
    expect(memory.used).toBe(50)
    finish({ kind: 'staged', value: { value: 1, bytes: 50 }, bytes: 50 })
    expect(await pending).toMatchObject({ kind: 'cancelled' })
    expect(memory.used).toBe(0)
    expect(discard).toHaveBeenCalledTimes(1)
    expect(memory.holds.filter(hold => hold.released)).toHaveLength(1)
  })

  it('keeps the caller signal live until commit and drops a staged set if it aborts', async () => {
    const memory = harness(), state = { value: 0 }, discard = vi.fn(), controller = new AbortController()
    const tx = new LinkedControlTransaction([{ ...target('a', state), discard }], memory)
    expect(await tx.stage(operation, { revision: 1, grantBytes: 100, signal: controller.signal })).toMatchObject({ kind: 'staged' })
    expect(memory.used).toBe(50)
    controller.abort()
    expect(tx.commit(1)).toMatchObject({ kind: 'stale' })
    expect(state.value).toBe(0)
    expect(discard).toHaveBeenCalledTimes(1)
    expect(memory.used).toBe(0)
  })
})
