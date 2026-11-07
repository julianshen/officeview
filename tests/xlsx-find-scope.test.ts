/**
 * Task 9 fix round 1, finding 7: Find must not lose ownership of a late or
 * rejected scope. A scope that settles after cancellation, or an acquired
 * scope rejected by the ownership guard, must still be disposed; byte charges
 * held for it must return to baseline.
 */
import { describe, expect, it } from 'vitest'
import { WorkbookFind, FindError } from '../src/xlsx/view/find'
import type { AuthoritativeCell } from '../src/xlsx/view/contracts'
import type { EffectiveFormatScope, FormatBudget } from '../src/xlsx/view/formatWorkspace'
import { drain, findInputOf, workbookOf } from './helpers/xlsx-find'

function trackingBudget(): { budget: FormatBudget; net: () => number; calls: string[] } {
  let held = 0
  const calls: string[] = []
  const budget: FormatBudget = {
    reserve: bytes => { held += bytes; calls.push(`reserve:${bytes}`) },
    release: bytes => { held -= bytes; calls.push(`release:${bytes}`) },
    retain: (_targets, bytes, commit) => { commit(); calls.push(`retain:${bytes}`) },
    settle: target => { calls.push('settle'); void target },
  }
  return { budget, net: () => held, calls }
}

interface TrackedScope extends EffectiveFormatScope {
  readonly bytes: number
  disposed: () => number
}

/** A scope that charges its budget on creation and refunds on dispose. */
function trackedScope(budget: FormatBudget, bytes = 100): TrackedScope {
  let disposed = 0
  budget.reserve(bytes)
  return {
    gen: 1,
    bytes,
    effectiveStyle: (_row, _col, base) => base,
    describe: () => undefined,
    dispose: () => { disposed++; budget.release(bytes) },
    disposed: () => disposed,
  }
}

const doc = () => workbookOf([{ sheetId: 'A', cells: [{ row: 0, col: 0, value: 'hit' }] }])

describe('late scope disposal on cancellation', () => {
  it('disposes a scope that settles after the search was closed', async () => {
    const { budget, net } = trackingBudget()
    let release!: (scope: TrackedScope) => void
    const gate = new Promise<TrackedScope>(resolve => { release = resolve })
    let started!: () => void
    const began = new Promise<void>(resolve => { started = resolve })
    const abort = new AbortController()
    const { input } = findInputOf(doc())
    const session = new WorkbookFind(input, { text: 'hit', caseSensitive: false }, {
      formatBudget: budget,
      textOf: cell => String(cell.value ?? ''),
      prepareScope: () => { started(); return gate },
      signal: abort.signal,
    })
    const pending = session.nextPage()
    // Let the scan reach the gated preparation, then close.
    await began
    abort.abort()
    await expect(pending).rejects.toMatchObject({ name: 'FindError', code: 'cancelled' })
    // The provider settles late: ownership must still end exactly once.
    const scope = trackedScope(budget)
    release(scope)
    await new Promise(resolve => setTimeout(resolve, 10))
    expect(scope.disposed()).toBe(1)
    expect(net()).toBe(0)
  })

  it('disposes an acquired scope rejected by an owner change, keeping stale classification', async () => {
    const { budget, net } = trackingBudget()
    const { input } = findInputOf(doc())
    let current = input.owner
    const seen: TrackedScope[] = []
    const session = new WorkbookFind(input, { text: 'hit', caseSensitive: false }, {
      formatBudget: budget,
      textOf: cell => String(cell.value ?? ''),
      prepareScope: () => {
        const scope = trackedScope(budget)
        seen.push(scope)
        current = { ...input.owner, globalViewGen: input.owner.globalViewGen + 1 }
        return Promise.resolve(scope)
      },
      currentOwner: () => current,
    })
    try {
      await expect(session.nextPage()).rejects.toMatchObject({ name: 'FindError', code: 'stale' })
      expect(seen).toHaveLength(1)
      expect(seen[0]!.disposed()).toBe(1)
      // The refused batch is admitted until the session closes.
      expect(net()).toBe(64)
    } finally {
      session.close()
    }
    expect(net()).toBe(0)
  })
})

describe('refused history accounting', () => {
  it('admits refused rects across pages and releases them on close', async () => {
    const { budget, net } = trackingBudget()
    const cells = Array.from({ length: 5000 }, (_, row) => ({ row, col: 0, value: row }))
    const { input } = findInputOf(workbookOf([{ sheetId: 'A', cells }]))
    const session = new WorkbookFind(input, { text: 'zzz-no-match', caseSensitive: false }, {
      formatBudget: budget,
      textOf: () => 'x',
      prepareScope: () => Promise.reject(new FindError('resource-limit', 'pressured')),
    })
    try {
      // Two scan pages refuse; nothing matches, so no materialization runs:
      // the refused index (two exact-cover rects) is admitted, not free.
      const hits = await drain(session)
      expect(hits).toEqual([])
      expect(net()).toBe(2 * 64)
      expect(session.stats.bufferedBytes).toBe(2 * 64)
    } finally {
      session.close()
    }
    expect(net()).toBe(0)
  })
})

describe('refused preparation path', () => {
  it('a refused preparation keeps base text without creating a scope', async () => {
    const { budget } = trackingBudget()
    let calls = 0
    const { input } = findInputOf(doc())
    const session = new WorkbookFind(input, { text: 'hit', caseSensitive: false }, {
      formatBudget: budget,
      textOf: (cell: AuthoritativeCell) => String(cell.value ?? ''),
      prepareScope: () => { calls++; return Promise.resolve(undefined) },
    })
    const hits = await drain(session)
    expect(hits.map(hit => hit.text)).toEqual(['hit'])
    expect(calls).toBeGreaterThan(0)
  })

  it('a materialize-time preparation failure refuses loudly instead of going stale', async () => {
    const { budget } = trackingBudget()
    const { input } = findInputOf(doc())
    let calls = 0
    const session = new WorkbookFind(input, { text: 'hit', caseSensitive: false }, {
      formatBudget: budget,
      textOf: (cell: AuthoritativeCell) => String(cell.value ?? ''),
      prepareScope: async () => {
        calls++
        // The scan batch prepares fine; the later hit re-read does not.
        if (calls > 1) throw new Error('transient pressure')
        return trackedScope(budget)
      },
    })
    await expect(drain(session)).rejects.toMatchObject({ name: 'FindError', code: 'resource-limit' })
    expect(calls).toBeGreaterThan(1)
  })
})
