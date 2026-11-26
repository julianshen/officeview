/** Stage 4c Task 4 (plan P2, P11): the canonical view, its hash, the ordered-sheet cap and the view holds. */
import { describe, expect, test } from 'bun:test'
import type { WireRange, WireSortSpec } from '../../src/stream/sheet-wire'
import { viewStateHash } from '../src/cache'
import { canonicalView, canonicalViewHash, IDENTITY_VIEW, isIdentityView, orderBytes, viewTier, type CanonicalView, type ResolvedFilter } from '../src/sheet/view-spec'
import { MAX_ORDERED_SHEETS, SessionViews, ViewStateLedger } from '../src/sheet/view-state'

const clock = { day: '2026-11-23', timeZone: 'UTC' }
const range: WireRange = { r0: 0, r1: 99, c0: 0, c1: 3, header: 1, totals: 0 }
const S: WireSortSpec = { range, keys: [{ col: 2, order: 'desc' }, { col: 0, order: 'asc' }], caseSensitive: false }
const F1: ResolvedFilter = {
  range,
  columns: [
    [3, { kind: 'values', list: { range, col: 3 }, positions: [4, 1, 4, 2], mode: 'include', blanks: false }],
    [1, { kind: 'text', op: 'contains', first: 'secret', wildcards: false }],
  ],
}
const hashOf = (view: CanonicalView) => viewStateHash({ view, sizes: [], authorHiddenTier: 'none', clock })

describe('the canonical view (P2)', () => {
  test('sortFilter keys the hash', () => {
    const a = hashOf({ sort: S, sortFilter: F1, filter: F1 })
    const b = hashOf({ sort: S, sortFilter: null, filter: F1 })
    expect(a).not.toBe(b)
    // Column order and values positions in the input do not change the hash; sort key order does (it is priority).
    const reordered: ResolvedFilter = { range, columns: [F1.columns[1]!, [3, { kind: 'values', list: { range, col: 3 }, positions: [2, 1, 4], mode: 'include', blanks: false }]] }
    expect(hashOf({ sort: S, sortFilter: reordered, filter: reordered })).toBe(a)
    expect(hashOf({ sort: { ...S, keys: [...S.keys].reverse() }, sortFilter: F1, filter: F1 })).not.toBe(a)
    // sortFilter is dropped when sort is null.
    expect(hashOf({ sort: null, sortFilter: F1, filter: F1 })).toBe(hashOf({ sort: null, sortFilter: null, filter: F1 }))
    expect(canonicalView({ sort: null, sortFilter: F1, filter: null })).toEqual(IDENTITY_VIEW)
  })

  test('canonicalView sorts filter columns and values positions and drops absent operands', () => {
    const c = canonicalView({ sort: S, sortFilter: F1, filter: F1 })
    expect(c.filter!.columns.map(([col]) => col)).toEqual([1, 3])
    expect(c.filter!.columns[1]![1]).toEqual({ kind: 'values', list: { range, col: 3 }, positions: [1, 2, 4], mode: 'include', blanks: false })
    expect(c.sort!.keys).toEqual(S.keys)
    const n = canonicalView({ sort: null, sortFilter: null, filter: { range, columns: [[0, { kind: 'number', op: 'gt', first: 3, second: undefined }]] } })
    expect(Object.keys(n.filter!.columns[0]![1])).toEqual(['kind', 'op', 'first'])
    expect(canonicalViewHash(n)).toBe(canonicalViewHash({ sort: null, sortFilter: null, filter: { range, columns: [[0, { kind: 'number', op: 'gt', first: 3 }]] } }))
    expect(canonicalViewHash(IDENTITY_VIEW)).toMatch(/^[0-9a-f]{64}$/)
    expect(canonicalViewHash(c)).not.toBe(canonicalViewHash(IDENTITY_VIEW))
  })

  test('identity, the C1 tier and orderBytes (P11)', () => {
    expect(isIdentityView(IDENTITY_VIEW)).toBe(true)
    expect(isIdentityView({ sort: null, sortFilter: F1, filter: null })).toBe(true)
    expect(isIdentityView({ sort: S, sortFilter: null, filter: null })).toBe(false)
    expect(isIdentityView({ sort: null, sortFilter: null, filter: F1 })).toBe(false)
    expect(viewTier('visible-only', { sort: null, sortFilter: null, filter: F1 })).toBe('c1')
    expect(viewTier('visible-only', { sort: S, sortFilter: F1, filter: null })).toBe('none')
    expect(viewTier('full', { sort: null, sortFilter: null, filter: F1 })).toBe('none')
    expect(orderBytes(IDENTITY_VIEW)).toBe(0)
    // One range of 99 body rows (100 rows, 1 header), shared by sort, sortFilter and filter, is counted once.
    expect(orderBytes({ sort: S, sortFilter: F1, filter: F1 })).toBe(8 * 99 + 160 * 1024)
    expect(orderBytes({ sort: S, sortFilter: null, filter: null })).toBe(8 * 99 + 160 * 1024)
    // Distinct ranges count the union of their body rows: 1..99 and 50..199 (totals row 200 excluded).
    const other: ResolvedFilter = { range: { r0: 49, r1: 200, c0: 0, c1: 0, header: 1, totals: 1 }, columns: [] }
    expect(orderBytes({ sort: S, sortFilter: null, filter: other })).toBe(8 * 199 + 160 * 1024)
  })

  test('the 4c view hashes the 4b way when it is the identity', () => {
    expect(hashOf(IDENTITY_VIEW)).toBe(viewStateHash({ sort: null, filters: [], sizes: [], authorHiddenTier: 'none', clock }))
    expect(viewStateHash({ view: { sort: null, sortFilter: null, filter: F1 }, sizes: [], authorHiddenTier: 'c1', clock })).not.toBe(hashOf({ sort: null, sortFilter: null, filter: F1 }))
  })
})

describe('session view state (P11, F4)', () => {
  const sorted: CanonicalView = { sort: S, sortFilter: null, filter: null }
  const holds = () => {
    const log: string[] = []
    let live = 0
    const hold = (name: string) => () => { live++; log.push(`+${name}`); return () => { live--; log.push(`-${name}`) } }
    return { log, hold, live: () => live }
  }

  test('two ordered sheets per sid', () => {
    const h = holds()
    const v = new SessionViews(new ViewStateLedger())
    expect(MAX_ORDERED_SHEETS).toBe(2)
    expect(v.commit(0, sorted, h.hold('s0'))).toBe(1)
    expect(v.commit(1, { sort: null, sortFilter: null, filter: F1 }, h.hold('s1'))).toBe(1)
    expect(v.orderedSheets()).toBe(2)
    expect(v.commit(2, sorted, h.hold('s2'))).toBe('view-state-limit')
    expect(h.log).toEqual(['+s0', '+s1'])
    expect(v.view(2)).toEqual(IDENTITY_VIEW); expect(v.viewGen(2)).toBe(0)
    // An ordered sheet may change its own view at the cap; the new hold is taken before the old one is released.
    expect(v.commit(0, { sort: S, sortFilter: F1, filter: F1 }, h.hold('s0b'))).toBe(2)
    expect(h.log.slice(2)).toEqual(['+s0b', '-s0'])
    // Clearing sheet 0 back to identity skips the hold and frees a slot.
    let called = false
    expect(v.commit(0, { sort: null, sortFilter: F1, filter: null }, () => { called = true; return () => {} })).toBe(3)
    expect(called).toBe(false)
    expect(h.log.slice(4)).toEqual(['-s0b'])
    expect(v.orderedSheets()).toBe(1)
    expect(v.view(0)).toEqual(IDENTITY_VIEW)
    // Resizes on sheet 2 still work, before and after its ordered commit, and share its viewGen.
    expect(v.apply(2, { axis: 'col', index: 0, size: 100 })).toEqual({ viewGen: 1, applied: 100 })
    expect(v.commit(2, sorted, h.hold('s2'))).toBe(2)
    expect(v.apply(2, { axis: 'row', index: 0, size: 30 })).toEqual({ viewGen: 3, applied: 30 })
    expect(v.canonical(2)).toEqual([{ axis: 'col', index: 0, px: 100 }, { axis: 'row', index: 0, px: 30 }])
    expect(v.view(2)).toEqual(canonicalView(sorted))
    expect(h.live()).toBe(2)
    v.dispose(); v.dispose()
    expect(h.live()).toBe(0)
    expect(v.orderedSheets()).toBe(0)
  })

  test('a refused hold or an unchanged view changes nothing', () => {
    const h = holds()
    const v = new SessionViews(new ViewStateLedger())
    expect(v.commit(0, sorted, () => null)).toBe('view-state-limit')
    expect(v.viewGen(0)).toBe(0); expect(v.orderedSheets()).toBe(0)
    expect(v.commit(0, sorted, h.hold('a'))).toBe(1)
    expect(v.commit(0, { sort: null, sortFilter: null, filter: F1 }, () => null)).toBe('view-state-limit')
    expect(v.view(0)).toEqual(canonicalView(sorted)); expect(v.viewGen(0)).toBe(1); expect(h.live()).toBe(1)
    // Committing an equal view takes no new hold and keeps viewGen.
    expect(v.commit(0, { sort: { ...S, keys: [...S.keys] }, sortFilter: null, filter: null }, h.hold('b'))).toBe(1)
    expect(h.log).toEqual(['+a'])
    // Identity on an identity sheet keeps viewGen too.
    expect(v.commit(1, IDENTITY_VIEW, h.hold('c'))).toBe(0)
    v.dispose()
    expect(h.log).toEqual(['+a', '-a'])
  })
})
