import { describe, expect, it, test } from 'bun:test'
import { DocumentLedger, LruBytes, renderKey, viewStateHash, type RenderKeyFields } from '../src/cache'

const base: RenderKeyFields = {
  docRef: 'a'.repeat(24), contentPolicyId: 'visible-only@1', layoutFingerprint: 'f'.repeat(64), recorderVersion: 1, scaleBucket: 1,
  markSettings: { visible: 'fixed', text: 'officeview · protected', forensic: false }, projectionMode: 'final-raster',
  policyDetails: { projection: 'visible-only@1' }, unitId: 0,
}

describe('§9.3 render key', () => {
  const variants: Array<[string, Partial<RenderKeyFields>]> = [
    ['docRef', { docRef: 'b'.repeat(24) }], ['contentPolicyId', { contentPolicyId: 'full@1' }], ['layoutFingerprint', { layoutFingerprint: 'e'.repeat(64) }],
    ['recorderVersion', { recorderVersion: 2 as never }], ['scaleBucket', { scaleBucket: 2 }],
    ['markSettings', { markSettings: { visible: 'viewer', text: 'v1', forensic: false } }], ['projectionMode', { projectionMode: 'commands' }],
    ['policyDetails', { policyDetails: { projection: 'visible-only@1', missingFaces: ['carlito-bold'] } }], ['unitId', { unitId: 1 }],
  ]
  it.each(variants)('changes when %s changes', (_n, patch) => { expect(renderKey({ ...base, ...patch })).not.toBe(renderKey(base)) })
  it('is independent of object key order', () => {
    const shuffled = Object.fromEntries(Object.entries(base).reverse()) as unknown as RenderKeyFields
    expect(renderKey(shuffled)).toBe(renderKey(base))
  })
  it('relaxed/protected and fixed/viewer settings share nothing, in both fill orders', () => {
    const keys = (order: number) => {
      const cache = new LruBytes<string>(1 << 20)
      const sessions = [{ ...base }, { ...base, projectionMode: 'commands' as const }, { ...base, markSettings: { visible: 'viewer' as const, text: 'alice', forensic: false as const } }, { ...base, markSettings: { visible: 'viewer' as const, text: 'bob', forensic: false as const } }]
      const seq = order ? [...sessions].reverse() : sessions
      return seq.map((s) => { const k = renderKey(s); const hit = cache.has(k); cache.put(k, 'x', 1); return hit })
    }
    expect(keys(0)).toEqual([false, false, false, false])
    expect(keys(1)).toEqual([false, false, false, false])
  })
})

describe('Stage 4a tile key (S4 §6.5)', () => {
  // Printed by `renderKey` at a56fa675, before `tile` and `viewStateHash` existed: unit keys keep their hashes.
  test('unit-key golden', () => {
    expect(renderKey(base)).toBe('e93595d59c4d7ec3c895b4775e16e0cf525cfe494199a7991068bb4b1a5f5d03')
    expect(renderKey({ ...base, projectionMode: 'commands', policyDetails: { projection: 'visible-only@1', missingFaces: ['carlito-bold'] }, unitId: 3, scaleBucket: 2 })).toBe('dbba2996e482a21267dcd8bea21bb2916e6dc9f235448af437c6f22a3c15392c')
  })
  const clock = { day: '2026-11-23', timeZone: 'UTC' }
  const vsh = viewStateHash({ sort: null, filters: [], sizes: [], authorHiddenTier: 'none', clock })
  const tileBase: RenderKeyFields = { ...base, scaleBucket: 1, policyDetails: { projection: 'visible-only@1', viewStateHash: vsh }, tile: { sheetIndex: 0, zoomStep: 5, dpr: 1, palette: 'light', col: 0, row: 0 } }
  test('viewStateHash is the sha256 hex of the canonical view, clock included', () => {
    expect(vsh).toMatch(/^[0-9a-f]{64}$/)
    expect(viewStateHash({ clock, authorHiddenTier: 'none', sizes: [], filters: [], sort: null })).toBe(vsh)
    expect(viewStateHash({ sort: null, filters: [], sizes: [], authorHiddenTier: 'none', clock: { ...clock, day: '2026-11-24' } })).not.toBe(vsh)
    expect(viewStateHash({ sort: null, filters: [], sizes: [], authorHiddenTier: 'none', clock: { ...clock, timeZone: 'Europe/London' } })).not.toBe(vsh)
  })
  const tileVariants: Array<[string, Partial<NonNullable<RenderKeyFields['tile']>>]> = [
    ['sheetIndex', { sheetIndex: 1 }], ['zoomStep', { zoomStep: 4 }], ['dpr', { dpr: 2 }], ['palette', { palette: 'dark' }], ['col', { col: 1 }], ['row', { row: 1 }],
  ]
  it.each(tileVariants)('a tile key changes when %s changes', (_n, patch) => {
    expect(renderKey({ ...tileBase, tile: { ...tileBase.tile!, ...patch } })).not.toBe(renderKey(tileBase))
  })
  test('the tile field and the view-state hash are keyed', () => {
    const { tile: _t, ...noTile } = tileBase
    expect(renderKey(noTile)).not.toBe(renderKey(tileBase))
    expect(renderKey({ ...tileBase, policyDetails: { projection: 'visible-only@1' } })).not.toBe(renderKey(tileBase))
    expect(renderKey({ ...tileBase, policyDetails: { projection: 'visible-only@1', viewStateHash: 'e'.repeat(64) } })).not.toBe(renderKey(tileBase))
  })
})

describe('byte-bounded LRU', () => {
  it('evicts least recently used first, never a pinned entry', () => {
    const evicted: string[] = []
    const c = new LruBytes<number>(100, (k) => evicted.push(k))
    c.put('a', 1, 40); c.put('b', 2, 40); c.get('a'); c.pin('b')
    c.put('c', 3, 40)
    expect(evicted).toEqual(['a'])       // b is pinned although older than a's touch
    expect(c.has('b') && c.has('c')).toBe(true)
    c.unpin('b'); c.put('d', 4, 40)
    expect(evicted).toEqual(['a', 'b'])
    expect(c.size).toBe(80)
    expect(c.put('huge', 5, 101)).toBe(false)
  })
})

describe('document ledger', () => {
  it('charges before parse, drops the least recently used past the ceiling, refuses what can never fit (review r2-3)', () => {
    const dropped: string[] = []
    const l = new DocumentLedger((d, p) => dropped.push(`${d}|${p}`), 100)
    expect(l.touch('d1', 'visible-only', 5)).toBe(true)   // 40
    expect(l.touch('d2', 'visible-only', 5)).toBe(true)   // 80
    expect(l.touch('d1', 'visible-only', 5)).toBe(true)   // refresh d1
    expect(l.touch('d3', 'full', 5)).toBe(true)           // 120 > 100 → drop d2
    expect(dropped).toEqual(['d2|visible-only'])
    expect(l.size).toBe(80)
    expect(l.touch('huge', 'visible-only', 13)).toBe(false)  // 104 > 100: refused, nothing held
    expect(l.size).toBe(80)
    expect(l.touch('d2', 'visible-only', 5)).toBe(true)   // a reparse after eviction is charged again
    expect(dropped).toEqual(['d2|visible-only', 'd1|visible-only'])
  })
})

describe('Stage 4b view-state hash (S4 §6.3)', () => {
  const clock = { day: '2026-11-23', timeZone: 'UTC' }
  test('identity hash unchanged', () => {
    // Printed by `viewStateHash` at 56431d14 (4a's identity view), under this clock.
    expect(viewStateHash({ sort: null, filters: [], sizes: [], authorHiddenTier: 'none', clock })).toBe('ed406e854c26f24450bbdd81d392b23810110369f6d89c8cd5d9302ebd51c645')
  })
  test('sizes are hashed', () => {
    const id = viewStateHash({ sort: null, filters: [], sizes: [], authorHiddenTier: 'none', clock })
    const a = viewStateHash({ sort: null, filters: [], sizes: [{ axis: 'col', index: 2, px: 1790 }], authorHiddenTier: 'none', clock })
    const b = viewStateHash({ sort: null, filters: [], sizes: [{ axis: 'col', index: 2, px: 1789 }], authorHiddenTier: 'none', clock })
    const c = viewStateHash({ sort: null, filters: [], sizes: [{ axis: 'row', index: 2, px: 1790 }], authorHiddenTier: 'none', clock })
    expect(new Set([id, a, b, c]).size).toBe(4)
    expect(viewStateHash({ clock, sizes: [{ px: 1790, index: 2, axis: 'col' }], authorHiddenTier: 'none', filters: [], sort: null })).toBe(a)
  })
})

// ---- Stage 4c Task 4 (P2, P11): the canonical view in the hash, and view holds in the document ledger ----
import { IDENTITY_VIEW } from '../src/sheet/view-spec'

describe('Stage 4c view-state hash (P2)', () => {
  const clock = { day: '2026-11-23', timeZone: 'UTC' }
  test('identity hashes unchanged', () => {
    // Printed by `viewStateHash` at fb624989 (4b), under this clock: the identity view, and a sizes-only view.
    const sizes = [{ axis: 'col' as const, index: 2, px: 1790 }, { axis: 'row' as const, index: 5, px: 40 }]
    expect(viewStateHash({ view: IDENTITY_VIEW, sizes: [], authorHiddenTier: 'none', clock })).toBe('ed406e854c26f24450bbdd81d392b23810110369f6d89c8cd5d9302ebd51c645')
    expect(viewStateHash({ view: IDENTITY_VIEW, sizes, authorHiddenTier: 'none', clock })).toBe('af1e4ca982de561a2e671905e6fe66234c5b6fca2fa693a2d7f461dc1aec9fe0')
    expect(viewStateHash({ sort: null, filters: [], sizes, authorHiddenTier: 'none', clock })).toBe('af1e4ca982de561a2e671905e6fe66234c5b6fca2fa693a2d7f461dc1aec9fe0')
  })
})

describe('Stage 4c view holds (P11)', () => {
  test('admission', () => {
    const MiB = 1 << 20
    const dropped: string[] = []
    const l = new DocumentLedger((d, p) => dropped.push(`${d}|${p}`), 10 << 20)
    expect(l.touch('d1', 'visible-only', (4 * MiB) / 8)).toBe(true)   // a 4 MiB parse
    expect(l.hold('v1', 8 * MiB)).toBe(true)                           // 12 MiB > 10 MiB: the parse is evicted
    expect(dropped).toEqual(['d1|visible-only'])
    expect(l.size).toBe(8 * MiB)
    expect(l.hold('v1', 8 * MiB)).toBe(true)                           // a second holder charges nothing more
    expect(l.size).toBe(8 * MiB)
    expect(l.hold('v2', 8 * MiB)).toBe(false)                          // a third distinct hold cannot fit
    expect(l.size).toBe(8 * MiB)
    expect(dropped).toEqual(['d1|visible-only'])
    expect(l.touch('d2', 'visible-only', (4 * MiB) / 8)).toBe(false)  // a held view is never evicted for a parse
    l.release('v1')
    expect(l.size).toBe(8 * MiB)
    l.release('v1')                                                     // both holders gone: freed
    expect(l.size).toBe(0)
    l.release('v1')                                                     // an extra release is a no-op
    expect(l.size).toBe(0)
    expect(l.hold('v2', 8 * MiB)).toBe(true)
    // A hold key never meets a parse key.
    expect(l.touch('v2', 'visible-only', 1)).toBe(true)
    expect(l.size).toBe(8 * MiB + 8)
    expect(dropped).toEqual(['d1|visible-only'])
  })
})
