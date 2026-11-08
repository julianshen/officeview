import { describe, expect, it } from 'bun:test'
import { DocumentLedger, LruBytes, renderKey, type RenderKeyFields } from '../src/cache'

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
  it('drops the least recently used parse past the ceiling', () => {
    const dropped: string[] = []
    const l = new DocumentLedger((d, p) => dropped.push(`${d}|${p}`), 100)
    l.touch('d1', 'visible-only@1', 'fp', 5)   // 40
    l.touch('d2', 'visible-only@1', 'fp', 5)   // 80
    l.touch('d1', 'visible-only@1', 'fp', 5)   // refresh d1
    l.touch('d3', 'full@1', 'fp', 5)           // 120 > 100 → drop d2
    expect(dropped).toEqual(['d2|visible-only@1'])
    expect(l.size).toBe(80)
  })
})
