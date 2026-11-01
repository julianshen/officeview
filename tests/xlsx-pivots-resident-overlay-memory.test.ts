// @vitest-environment node
import { describe, expect, it } from 'vitest'
import { compositeOverlay } from '../src/components/xlsx/pivot/pivotClient'

describe('resident pivot overlay memory', () => {
  it('does not build a per-cell Map when composing an overlay', () => {
    const count = 20_000
    const cells = Array.from({ length: count }, (_, index) => ({
      row: Math.floor(index / 10), col: index % 10,
      value: { kind: 'number' as const, value: index },
    }))
    const rect = { r0: 0, r1: Math.floor((count - 1) / 10), c0: 0, c1: 9 }
    const layout = { rect: { minRow: 0, maxRow: rect.r1, minCol: 0, maxCol: 9 }, axes: { rows: [], cols: [] }, cells, merges: [] }
    const originalSet = Map.prototype.set
    let mapWrites = 0
    Map.prototype.set = function <K, V>(this: Map<K, V>, key: K, value: V) {
      mapWrites++
      return originalSet.call(this, key, value)
    }
    let overlay: ReturnType<typeof compositeOverlay>
    try {
      overlay = compositeOverlay([{ saved: [rect], current: [rect], layout: layout as never }])
    } finally {
      Map.prototype.set = originalSet
    }
    expect(mapWrites).toBeLessThan(count / 2)
    expect(overlay!.cell(1234, 5)).toEqual({ value: 12_345 })
    expect(overlay!.indexBytes).toBeLessThan(count * 2)
  })
})
