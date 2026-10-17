// @vitest-environment node
import { runInNewContext } from 'node:vm'
import { describe, expect, it } from 'vitest'
import { adoptRowOrder, createIdentityOrder } from '../src/xlsx/view/order'
import type { RowPermutationBand } from '../src/xlsx/view/contracts'

const band = (): RowPermutationBand => ({ r0: 2, r1: 4, sourceAtOffset: new Uint32Array([4, 2, 3]), displayOffsetOfSource: new Uint32Array([1, 2, 0]) })

describe('adoptRowOrder exclusive buffer handoff', () => {
  it('adopts both arrays without copying, sorts bands, and preserves identity outside them', () => {
    const b = band(), order = adoptRowOrder([b, { r0: 0, r1: 0, sourceAtOffset: new Uint32Array([0]), displayOffsetOfSource: new Uint32Array([0]) }])
    expect(order.bands[1].sourceAtOffset).toBe(b.sourceAtOffset)
    expect(order.bands[1].displayOffsetOfSource).toBe(b.displayOffsetOfSource)
    expect([1, 2, 3, 4, 5].map(order.displayToSource)).toEqual([1, 4, 2, 3, 5])
    expect([1, 2, 3, 4, 5].map(order.sourceToDisplay)).toEqual([1, 3, 4, 2, 5])
    expect(order.diagnostics.typedArrayBytes).toBe(32)
    expect(Object.isFrozen(order.bands[1])).toBe(true)
    expect(Object.isFrozen(order.bands)).toBe(true)
    expect(adoptRowOrder([])).toBe(createIdentityOrder())
  })
  it('re-views foreign-realm Uint32Arrays without copying their buffers', () => {
    const b = runInNewContext('({ r0: 2, r1: 4, sourceAtOffset: new Uint32Array([4,2,3]), displayOffsetOfSource: new Uint32Array([1,2,0]) })') as RowPermutationBand
    const order = adoptRowOrder([b])
    expect(order.bands[0].sourceAtOffset).toBeInstanceOf(Uint32Array)
    expect(order.bands[0].sourceAtOffset.buffer).toBe(b.sourceAtOffset.buffer)
    expect(order.sourceToDisplay(4)).toBe(2)
  })
  it.each([
    ['same array', () => { const a = new Uint32Array([0, 1]); return [{ r0: 0, r1: 1, sourceAtOffset: a, displayOffsetOfSource: a }] }],
    ['distinct views of one buffer', () => { const a = new Uint32Array([0, 1, 0, 1]); return [{ r0: 0, r1: 1, sourceAtOffset: a.subarray(0, 2), displayOffsetOfSource: a.subarray(2) }] }],
    ['buffer reused between bands', () => { const a = new Uint32Array([0, 1]); return [{ r0: 0, r1: 0, sourceAtOffset: a.subarray(0, 1), displayOffsetOfSource: new Uint32Array([0]) }, { r0: 1, r1: 1, sourceAtOffset: a.subarray(1), displayOffsetOfSource: new Uint32Array([0]) }] }],
    ['padded whole-prefix view', () => [{ r0: 0, r1: 0, sourceAtOffset: new Uint32Array(new ArrayBuffer(1024), 0, 1), displayOffsetOfSource: new Uint32Array([0]) }]],
    ['offset padded view', () => [{ r0: 0, r1: 0, sourceAtOffset: new Uint32Array(new ArrayBuffer(1024), 128, 1), displayOffsetOfSource: new Uint32Array([0]) }]],
    ['SharedArrayBuffer', () => [{ r0: 0, r1: 0, sourceAtOffset: new Uint32Array(new SharedArrayBuffer(4)), displayOffsetOfSource: new Uint32Array([0]) }]],
  ] as const)('rejects %s because ownership cannot be exclusive', (_name, make) => {
    expect(() => adoptRowOrder(make())).toThrow(RangeError)
  })
  it.each([
    ['wrong source type', { sourceAtOffset: new Int32Array([4, 2, 3]) }],
    ['wrong inverse type', { displayOffsetOfSource: [1, 2, 0] }],
    ['short source', { sourceAtOffset: new Uint32Array([4, 2]) }],
    ['short inverse', { displayOffsetOfSource: new Uint32Array([1, 2]) }],
    ['duplicate source', { sourceAtOffset: new Uint32Array([4, 4, 3]) }],
    ['wrong inverse', { displayOffsetOfSource: new Uint32Array([2, 1, 0]) }],
    ['source below band', { sourceAtOffset: new Uint32Array([1, 2, 3]) }],
    ['source above band', { sourceAtOffset: new Uint32Array([5, 2, 3]) }],
    ['reversed band', { r0: 4, r1: 2 }],
    ['fractional band', { r0: 2.5 }],
    ['outside worksheet', { r1: 1048576 }],
  ])('rejects malformed %s', (_name, change) => {
    expect(() => adoptRowOrder([{ ...band(), ...change } as RowPermutationBand])).toThrow(RangeError)
  })
  it('rejects inclusive overlapping bands', () => {
    expect(() => adoptRowOrder([band(), band()])).toThrow(/Overlapping/)
  })
})
