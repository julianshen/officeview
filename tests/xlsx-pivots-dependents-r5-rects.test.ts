// Round 5 (codex r5 finding 1): the rectangles a pivot owes its dependents stay one canonical, bounded set.
import { expect, it } from 'vitest'
import { mergeOwedRects, type OwedRect } from '../src/xlsx/pivots/session/dependentRects'

const key = (rect: OwedRect) => `${rect.r0}:${rect.r1}:${rect.c0}:${rect.c1}`

it('20 consecutive over-budget commits leave the changed rectangles merged, with no duplicates', () => {
  const saved = { r0: 2, r1: 23, c0: 0, c1: 8 }, collapsed = { r0: 2, r1: 18, c0: 0, c1: 8 }, grown = { r0: 2, r1: 30, c0: 0, c1: 8 }
  let owed: OwedRect[] = []
  // Each commit owes saved ∪ before ∪ current; nothing ran, so the owed set carries into the next commit.
  for (let commit = 0; commit < 20; commit++) {
    const before = commit % 2 ? collapsed : grown, current = commit % 2 ? grown : collapsed
    owed = mergeOwedRects(owed, [saved, before, current])
    expect(new Set(owed.map(key)).size).toBe(owed.length)
    expect(owed.length).toBeLessThanOrEqual(3)
  }
  // The grown area contains the other two.
  expect(owed).toEqual([grown])
})

it('disjoint areas are all kept, once each', () => {
  const a = { r0: 0, r1: 4, c0: 0, c1: 4 }, b = { r0: 10, r1: 12, c0: 10, c1: 12 }
  let owed: OwedRect[] = []
  for (let commit = 0; commit < 20; commit++) owed = mergeOwedRects(owed, [a, b, { ...a }])
  expect(owed.map(key).sort()).toEqual([key(a), key(b)].sort())
})
