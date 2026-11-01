/** Rectangles (source coordinates, inclusive) whose dependent readers a pivot still owes a recompute. */
export interface OwedRect { readonly r0: number; readonly r1: number; readonly c0: number; readonly c1: number }

const contains = (outer: OwedRect, inner: OwedRect) => outer.r0 <= inner.r0 && outer.r1 >= inner.r1 && outer.c0 <= inner.c0 && outer.c1 >= inner.c1

/** The canonical owed set of `owed` ∪ `added` (T31 r5): no duplicate and no rectangle another one contains, so
 * repeated commits that owe the same areas (every over-budget retry re-adds the saved and shown areas) keep it bounded
 * by the distinct areas, never by the number of commits. */
export function mergeOwedRects(owed: readonly OwedRect[], added: readonly OwedRect[]): OwedRect[] {
  const all = [...owed, ...added].sort((a, b) => (b.r1 - b.r0 + 1) * (b.c1 - b.c0 + 1) - (a.r1 - a.r0 + 1) * (a.c1 - a.c0 + 1))
  const kept: OwedRect[] = []
  for (const rect of all) if (!kept.some(other => contains(other, rect))) kept.push(rect)
  return kept
}
