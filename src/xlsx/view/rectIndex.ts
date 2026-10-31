/**
 * Cell-to-owner lookups over owned rectangles (Task 33 perf fix): a pivot's overlap check and the composite pivot
 * overlays ask, per cell, which pivots' rectangles hold it. A `RectangleIndex` answers in O(log n + hits) instead of a
 * walk over every pivot and each of its rectangles. Shared by the main thread and the worker.
 */
import type { SourceRect } from './coordinates'
import { RectangleIndex } from './structure'

let queries = 0, visits = 0
/** Point queries made through `OwnedRects`, and the index nodes they visited (tests bound the work with these). */
export const ownedRectsWork = (): { queries: number; visits: number } => ({ queries, visits })
const count = (n: number) => { visits += n }

/** `rects` without repeats (by their four bounds), first occurrence kept, in one keyed pass. */
export function uniqueRectList<R extends SourceRect>(rects: Iterable<R>): R[] {
  const seen = new Set<string>(), out: R[] = []
  for (const rect of rects) {
    const key = `${rect.r0},${rect.r1},${rect.c0},${rect.c1}`
    if (!seen.has(key)) { seen.add(key); out.push(rect) }
  }
  return out
}

/** Rectangles with an owner each, queried by cell. */
export class OwnedRects<T> {
  private readonly index: RectangleIndex<T>
  readonly size: number
  constructor(entries: Iterable<{ rect: SourceRect; owner: T }>) {
    const list: { rect: SourceRect; value: T }[] = []
    for (const { rect, owner } of entries) list.push({ rect, value: owner })
    this.index = RectangleIndex.build(list)
    this.size = list.length
  }
  /** Retained bytes of the index (its entries and interval tree). */
  get estimatedBytes(): number { return this.index.estimatedBytes }
  /** The owner of every rectangle holding the cell (one per rectangle). */
  at(row: number, col: number): readonly T[] {
    if (!this.size) return []
    queries++
    return this.index.walkFrom({ r0: row, r1: row, c0: col, c1: col }, 0, Number.MAX_SAFE_INTEGER, { countVisits: count }).items
  }
  /** The first owner holding the cell that `pred` accepts (early exit). */
  find(row: number, col: number, pred: (owner: T) => boolean): T | undefined {
    if (!this.size) return undefined
    queries++
    return this.index.walkFrom({ r0: row, r1: row, c0: col, c1: col }, 0, 1, { pred, countVisits: count }).items[0]
  }
  /** Whether any rectangle holds the cell. */
  has(row: number, col: number): boolean { return this.find(row, col, () => true) !== undefined }
}

/** Rectangles queried for containment: whether one of them contains all of a rectangle. */
export class RectCover {
  private readonly owned: OwnedRects<SourceRect>
  constructor(bounds: Iterable<SourceRect>) {
    const entries: { rect: SourceRect; owner: SourceRect }[] = []
    for (const rect of bounds) entries.push({ rect, owner: rect })
    this.owned = new OwnedRects(entries)
  }
  /** A containing bound holds the rectangle's first cell, so only the bounds at that cell are compared. */
  contains(rect: SourceRect): boolean { return this.owned.find(rect.r0, rect.c0, bound => bound.r1 >= rect.r1 && bound.c1 >= rect.c1) !== undefined }
  has(row: number, col: number): boolean { return this.owned.has(row, col) }
}
