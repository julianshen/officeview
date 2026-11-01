/**
 * Cell-to-owner lookups over owned rectangles (Task 33 perf fix): a pivot's overlap check and the composite pivot
 * overlays ask, per cell, which pivots' rectangles hold it. A two-level index answers in O(log² n + hits) instead of a
 * walk over every pivot and each of its rectangles. Shared by the main thread and the worker.
 */
import type { SourceRect } from './coordinates'

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

/**
 * A two-level point-stabbing index (perf round 2: column-aware). A segment tree over the rows' elementary intervals
 * stores each rectangle at its O(log n) canonical nodes; each node keeps its rectangles in a column stabbing tree
 * (sorted by `c0`, pruned by the subtree's largest `c1`). A point query walks one root-to-leaf row path and stabs each
 * node's column tree: O(log² n + hits) whatever the shape, side-by-side pivots in one row band included. Bounds are
 * inclusive; a rectangle given twice is reported twice.
 */
class PointIndex<T> {
  private readonly rows: number[]
  /** Per segment-tree node: its rectangles' columns, sorted by c0, and their max-c1 tree. */
  private readonly nodes: Array<{ c0: Float64Array; c1: Float64Array; ids: Int32Array; tree: Float64Array } | undefined>
  private readonly values: readonly T[]
  readonly bytes: number
  constructor(entries: readonly { rect: SourceRect; value: T }[]) {
    this.values = entries.map(entry => entry.value)
    const bounds = new Set<number>()
    for (const { rect } of entries) { bounds.add(rect.r0); bounds.add(rect.r1 + 1) }
    this.rows = [...bounds].sort((a, b) => a - b)
    const segments = Math.max(0, this.rows.length - 1)
    const lists: number[][] = []
    const lower = (value: number) => { let lo = 0, hi = this.rows.length; while (lo < hi) { const mid = (lo + hi) >> 1; if (this.rows[mid]! < value) lo = mid + 1; else hi = mid } return lo }
    const insert = (node: number, lo: number, hi: number, a: number, b: number, id: number) => {
      if (b <= lo || hi <= a) return
      if (a <= lo && hi <= b) { (lists[node] ??= []).push(id); return }
      const mid = (lo + hi) >> 1
      insert(node * 2, lo, mid, a, b, id); insert(node * 2 + 1, mid, hi, a, b, id)
    }
    entries.forEach(({ rect }, id) => { if (segments) insert(1, 0, segments, lower(rect.r0), lower(rect.r1 + 1), id) })
    // Its row bounds, the node table, and the owners array (one reference per rectangle: codex perf round 2 follow-up).
    let bytes = 64 + 8 * this.rows.length + 8 * lists.length + 16 + 8 * this.values.length
    this.nodes = lists.map(list => {
      if (!list) return undefined
      const sorted = [...list].sort((a, b) => entries[a]!.rect.c0 - entries[b]!.rect.c0)
      const c0 = Float64Array.from(sorted, id => entries[id]!.rect.c0), c1 = Float64Array.from(sorted, id => entries[id]!.rect.c1)
      const tree = new Float64Array(4 * sorted.length).fill(-1)
      const build = (t: number, l: number, r: number): number => {
        if (l + 1 === r) return (tree[t] = c1[l]!)
        const m = (l + r) >> 1
        return (tree[t] = Math.max(build(t * 2, l, m), build(t * 2 + 1, m, r)))
      }
      build(1, 0, sorted.length)
      bytes += 96 + 20 * sorted.length + tree.byteLength
      return { c0, c1, ids: Int32Array.from(sorted), tree }
    })
    this.bytes = bytes
  }
  /** Calls `hit` with each owner holding the cell until it returns true; returns the node visits. */
  stab(row: number, col: number, hit: (value: T) => boolean): number {
    const segments = this.rows.length - 1
    if (segments <= 0 || row < this.rows[0]! || row >= this.rows[segments]!) return 0
    let lo = 0, hi = this.rows.length
    while (lo + 1 < hi) { const mid = (lo + hi) >> 1; if (this.rows[mid]! <= row) lo = mid; else hi = mid }
    const segment = lo
    let visits = 0, stopped = false
    const column = (entry: NonNullable<PointIndex<T>['nodes'][number]>) => {
      let a = 0, b = entry.c0.length
      while (a < b) { const m = (a + b) >> 1; if (entry.c0[m]! <= col) a = m + 1; else b = m }
      const end = a
      const walk = (t: number, l: number, r: number): void => {
        visits++
        if (stopped || l >= end || entry.tree[t]! < col) return
        if (l + 1 === r) { if (entry.c1[l]! >= col && hit(this.values[entry.ids[l]!]!)) stopped = true; return }
        const m = (l + r) >> 1
        walk(t * 2, l, m); walk(t * 2 + 1, m, r)
      }
      if (end > 0) walk(1, 0, entry.c0.length)
    }
    let node = 1, l = 0, r = segments
    for (;;) {
      visits++
      const entry = this.nodes[node]
      if (entry) column(entry)
      if (stopped || l + 1 === r) break
      const m = (l + r) >> 1
      if (segment < m) { node = node * 2; r = m } else { node = node * 2 + 1; l = m }
    }
    return visits
  }
}

/** Rectangles with an owner each, queried by cell. */
export class OwnedRects<T> {
  private readonly index: PointIndex<T>
  readonly size: number
  constructor(entries: Iterable<{ rect: SourceRect; owner: T }>) {
    const list: { rect: SourceRect; value: T }[] = []
    for (const { rect, owner } of entries) list.push({ rect, value: owner })
    this.index = new PointIndex(list)
    this.size = list.length
  }
  /** Retained bytes of the index (its row bounds, owners, node column arrays and trees). */
  get estimatedBytes(): number { return this.index.bytes }
  /** An upper bound of the bytes `rects` rectangles add to an index of up to 2^21 rectangles (each stored at no more
   * than 2·21 + 2 canonical nodes; per stored entry its columns, id and tree slots, plus a node's overhead; per
   * rectangle its row bounds and owner reference), so a
   * composite of several owners' shares is charged by summing them. */
  static upperBytesFor(rects: number): number { return 80 + rects * (24 + 44 * (20 + 32 + 104)) }
  /** The owner of every rectangle holding the cell (one per rectangle). */
  at(row: number, col: number): readonly T[] {
    if (!this.size) return []
    queries++
    const out: T[] = []
    count(this.index.stab(row, col, owner => { out.push(owner); return false }))
    return out
  }
  /** The first owner holding the cell that `pred` accepts (early exit). */
  find(row: number, col: number, pred: (owner: T) => boolean): T | undefined {
    if (!this.size) return undefined
    queries++
    let found: T | undefined
    count(this.index.stab(row, col, owner => { if (!pred(owner)) return false; found = owner; return true }))
    return found
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
