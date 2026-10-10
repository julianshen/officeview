/**
 * C4 phase 2 private iterative-calculation executor.
 *
 * Nonrecursive dependency/SCC analysis and the bounded in-place pass driver.
 * This module owns NO workbook state, is never exported from the public barrel,
 * adds no runtime dependency, and defines no public ABI: the caller supplies the
 * ACTUAL (lazily discovered) dependency edges, trusted numeric seeds, and a
 * per-cell evaluation callback.
 *
 * Native-established semantics (root-native-iteration/ITERATION-FACTS.json):
 *  - a cyclic cohort plus its downstream dependents is re-evaluated in canonical
 *    row/column order, in place: each cell reads already-updated earlier values
 *    and previous-pass later values (never a final post-pass dependent recompute);
 *  - the convergence metric is the GLOBAL strict maximum absolute change over the
 *    whole cohort (dependents and independent cycles included) and the stop test
 *    is `maxChange < delta`, never `<=`;
 *  - the pass count is bounded by iterateCount.
 */

/**
 * Upper bound on the number of iterative passes this executor will run. Above
 * this, the boundary is UNVERIFIED native behavior and the caller retains the
 * existing caches instead of guessing. (iterateCount=0 is likewise unverified.)
 */
export const ITERATIVE_PASS_HARD_CAP = 10000

/**
 * Iterative Tarjan strongly-connected components (explicit stack — never
 * recursive). Returns the SCC index of every reachable node.
 *
 * `edges` must only yield nodes contained in `nodes`; callers restrict the
 * graph to the recalculation set so retained caches never invent edges.
 */
export function stronglyConnectedComponents(
  nodes: readonly string[],
  edges: (node: string) => Iterable<string>,
): Map<string, number> {
  const index = new Map<string, number>()
  const low = new Map<string, number>()
  const onStack = new Set<string>()
  const stack: string[] = []
  const sccOf = new Map<string, number>()
  let counter = 0
  let sccCount = 0

  interface Work { node: string; iter: Iterator<string> }

  for (const root of nodes) {
    if (index.has(root)) continue
    index.set(root, counter)
    low.set(root, counter)
    counter++
    stack.push(root)
    onStack.add(root)
    const work: Work[] = [{ node: root, iter: edges(root)[Symbol.iterator]() }]

    while (work.length > 0) {
      const frame = work[work.length - 1]
      const next = frame.iter.next()
      if (!next.done) {
        const w = next.value
        if (!index.has(w)) {
          index.set(w, counter)
          low.set(w, counter)
          counter++
          stack.push(w)
          onStack.add(w)
          work.push({ node: w, iter: edges(w)[Symbol.iterator]() })
        } else if (onStack.has(w)) {
          low.set(frame.node, Math.min(low.get(frame.node)!, index.get(w)!))
        }
        continue
      }
      work.pop()
      if (work.length > 0) {
        const parent = work[work.length - 1].node
        low.set(parent, Math.min(low.get(parent)!, low.get(frame.node)!))
      }
      if (low.get(frame.node) === index.get(frame.node)) {
        let member: string
        do {
          member = stack.pop()!
          onStack.delete(member)
          sccOf.set(member, sccCount)
        } while (member !== frame.node)
        sccCount++
      }
    }
  }
  return sccOf
}

/**
 * Cyclic cells = members of a multi-node SCC or holders of an actual self-edge.
 * Self-edges are read from the ACTUAL graph, so a lazy branch that never reads
 * the cell back (IF(FALSE,A1+1,7)) is not cyclic.
 */
export function findCyclicCells(
  nodes: readonly string[],
  edges: (node: string) => Iterable<string>,
): { cyclic: Set<string>; sccOf: Map<string, number> } {
  const sccOf = stronglyConnectedComponents(nodes, edges)
  const sizes = new Map<number, number>()
  for (const s of sccOf.values()) sizes.set(s, (sizes.get(s) ?? 0) + 1)
  const cyclic = new Set<string>()
  for (const node of nodes) {
    if ((sizes.get(sccOf.get(node)!) ?? 0) > 1) {
      cyclic.add(node)
      continue
    }
    for (const w of edges(node)) {
      if (w === node) {
        cyclic.add(node)
        break
      }
    }
  }
  return { cyclic, sccOf }
}

/**
 * Transitive downstream closure (seeds included) over the reverse of the actual
 * dependency graph. A dependent participates in the cohort at its own row/column
 * position, so it must be re-evaluated every pass and measured by the metric.
 */
export function transitiveDependents(
  seeds: Iterable<string>,
  edges: (node: string) => Iterable<string>,
  universe: ReadonlySet<string>,
): Set<string> {
  const reverse = new Map<string, string[]>()
  for (const node of universe) {
    for (const w of edges(node)) {
      if (!universe.has(w)) continue
      const readers = reverse.get(w)
      if (readers) readers.push(node)
      else reverse.set(w, [node])
    }
  }
  const out = new Set<string>()
  const queue: string[] = []
  for (const seed of seeds) {
    if (!universe.has(seed) || out.has(seed)) continue
    out.add(seed)
    queue.push(seed)
  }
  while (queue.length > 0) {
    const current = queue.pop()!
    for (const reader of reverse.get(current) ?? []) {
      if (out.has(reader)) continue
      out.add(reader)
      queue.push(reader)
    }
  }
  return out
}

/**
 * Canonical order for the ONE-SHEET verified native profile: sheet workbook
 * index, then row, then column. The native workbook order for a cohort that
 * spans multiple sheets is NOT claimed here — the caller gates cross-sheet
 * cohorts as unverified and retains their caches.
 */
export function canonicalOrder(
  keys: Iterable<string>,
  positionOf: (key: string) => readonly [number, number, number],
): string[] {
  return [...keys].sort((a, b) => {
    const pa = positionOf(a)
    const pb = positionOf(b)
    if (pa[0] !== pb[0]) return pa[0] - pb[0]
    if (pa[1] !== pb[1]) return pa[1] - pb[1]
    if (pa[2] !== pb[2]) return pa[2] - pb[2]
    return a < b ? -1 : a > b ? 1 : 0
  })
}

/**
 * Bounded in-place pass driver. `pass` evaluates one whole cohort in order and
 * returns that pass's maximum absolute change; the loop stops as soon as the
 * strict inequality `maxChange < delta` holds, otherwise at `count` passes.
 */
export function runBoundedPasses(
  count: number,
  delta: number,
  pass: () => number,
  aborted?: () => boolean,
): { passes: number; converged: boolean; aborted: boolean } {
  let passes = 0
  for (let i = 0; i < count; i++) {
    passes++
    const maxChange = pass()
    // The pass may discover an unverified condition (e.g. an adaptive dependency
    // change); the caller aborts and retains the shipped caches.
    if (aborted?.()) return { passes, converged: false, aborted: true }
    if (maxChange < delta) return { passes, converged: true, aborted: false }
  }
  return { passes, converged: false, aborted: false }
}
