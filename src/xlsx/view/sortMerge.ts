import type { CooperativeSlice } from './scheduler'

/** Stable natural-run merge over owned index buffers. Boundary storage is supplied and charged by the caller. */
export async function mergeIndexRuns<Phase extends string>(
  index: Uint32Array, scratch: Uint32Array, starts: Uint32Array, count: number,
  compare: (a: number, b: number) => number, slice: CooperativeSlice<Phase>, phase: Phase,
): Promise<Uint32Array> {
  let runs = 0, start = 0
  starts[0] = 0
  while (start < count) {
    let end = start + 1
    if (end < count) {
      const descending = compare(index[end - 1], index[end]) > 0
      if (slice.due()) await slice.pause(phase)
      end++
      while (end < count) {
        const comparison = compare(index[end - 1], index[end])
        if (slice.due()) await slice.pause(phase)
        // Strict descending runs only: reversing equal keys would break stable ties.
        if (descending ? comparison <= 0 : comparison > 0) break
        end++
      }
      if (descending) for (let a = start, b = end - 1; a < b; a++, b--) {
        const swap = index[a]; index[a] = index[b]
        if (slice.due()) await slice.pause(phase)
        index[b] = swap
        if (slice.due()) await slice.pause(phase)
      }
    }
    starts[++runs] = end
    start = end
  }
  let input = index, output = scratch
  while (runs > 1) {
    let nextRuns = 0
    for (let run = 0; run < runs; run += 2) {
      const start = starts[run], middle = starts[run + 1], end = starts[Math.min(run + 2, runs)]
      let ordered = middle === end
      if (!ordered) {
        ordered = compare(input[middle - 1], input[middle]) <= 0
        if (slice.due()) await slice.pause(phase)
      }
      if (ordered) {
        // No comparator work for an already ordered boundary; copies stay bounded.
        for (let at = start; at < end; at++) {
          output[at] = input[at]
          if (slice.due()) await slice.pause(phase)
        }
      } else {
        let left = start, right = middle
        for (let at = start; at < end; at++) {
          output[at] = left < middle && (right >= end || compare(input[left], input[right]) <= 0) ? input[left++] : input[right++]
          if (slice.due()) await slice.pause(phase)
        }
      }
      starts[++nextRuns] = end
    }
    runs = nextRuns
    const swap = input; input = output; output = swap
  }
  return input
}
