import type { SourceStore } from './contracts'

/** Internal capability: only the prepared store opts in; raw adapters keep
 * cell entries so requesting lanes cannot expand repeated giant strings. */
export const BOUNDED_SORT_LANES = Symbol('bounded-prepared-sort-lanes')
export type BoundedSortLaneStore = SourceStore & { readonly [BOUNDED_SORT_LANES]?: true }
