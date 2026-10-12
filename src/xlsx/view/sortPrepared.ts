import type { SingleCellPayload, SourceStamp, SourceStore } from './contracts'

/** Internal prepared-store capability. Complete, validated values remain pinned
 * during the read; the consumer keeps its own key/ownership ledger. */
export const PREPARED_SORT_READER = Symbol('prepared-sort-reader')
export interface PreparedSortLease {
  readonly stamp: SourceStamp
  read(r0: number, r1: number, put: (row: number, value: SingleCellPayload['value'], error: boolean, unavailable: boolean) => void,
    cooperate: { due(): boolean; pause(): Promise<void> }): Promise<'complete' | 'stale'>
  release(): void
}
export type PreparedSortStore = SourceStore & {
  [PREPARED_SORT_READER]?: (col: number, stamp: SourceStamp) => PreparedSortLease | undefined
}
