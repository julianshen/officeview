import type { XlsxMergeRange } from '../types'
import type { AuthoritativeCell, CacheIdentity, Completeness, LogicalCellBox, LogicalViewport, ViewportResult } from './contracts'

/** Request-local wire evidence; it never certifies a row-cache chunk. */
export type OverflowEnd = 'blocker' | 'paint-boundary' | 'unknown'
export interface OverflowEntry {
  id: number
  row: number
  col: number
  left: number
  right: number
  leftEnd: OverflowEnd
  rightEnd: OverflowEnd
  x0: number
  x1: number
  h0: number
  h1: number
  c0: number
  c1: number
}
export interface OverflowProof {
  version: 1
  owner: CacheIdentity
  window: LogicalViewport
  profile: 'D' | 'M'
  scope: 'frame' | 'tile'
  entries: OverflowEntry[]
  pending: Array<{ bounds: LogicalCellBox; error: boolean }>
}
/** Compatibility name for existing consumers of the now-typed wire field. */
export type OverflowViewportResult = ViewportResult
export interface OverflowMerge { range: XlsxMergeRange; bounds: LogicalCellBox }
export interface OverflowCapture {
  structure(completeness: Completeness, merges: readonly OverflowMerge[]): void
  cell(cell: AuthoritativeCell): void
  run(range: { r0: number; r1: number; c0: number; c1: number }, complete: boolean): void
  page(payload: unknown, records: number): void
}
