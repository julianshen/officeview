/**
 * Stage 4a (S4 §6.1, R16): the narrowed source surface `InteractiveSheetDoc` reaches through a non-resident workbook.
 * Two sources implement it: `RemoteWorkbook` (paged, kind 'paged': a resident worker owns the cells) and
 * `StreamWorkbook` (kind 'stream': the server paints the tiles and the client holds no cell data). The viewer reaches
 * these shared members on either, and every other member only under `kind === 'paged'` (widened, never subclassed:
 * a paged method added later is not reachable on a stream unless someone adds it here on purpose).
 *
 * Type-only imports: this module adds no runtime edge (the discriminant is a plain field).
 */
import type { CacheIdentity, FilterKeySet, FilterSpec, SheetLayout, SortSpec, SourceStore } from '../../xlsx/view/contracts'
import type { XlsxMergeRange } from '../../xlsx/types'
import type { SheetGeometry, WorkbookSource } from '../../worker/viewerSession'
import type { RemoteCommit, RemoteSheetView, RemoteWorkbook } from '../../worker/viewerRemoteSource'
import type { StreamWorkbook } from './streamWorkbook'

export type SourceKind = 'paged' | 'stream'

/** R16's shared members (signatures copied from `RemoteWorkbook`), plus `activate`, which the recount adds: on a
 * stream it fetches the sheet's `sheet-layout`, so it is asynchronous there (paged returns at once). */
export interface ViewerSheetSource extends WorkbookSource {
  readonly kind: SourceKind
  layout(sheetId: string): SheetLayout
  merges(sheetId: string): { gen: number; merges: readonly XlsxMergeRange[]; sourceMerges: readonly XlsxMergeRange[] }
  store(sheetId: string): SourceStore
  identity(sheetId: string): CacheIdentity
  geometry(sheetId: string): SheetGeometry
  readonly isClosed: boolean
  subscribe(listener: () => void): () => void
  loadState(sheetId: string): { completeness: 'complete' | 'loading' | 'incomplete'; failure?: string; code?: string }
  view(sheetId: string): RemoteSheetView
  sort(owner: CacheIdentity, spec: SortSpec | null, signal: AbortSignal, progress?: (completed: number, total: number) => void): Promise<RemoteCommit>
  filter(owner: CacheIdentity, spec: FilterSpec | null, keySets: readonly FilterKeySet[], signal: AbortSignal, progress?: (completed: number, total: number) => void): Promise<RemoteCommit>
  activate(sheetId: string): void | Promise<void>
}

/** The viewer's discriminant check (replaces the old `instanceof RemoteWorkbook` guard): a resident source is neither. */
export function sheetSourceOf(source: WorkbookSource | undefined): { paged: RemoteWorkbook | null; stream: StreamWorkbook | null } {
  const kind = (source as { readonly kind?: unknown } | undefined)?.kind
  return {
    paged: kind === 'paged' ? source as RemoteWorkbook : null,
    stream: kind === 'stream' ? source as StreamWorkbook : null,
  }
}

/** Compile-time only: both sources implement the shared surface (a signature drift fails tsc, no runtime code). */
type AssertExtends<A extends B, B> = A
export type PagedViewerSource = AssertExtends<RemoteWorkbook, ViewerSheetSource>
export type StreamViewerSource = AssertExtends<StreamWorkbook, ViewerSheetSource>
