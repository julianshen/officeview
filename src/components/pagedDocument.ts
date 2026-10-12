/**
 * Shell documents for worker-backed (paged) spreadsheets (plan S5.4). With
 * `interactiveSheets`, a paged admission opens in the resident viewer worker;
 * the loader returns an XlsxDocument carrying sheet names only (no rows), and
 * this marker carries the worker-backed source to <OfficeDoc>. Eager and tiny:
 * the source itself lives in a lazy chunk.
 */
import type { WorkbookSource } from '../worker/viewerSession'

export const PAGED_WORKBOOK = Symbol.for('officeview.pagedWorkbook')
export interface PagedWorkbookHandle {
  /** Borrowed by the viewer; releasing it there is a no-op. */
  readonly source: WorkbookSource
  /** Closes the worker document and its worker (owned by the loader's caller). */
  release(): Promise<void>
}
export function pagedWorkbookOf(doc: unknown): PagedWorkbookHandle | undefined {
  return typeof doc === 'object' && doc !== null ? (doc as { [PAGED_WORKBOOK]?: PagedWorkbookHandle })[PAGED_WORKBOOK] : undefined
}
