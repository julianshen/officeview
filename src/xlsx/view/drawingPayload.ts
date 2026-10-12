/**
 * One sheet's drawing payload for the worker-backed viewer (S13.2; spec
 * RENDER-02, §12.4). Both source adapters hand the main thread the drawings the
 * resident parser produced, re-indexed per sheet, within one transfer budget:
 * pictures (and any other binary asset) are charged in authored paint order and
 * an asset that does not fit is withdrawn and named in a diagnostic, never
 * dropped silently. The drawable itself stays (it paints as a refused picture
 * does on the resident path: blank), so placement parity is kept.
 *
 * Every binary in the returned payload is owned by it (callers pass freshly
 * parsed or cloned drawings): the worker transfers those buffers.
 */
import type { XlsxDrawing } from '../types'
import type { ThemeContext } from '../../drawing/style'
import type { ProfileId, SheetDrawings, WorkerDiagnostic } from './contracts'
import { collectXlsxImages } from '../drawingImages'
import { MAIN_THREAD_BUDGETS } from './tileCache'

export const DRAWING_ASSET_REFUSED = 'drawing-asset-refused'
export const DRAWING_PART_REFUSED = 'drawing-part-refused'
export const DRAWINGS_UNAVAILABLE = 'drawings-unavailable'

/** Binary bytes one sheet's drawings may carry to the main thread: a quarter
 * of the main-thread profile budget (D 16 MiB, M 8 MiB), like the row cache. */
export function drawingPayloadBudget(profile: ProfileId): number { return MAIN_THREAD_BUDGETS[profile] / 4 }

/** An object that owns picture bytes: a scene/content image or an SVG candidate holder. */
interface Asset { data?: Uint8Array; svg?: { bytes: Uint8Array } }
function isAsset(value: object): value is Asset {
  // isView, not instanceof: a cloned array may come from another realm.
  const candidate = value as Asset
  return ArrayBuffer.isView(candidate.data) || ArrayBuffer.isView(candidate.svg?.bytes)
}

/**
 * Charges and indexes `drawings` (mutated in place: they must be owned by the
 * payload). Assets are visited depth-first in paint order; a buffer shared by
 * several assets is charged once.
 */
export function sheetDrawingPayload(drawings: XlsxDrawing[], theme: ThemeContext | undefined, diagnostics: WorkerDiagnostic[], budgetBytes: number): SheetDrawings {
  const charged = new Set<ArrayBufferLike>()
  const seen = new Set<object>()
  let bytes = 0, refused = 0, refusedBytes = 0
  const sizeOf = (asset: Asset): number => {
    let size = 0
    // A view transfers (and a clone copies) its whole buffer: charge that.
    const buffers = new Set<ArrayBufferLike>()
    for (const array of [asset.data, asset.svg?.bytes]) if (array && !charged.has(array.buffer)) buffers.add(array.buffer)
    for (const buffer of buffers) size += buffer.byteLength
    return size
  }
  const reached = new Set<ArrayBufferLike>()
  const visit = (value: unknown): void => {
    if (typeof value !== 'object' || value === null || seen.has(value)) return
    if (ArrayBuffer.isView(value)) { reached.add(value.buffer); return }
    seen.add(value)
    if (isAsset(value)) {
      const size = sizeOf(value)
      if (bytes + size <= budgetBytes) {
        bytes += size
        for (const array of [value.data, value.svg?.bytes]) if (array) charged.add(array.buffer)
      } else {
        refused++
        refusedBytes += size
        value.data = undefined
        delete value.svg
      }
    }
    // Indexes are re-assigned below over the bytes that stayed.
    if ('imageIndex' in value) delete (value as { imageIndex?: number }).imageIndex
    for (const entry of Array.isArray(value) ? value : Object.values(value)) visit(entry)
  }
  visit(drawings)
  // Binaries outside a picture asset still cross the boundary: charge them, and
  // carry no drawings at all rather than exceed the budget.
  for (const buffer of reached) if (!charged.has(buffer)) bytes += buffer.byteLength
  if (bytes > budgetBytes) return unavailableDrawings(`This sheet's drawings need more than the ${budgetBytes}-byte large-file drawing budget and are not shown`)
  if (refused) {
    diagnostics.push({ code: DRAWING_ASSET_REFUSED, message: `${refused} picture${refused === 1 ? '' : 's'} (${refusedBytes} bytes) exceed the ${budgetBytes}-byte large-file drawing budget and are not shown` })
  }
  const images = collectXlsxImages([{ name: '', rows: [], cols: [], merges: [], mergeRanges: [], drawings }])
  return { completeness: 'complete', drawings, images, ...(theme ? { theme } : {}), diagnostics, bytes }
}

/** A typed outcome without drawings: the reason is always a diagnostic. */
export function unavailableDrawings(message: string, error?: SheetDrawings['error']): SheetDrawings {
  return { completeness: 'incomplete', drawings: [], images: [], diagnostics: [{ code: DRAWINGS_UNAVAILABLE, message }], bytes: 0, ...(error ? { error } : {}) }
}
