/**
 * Drawing content of one paged sheet (S13.2; spec RENDER-02, AC-RENDER-01).
 *
 * The paged import keeps placement-only anchors. When the viewer asks, this
 * reads the sheet's drawing parts through the range reader (never the whole
 * package): the worksheet's relationships, each referenced drawing/VML part and
 * the parts those reference (charts, diagrams, media, nested relationships),
 * plus the workbook theme. Every part is size-checked from the inventory before
 * it is inflated: drawing-family and other XML parts against the object path's
 * drawing part cap, media against the drawing transfer budget, all of them
 * against one per-sheet read budget. The bounded parts then go through the
 * resident DrawingML parser unchanged, so a paged sheet paints what the
 * resident sheet paints. Anchors are placed on the geometry the paged cells
 * paint with (O(log n) per marker). A part kept out is a diagnostic.
 */
import { OfficePackage } from '../../core/zip'
import { readPartBytes, type RangeSource, type ZipInventory } from '../../core/zip-stream'
import { isPackageError } from '../../core/package-error'
import { attrs, getChildren, parseXml, parseXmlOrdered } from '../../core/xml'
import { resolvePartTarget } from '../../drawing/parts'
import { parseThemeContext, type ThemeContext } from '../../drawing/style'
import { parseWorksheetDrawings } from '../drawing'
import type { XlsxSheet } from '../types'
import type { ProfileId, SheetDrawings, WorkerDiagnostic } from '../view/contracts'
import { placementAxes } from '../view/drawingBounds'
import { DRAWING_PART_REFUSED, drawingPayloadBudget, sheetDrawingPayload } from '../view/drawingPayload'
import type { SheetGeometry } from '../../worker/viewerSession'
import { classifyPart } from './admission'
import { importLimitsFor } from './importSession'

const R_NS = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships'
/** Relationship hops followed from a worksheet drawing reference (drawing → chart → chart rels → media). */
const MAX_DEPTH = 4

export interface PagedDrawingLimits {
  /** Per XML (or other non-media) part. */
  partBytes: number
  /** Per media part. */
  mediaBytes: number
  /** All parts read for one sheet. */
  readBytes: number
  maxParts: number
  /** Binary bytes the payload may carry to the main thread. */
  payloadBytes: number
}
export function pagedDrawingLimits(profile: ProfileId, payloadBytes = drawingPayloadBudget(profile)): PagedDrawingLimits {
  const object = importLimitsFor(profile).object
  return { partBytes: object.drawingPartBytes, mediaBytes: payloadBytes, readBytes: payloadBytes + 4 * object.drawingPartBytes, maxParts: object.entries, payloadBytes }
}

export interface PagedDrawingSheet {
  name: string
  partPath: string
  drawingIds: readonly (string | undefined)[]
  legacyDrawingIds?: readonly (string | undefined)[]
}

function relsPathOf(part: string): string {
  const slash = part.lastIndexOf('/')
  return `${slash < 0 ? '' : `${part.slice(0, slash)}/`}_rels/${part.slice(slash + 1)}.rels`
}
const escapeAttr = (value: string) => value.replace(/&/g, '&amp;').replace(/</g, '&lt;').replace(/"/g, '&quot;')
const utf8 = new TextDecoder()

export async function loadPagedSheetDrawings(
  pkg: { src: RangeSource; inventory: ZipInventory },
  sheet: PagedDrawingSheet,
  geometry: SheetGeometry,
  limits: PagedDrawingLimits,
  signal: AbortSignal,
): Promise<SheetDrawings> {
  const legacy = sheet.legacyDrawingIds ?? []
  if (!sheet.drawingIds.length && !legacy.length) return sheetDrawingPayload([], undefined, [], limits.payloadBytes)
  const diagnostics: WorkerDiagnostic[] = []
  const parts = new Map<string, Uint8Array>()
  let readBytes = 0, refusedParts = 0
  const refuse = (name: string, reason: string) => {
    refusedParts++
    // Bounded: the first few parts are named, the rest are counted below.
    if (refusedParts <= 4) diagnostics.push({ code: DRAWING_PART_REFUSED, message: `Drawing part ${name} is not shown in the large-file view: ${reason}`, partId: name })
  }
  /** Reads one part within its caps; undefined when absent or kept out. */
  const read = async (name: string): Promise<Uint8Array | undefined> => {
    if (parts.has(name)) return parts.get(name)
    const entry = pkg.inventory.find(name)
    if (!entry) return undefined
    const cap = classifyPart(name).cls === 'media' ? limits.mediaBytes : limits.partBytes
    if (entry.uncompressedSize > cap) { refuse(name, `${entry.uncompressedSize} bytes exceed the ${cap}-byte part limit`); return undefined }
    if (readBytes + entry.uncompressedSize > limits.readBytes) { refuse(name, `the sheet's drawing parts exceed ${limits.readBytes} bytes`); return undefined }
    if (parts.size >= limits.maxParts) { refuse(name, `more than ${limits.maxParts} drawing parts`); return undefined }
    try {
      // The declared size was admitted above; inflation is held to it exactly.
      const bytes = await readPartBytes(pkg.src, entry, cap, { signal, budget: { charge: () => {} } })
      readBytes += bytes.byteLength
      parts.set(name, bytes)
      return bytes
    } catch (error) {
      if (signal.aborted || (isPackageError(error) && error.code === 'cancelled')) throw error
      refuse(name, error instanceof Error ? error.message : String(error))
      return undefined
    }
  }
  /** A part's internal relationships (id → resolved target present in the package). */
  const relationships = async (owner: string): Promise<Map<string, { path: string; type: string }>> => {
    const out = new Map<string, { path: string; type: string }>()
    const rels = await read(relsPathOf(owner))
    if (!rels) return out
    try {
      for (const a of getChildren(parseXml(utf8.decode(rels)), 'Relationship').map(attrs)) {
        if (!a.Id || !a.Target || a.TargetMode === 'External') continue
        const path = resolvePartTarget(owner, a.Target)
        if (pkg.inventory.find(path)) out.set(a.Id, { path, type: a.Type ?? '' })
      }
    } catch { /* an unreadable part leaves its references unresolved; the parser diagnoses them */ }
    return out
  }

  // Theme: the workbook relationship, exactly as parse.ts finds it.
  const themePath = [...(await relationships('xl/workbook.xml')).values()].find(rel => rel.type.endsWith('/theme'))?.path
  if (themePath) await read(themePath)

  // The sheet's drawing and VML parts and everything they reach, breadth first.
  // Only the worksheet's drawing references are followed (not its tables or comments).
  const queue: Array<{ path: string; depth: number }> = []
  const queued = new Set<string>()
  const enqueue = (path: string, depth: number) => { if (!queued.has(path) && depth <= MAX_DEPTH) { queued.add(path); queue.push({ path, depth }) } }
  const sheetRels = await relationships(sheet.partPath)
  for (const id of [...sheet.drawingIds, ...legacy]) {
    const rel = id === undefined ? undefined : sheetRels.get(id)
    if (rel) enqueue(rel.path, 1)
  }
  for (let next = queue.shift(); next; next = queue.shift()) {
    if (signal.aborted) throw signal.reason ?? new DOMException('Aborted', 'AbortError')
    if (!await read(next.path)) continue
    for (const rel of (await relationships(next.path)).values()) enqueue(rel.path, next.depth + 1)
  }
  if (refusedParts > 4) diagnostics.push({ code: DRAWING_PART_REFUSED, message: `${refusedParts - 4} more drawing parts are not shown in the large-file view` })
  const officePackage = OfficePackage.fromParts(parts)
  let theme: ThemeContext = parseThemeContext()
  if (themePath && parts.has(themePath)) {
    try { theme = parseThemeContext(await officePackage.xmlOrdered(themePath)) }
    catch { /* optional theme cannot invalidate usable drawings (parse.ts) */ }
  }

  // The worksheet's drawing references, as parse.ts hands them to the parser.
  const refs = [
    ...sheet.drawingIds.map(id => id === undefined ? '<drawing/>' : `<drawing r:id="${escapeAttr(id)}"/>`),
    ...legacy.map(id => id === undefined ? '<legacyDrawing/>' : `<legacyDrawing r:id="${escapeAttr(id)}"/>`),
  ]
  const root = parseXmlOrdered(`<worksheet xmlns:r="${R_NS}">${refs.join('')}</worksheet>`)
  const worksheet = getChildren(root, 'worksheet')[0] ?? root
  const shell: XlsxSheet = { name: sheet.name, sourcePartPath: sheet.partPath, rows: [], cols: [], merges: [], mergeRanges: [] }
  const axes = placementAxes(geometry)
  await parseWorksheetDrawings(officePackage, shell, sheet.partPath, worksheet, theme, { col: index => axes.cols.start(index), row: index => axes.rows.start(index) })
  if (signal.aborted) throw signal.reason ?? new DOMException('Aborted', 'AbortError')
  return sheetDrawingPayload(shell.drawings ?? [], theme, diagnostics, limits.payloadBytes)
}
