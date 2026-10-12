/**
 * Drawing content of one paged sheet (S13.2; spec RENDER-02, AC-RENDER-01).
 *
 * The paged import keeps placement-only anchors. When the viewer asks, the
 * resident DrawingML parser runs unchanged over a package view that reads parts
 * on demand through the range reader (never the whole package, and only the
 * parts the parser actually opens: the worksheet relationships, the drawing
 * and VML parts it references, and the charts, diagrams, media and nested
 * relationships those reference, plus the workbook theme). Every part is
 * admitted from its inventory size before it is inflated: media against the
 * drawing transfer budget, any other part against the object path's drawing
 * part cap, all of them together against one per-sheet read budget. A part kept
 * out reads as absent to the parser (as a missing part does on the resident
 * path) and is named in a diagnostic. Anchors are placed on the geometry the
 * paged cells paint with (O(log n) per marker), so a paged sheet paints what
 * the resident sheet paints.
 */
import JSZip from 'jszip'
import { OfficePackage } from '../../core/zip'
import { readPartBytes, type RangeSource, type ZipInventory } from '../../core/zip-stream'
import { isPackageError } from '../../core/package-error'
import { attrs, getChildren, parseXml, parseXmlOrdered, type XmlNode } from '../../core/xml'
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
/** Refused parts named one by one in diagnostics; the rest are counted. */
const NAMED_REFUSALS = 4

export interface PagedDrawingLimits {
  /** Per non-media part (drawing XML, charts, relationships, VML, ...). */
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

const utf8 = new TextDecoder()
/** OfficePackage over the range reader: bounded on-demand part reads, cached per part. */
class RangePackage extends OfficePackage {
  private readonly reads = new Map<string, Promise<Uint8Array | undefined>>()
  private readonly trees = new Map<string, Promise<XmlNode | undefined>>()
  private readonly ordered = new Map<string, Promise<XmlNode | undefined>>()
  private readBytes = 0
  private partsRead = 0
  private refusals = 0
  readonly diagnostics: WorkerDiagnostic[] = []

  constructor(private readonly src: RangeSource, private readonly inventory: ZipInventory, private readonly limits: PagedDrawingLimits, private readonly signal: AbortSignal) {
    super(new JSZip())
  }
  private cap(path: string): number { return classifyPart(path).cls === 'media' ? this.limits.mediaBytes : this.limits.partBytes }
  private refuse(path: string, reason: string): void {
    if (++this.refusals <= NAMED_REFUSALS) this.diagnostics.push({ code: DRAWING_PART_REFUSED, message: `Drawing part ${path} is not shown in the large-file view: ${reason}`, partId: path })
  }
  finish(): WorkerDiagnostic[] {
    if (this.refusals > NAMED_REFUSALS) this.diagnostics.push({ code: DRAWING_PART_REFUSED, message: `${this.refusals - NAMED_REFUSALS} more drawing parts are not shown in the large-file view` })
    return this.diagnostics
  }
  /** Present and within its own cap (a part beyond it reads as absent). */
  override has(path: string): boolean {
    const entry = this.inventory.find(path)
    if (!entry) return false
    if (entry.uncompressedSize <= this.cap(path)) return true
    if (!this.reads.has(path)) {
      this.reads.set(path, Promise.resolve(undefined))
      this.refuse(path, `${entry.uncompressedSize} bytes exceed the ${this.cap(path)}-byte part limit`)
    }
    return false
  }
  override bytes(path: string, maxBytes?: number): Promise<Uint8Array | undefined> {
    let pending = this.reads.get(path)
    if (!pending) {
      pending = this.read(path)
      this.reads.set(path, pending)
    }
    return maxBytes === undefined ? pending : pending.then(bytes => {
      if (bytes && bytes.byteLength > maxBytes) throw new Error('Part byte limit exceeded')
      return bytes
    })
  }
  override async text(path: string): Promise<string | undefined> {
    const bytes = await this.bytes(path)
    return bytes && utf8.decode(bytes)
  }
  override xml(path: string): Promise<XmlNode | undefined> {
    let tree = this.trees.get(path)
    if (!tree) this.trees.set(path, tree = this.text(path).then(text => text === undefined ? undefined : parseXml(text)))
    return tree
  }
  override xmlOrdered(path: string): Promise<XmlNode | undefined> {
    let tree = this.ordered.get(path)
    if (!tree) this.ordered.set(path, tree = this.text(path).then(text => text === undefined ? undefined : parseXmlOrdered(text)))
    return tree
  }
  override paths(predicate?: (p: string) => boolean): string[] {
    return this.inventory.entries.map(entry => entry.name).filter(name => !predicate || predicate(name))
  }
  private async read(path: string): Promise<Uint8Array | undefined> {
    if (!this.has(path)) return undefined
    const entry = this.inventory.find(path)!
    if (this.readBytes + entry.uncompressedSize > this.limits.readBytes) { this.refuse(path, `the sheet's drawing parts exceed ${this.limits.readBytes} bytes`); return undefined }
    if (this.partsRead >= this.limits.maxParts) { this.refuse(path, `more than ${this.limits.maxParts} drawing parts`); return undefined }
    // Charged before the read, so concurrent reads cannot overrun the budget.
    this.readBytes += entry.uncompressedSize
    this.partsRead++
    try {
      // The declared size was admitted; inflation is held to it exactly.
      return await readPartBytes(this.src, entry, this.cap(path), { signal: this.signal, budget: { charge: () => {} } })
    } catch (error) {
      if (this.signal.aborted || (isPackageError(error) && error.code === 'cancelled')) throw error
      this.refuse(path, error instanceof Error ? error.message : String(error))
      return undefined
    }
  }
}

const escapeAttr = (value: string) => value.replace(/&/g, '&amp;').replace(/</g, '&lt;').replace(/"/g, '&quot;')
const aborted = (signal: AbortSignal) => signal.reason ?? new DOMException('Aborted', 'AbortError')

export async function loadPagedSheetDrawings(
  source: { src: RangeSource; inventory: ZipInventory },
  sheet: PagedDrawingSheet,
  geometry: SheetGeometry,
  limits: PagedDrawingLimits,
  signal: AbortSignal,
): Promise<SheetDrawings> {
  const legacy = sheet.legacyDrawingIds ?? []
  if (!sheet.drawingIds.length && !legacy.length) return sheetDrawingPayload([], undefined, [], limits.payloadBytes)
  const pkg = new RangePackage(source.src, source.inventory, limits, signal)
  // Theme: the workbook relationship, exactly as parse.ts finds it.
  let theme: ThemeContext = parseThemeContext()
  try {
    const rels = getChildren(await pkg.xml('xl/_rels/workbook.xml.rels'), 'Relationship').map(attrs)
    const rel = rels.find(a => a.Type?.endsWith('/theme') && a.TargetMode !== 'External' && a.Target)
    if (rel) theme = parseThemeContext(await pkg.xmlOrdered(resolvePartTarget('xl/workbook.xml', rel.Target!)))
  } catch {
    // An optional theme cannot invalidate usable drawings (parse.ts); a cancel still ends the read.
    if (signal.aborted) throw aborted(signal)
  }
  // The worksheet's drawing references, handed to the parser as parse.ts hands them.
  const refs = [
    ...sheet.drawingIds.map(id => id === undefined ? '<drawing/>' : `<drawing r:id="${escapeAttr(id)}"/>`),
    ...legacy.map(id => id === undefined ? '<legacyDrawing/>' : `<legacyDrawing r:id="${escapeAttr(id)}"/>`),
  ]
  const root = parseXmlOrdered(`<worksheet xmlns:r="${R_NS}">${refs.join('')}</worksheet>`)
  const worksheet = getChildren(root, 'worksheet')[0] ?? root
  const shell: XlsxSheet = { name: sheet.name, sourcePartPath: sheet.partPath, rows: [], cols: [], merges: [], mergeRanges: [] }
  const axes = placementAxes(geometry)
  await parseWorksheetDrawings(pkg, shell, sheet.partPath, worksheet, theme, { col: index => axes.cols.start(index), row: index => axes.rows.start(index) })
  if (signal.aborted) throw aborted(signal)
  return sheetDrawingPayload(shell.drawings ?? [], theme, pkg.finish(), limits.payloadBytes)
}
