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
import { DOCUMENT_DRAWING_NODE_LIMIT, drawingPartContext, resolvePartTarget } from '../../drawing/parts'
import { parseThemeContext, type ThemeContext } from '../../drawing/style'
import { parseWorksheetDrawings } from '../drawing'
import type { XlsxSheet } from '../types'
import type { DrawingMemory, ProfileId, SheetDrawings, WorkerDiagnostic } from '../view/contracts'
import { placementAxes } from '../view/drawingBounds'
import { DRAWING_MEMORY_REFUSED, DRAWING_NODE_BUDGET, DRAWING_PART_REFUSED, drawingFootprint, drawingPayloadBudget, sheetDrawingPayload, unavailableDrawings } from '../view/drawingPayload'
import type { SheetGeometry } from '../../worker/viewerSession'
import { classifyPart } from './admission'
import { importLimitsFor } from './importSession'

const R_NS = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships'
/** Refused parts named one by one in diagnostics; the rest are counted. */
const NAMED_REFUSALS = 4
/** Retained bytes per element of a parsed (ordered or plain) XML tree, attributes
 * included. Calibrated from heap measurements of parseXmlOrdered over drawing
 * parts in JavaScriptCore (380-560 B per element at 60k-120k elements), rounded
 * up; charged per tree actually built, on top of the part's bytes and text. */
export const XML_ELEMENT_BYTES = 640
/** Element start tags in an XML byte stream (`<` not opening an end tag, PI, comment or doctype). */
export function countXmlElements(bytes: Uint8Array): number {
  let count = 0
  for (let i = 0; i < bytes.length; i++) {
    if (bytes[i] !== 0x3c) continue
    const next = bytes[i + 1]
    if (next !== 0x2f && next !== 0x3f && next !== 0x21) count++
  }
  return count
}


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
  /** Charged working-set estimate (inflated parts, their text, parsed trees). */
  workingSet = 0
  /** A part did not fit the document pool: the whole read is refused, never partial. */
  memoryRefused = false
  /** Drawing nodes the parse had reserved when the pool refused. */
  nodesAtRefusal = 0
  private readonly texts = new Map<string, Promise<string | undefined>>()
  private settleRefusal: () => void = () => {}
  /** Settles when the pool refuses, so the caller stops waiting on the parse. */
  readonly refusal = new Promise<void>(resolve => { this.settleRefusal = resolve })
  readonly diagnostics: WorkerDiagnostic[] = []

  constructor(private readonly src: RangeSource, private readonly inventory: ZipInventory, private readonly limits: PagedDrawingLimits, private readonly signal: AbortSignal, private readonly memory?: DrawingMemory) {
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
  override text(path: string): Promise<string | undefined> {
    let text = this.texts.get(path)
    if (!text) this.texts.set(path, text = this.bytes(path).then(bytes => bytes && utf8.decode(bytes)))
    return text
  }
  override xml(path: string): Promise<XmlNode | undefined> {
    let tree = this.trees.get(path)
    if (!tree) this.trees.set(path, tree = this.tree(path, parseXml))
    return tree
  }
  override xmlOrdered(path: string): Promise<XmlNode | undefined> {
    let tree = this.ordered.get(path)
    if (!tree) this.ordered.set(path, tree = this.tree(path, parseXmlOrdered))
    return tree
  }
  /** Every tree built is admitted first, by its element count. */
  private async tree(path: string, parse: (text: string) => XmlNode): Promise<XmlNode | undefined> {
    const bytes = await this.bytes(path)
    if (!bytes || !this.admit(countXmlElements(bytes) * XML_ELEMENT_BYTES)) return undefined
    const text = await this.text(path)
    return text === undefined ? undefined : parse(text)
  }
  /** Grows the pool share by `cost`; on refusal the read is refused as a whole and
   * the parse winds down: no further part reads and no further drawing nodes. */
  private admit(cost: number): boolean {
    if (this.memoryRefused) return false
    if (!this.memory || this.memory.set(this.workingSet + cost)) { this.workingSet += cost; return true }
    this.memoryRefused = true
    const context = drawingPartContext(this)
    this.nodesAtRefusal = context.nodes
    context.nodes = DOCUMENT_DRAWING_NODE_LIMIT
    this.settleRefusal()
    return false
  }
  override paths(predicate?: (p: string) => boolean): string[] {
    return this.inventory.entries.map(entry => entry.name).filter(name => !predicate || predicate(name))
  }
  private async read(path: string): Promise<Uint8Array | undefined> {
    if (this.memoryRefused || !this.has(path)) return undefined
    const entry = this.inventory.find(path)!
    if (this.readBytes + entry.uncompressedSize > this.limits.readBytes) { this.refuse(path, `the sheet's drawing parts exceed ${this.limits.readBytes} bytes`); return undefined }
    if (this.partsRead >= this.limits.maxParts) { this.refuse(path, `more than ${this.limits.maxParts} drawing parts`); return undefined }
    // The pool admits the inflated bytes (and, for XML, their UTF-16 text) before anything is inflated.
    if (!this.admit(entry.uncompressedSize * (classifyPart(path).cls === 'media' ? 1 : 3))) return undefined
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
  options: { memory?: DrawingMemory; startNodes?: number } = {},
): Promise<PagedSheetDrawings> {
  const legacy = sheet.legacyDrawingIds ?? []
  if (!sheet.drawingIds.length && !legacy.length) return { drawings: sheetDrawingPayload([], undefined, [], limits.payloadBytes), nodes: 0, workingSet: 0 }
  const pkg = new RangePackage(source.src, source.inventory, limits, signal, options.memory)
  const context = drawingPartContext(pkg)
  const startNodes = context.nodes = options.startNodes ?? 0
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
  const parsed = parseWorksheetDrawings(pkg, shell, sheet.partPath, worksheet, theme, { col: index => axes.cols.start(index), row: index => axes.rows.start(index) })
  // A pool refusal ends the read at once; the parse it leaves behind reads no parts and reserves no nodes.
  parsed.catch(() => {})
  await Promise.race([parsed, pkg.refusal])
  if (signal.aborted) throw aborted(signal)
  const nodes = (pkg.memoryRefused ? pkg.nodesAtRefusal : context.nodes) - startNodes
  if (pkg.memoryRefused) {
    return { drawings: unavailableDrawings('This sheet\'s drawings need more worker memory than the document has free and are not shown', { code: 'resource-limit', retryable: false, message: 'worker memory budget' }, DRAWING_MEMORY_REFUSED), nodes, workingSet: pkg.workingSet }
  }
  const diagnostics = pkg.finish()
  if (context.diagnostics.some(issue => issue.kind === 'node-budget' && issue.reason === 'document-budget')) {
    diagnostics.push({ code: DRAWING_NODE_BUDGET, message: `Some drawings exceed the document's ${DOCUMENT_DRAWING_NODE_LIMIT}-node drawing budget and are not shown` })
  }
  const drawings = sheetDrawingPayload(shell.drawings ?? [], theme, diagnostics, limits.payloadBytes)
  const workingSet = pkg.workingSet + drawingFootprint(drawings.drawings)
  if (options.memory && !options.memory.set(workingSet)) {
    return { drawings: unavailableDrawings('This sheet\'s drawings need more worker memory than the document has free and are not shown', { code: 'resource-limit', retryable: false, message: 'worker memory budget' }, DRAWING_MEMORY_REFUSED), nodes, workingSet: pkg.workingSet }
  }
  return { drawings, nodes, workingSet }
}
export interface PagedSheetDrawings {
  drawings: SheetDrawings
  /** Drawing nodes this parse used (for the document ledger). */
  nodes: number
  /** Charged working-set estimate at the end of the parse. */
  workingSet: number
}
