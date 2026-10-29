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
import { parseWorksheetDrawings, settleChartFallbacks } from '../drawing'
import type { XlsxDrawing, XlsxSheet } from '../types'
import { hasXlsxCharts, loadChartModule } from '../chartLoader'
import { parseDefinedNames } from '../formula/names'
import type { DefinedNameMetadata } from '../formula/types'
import type { DrawingMemory, ProfileId, RemotePivotDisplay, SheetDrawings, WorkerDiagnostic } from '../view/contracts'
import { placementAxes } from '../view/drawingBounds'
import { DRAWING_MEMORY_REFUSED, DRAWING_NODE_BUDGET, DRAWING_PART_REFUSED, drawingFootprint, drawingPayloadBudget, sheetDrawingPayload, unavailableDrawings } from '../view/drawingPayload'
import type { SheetGeometry } from '../../worker/viewerSession'
import { classifyPart } from './admission'
import { importLimitsFor } from './importSession'
import type { PivotRegionRef } from '../pivots/stub'
import type { PivotTableDefinition } from '../pivots/model'
import type { TableStyleSource } from '../tableStyles/model'
import { resolveSheetControls } from '../pivotLoader'
import { PIVOT_PART_MAX_BYTES } from '../pivots/stub'
import type { SheetControls } from '../pivots/slicers/model'

const R_NS = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships'
/** Refused parts named one by one in diagnostics; the rest are counted. */
const NAMED_REFUSALS = 4
/** Retained bytes per element of a parsed (ordered or plain) XML tree, attributes
 * included. Calibrated from heap measurements of parseXmlOrdered over drawing
 * parts in JavaScriptCore (380-560 B per element at 60k-120k elements), rounded
 * up; charged per tree actually built, on top of the part's bytes and text. */
export const XML_ELEMENT_BYTES = 640
/** Element start tags in an XML byte stream. Skips end tags, processing
 * instructions, comments, CDATA sections and doctype declarations (their
 * content is never counted), and `<` inside quoted attribute values. */
export function countXmlElements(bytes: Uint8Array): number {
  const LT = 0x3c, GT = 0x3e, SLASH = 0x2f, QUESTION = 0x3f, BANG = 0x21, DQ = 0x22, SQ = 0x27, LSQB = 0x5b, RSQB = 0x5d
  const startsWith = (at: number, text: string) => { for (let k = 0; k < text.length; k++) if (bytes[at + k] !== text.charCodeAt(k)) return false; return true }
  /** Index just past `terminator` from `from`, or the end. */
  const skipTo = (from: number, terminator: string) => {
    for (let i = from; i < bytes.length; i++) if (startsWith(i, terminator)) return i + terminator.length
    return bytes.length
  }
  let count = 0, i = 0
  while (i < bytes.length) {
    if (bytes[i] !== LT) { i++; continue }
    const next = bytes[i + 1]
    if (next === BANG) {
      if (startsWith(i, '<!--')) i = skipTo(i + 4, '-->')
      else if (startsWith(i, '<![CDATA[')) i = skipTo(i + 9, ']]>')
      else {
        // A declaration (doctype); an internal subset [ ... ] may hold `>`.
        let depth = 0, quote = 0
        for (i += 2; i < bytes.length; i++) {
          const c = bytes[i]
          if (quote) { if (c === quote) quote = 0 }
          else if (c === DQ || c === SQ) quote = c
          else if (c === LSQB) depth++
          else if (c === RSQB) depth--
          else if (c === GT && depth <= 0) { i++; break }
        }
      }
      continue
    }
    if (next === QUESTION) { i = skipTo(i + 2, '?>'); continue }
    if (next !== SLASH) count++
    // Skip the rest of the tag; quoted attribute values may contain `<` or `>`.
    let quote = 0
    for (i++; i < bytes.length; i++) {
      const c = bytes[i]
      if (quote) { if (c === quote) quote = 0 }
      else if (c === DQ || c === SQ) quote = c
      else if (c === GT) { i++; break }
    }
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
  /** Independent cap for pivot, slicer and timeline definition parts (Ruling AL). */
  pivotPartBytes: number
}
export function pagedDrawingLimits(profile: ProfileId, payloadBytes = drawingPayloadBudget(profile)): PagedDrawingLimits {
  const object = importLimitsFor(profile).object
  return { partBytes: object.drawingPartBytes, mediaBytes: payloadBytes, readBytes: payloadBytes + 4 * object.drawingPartBytes, maxParts: object.entries, payloadBytes, pivotPartBytes: PIVOT_PART_MAX_BYTES }
}

export interface PagedDrawingSheet {
  name: string
  partPath: string
  drawingIds: readonly (string | undefined)[]
  legacyDrawingIds?: readonly (string | undefined)[]
  slicerIds?: readonly string[]
  timelineIds?: readonly string[]
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
  private pivotRefusals = 0
  /** Charged working-set estimate (inflated parts, their text, parsed trees). */
  workingSet = 0
  /** A chart fill's live scratch (its cell maps and typed arrays): part of every admission until the fill ends. */
  private fillScratch = 0
  /** A part did not fit the document pool: the whole read is refused, never partial. */
  memoryRefused = false
  /** Drawing nodes the parse had reserved when the pool refused. */
  nodesAtRefusal = 0
  private readonly texts = new Map<string, Promise<string | undefined>>()
  private settleRefusal: () => void = () => {}
  /** Settles when the pool refuses, so the caller stops waiting on the parse. */
  readonly refusal = new Promise<void>(resolve => { this.settleRefusal = resolve })
  readonly diagnostics: WorkerDiagnostic[] = []
  /** The parent ledger of this read's chart workspace (charts/context `chartParseContextFor`; D0): the
   * document pool share this read is charged to. A chart allocation is counted once here, through the
   * workspace, and refunded when the workspace releases it, so the share returns to its baseline. */
  readonly chartLedger = {
    reserve: (bytes: number, tag: string) => transientChartTag(tag) || this.admit(bytes),
    release: (tag: string, bytes: number) => { if (!transientChartTag(tag)) this.refund(bytes) },
  }
  /** The chart workspace's profile (desktop / mobile limits). */
  readonly chartProfile: 'desktop' | 'mobile'

  constructor(private readonly src: RangeSource, private readonly inventory: ZipInventory, private readonly limits: PagedDrawingLimits, private readonly signal: AbortSignal, private readonly memory?: DrawingMemory, profile: ProfileId = 'D', readonly chartDocument?: { workspace?: unknown }) {
    super(new JSZip())
    this.chartProfile = profile === 'M' ? 'mobile' : 'desktop'
  }
  private isPivotPart(path: string): boolean { return /\/(?:pivotTables|pivotCache|slicerCaches|slicers|timelineCaches|timelines)\//i.test(`/${path}`) }
  private cap(path: string): number { return this.isPivotPart(path) ? this.limits.pivotPartBytes : classifyPart(path).cls === 'media' ? this.limits.mediaBytes : this.limits.partBytes }
  private refuse(path: string, reason: string): void {
    const pivot = this.isPivotPart(path)
    if (pivot) {
      if (++this.pivotRefusals <= NAMED_REFUSALS) this.diagnostics.push({ code: 'pivot-part-too-large', message: `Pivot/control part ${path} was not read: ${reason}`, partId: path })
    }
    else if (++this.refusals <= NAMED_REFUSALS) this.diagnostics.push({ code: DRAWING_PART_REFUSED, message: `Drawing part ${path} is not shown in the large-file view: ${reason}`, partId: path })
  }
  finish(): WorkerDiagnostic[] {
    if (this.refusals > NAMED_REFUSALS) this.diagnostics.push({ code: DRAWING_PART_REFUSED, message: `${this.refusals - NAMED_REFUSALS} more drawing parts are not shown in the large-file view` })
    if (this.pivotRefusals > NAMED_REFUSALS) this.diagnostics.push({ code: 'pivot-part-too-large', message: `${this.pivotRefusals - NAMED_REFUSALS} additional pivot/control parts exceed their part cap` })
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
  /** Shrinks the pool share by `cost` (a released chart allocation). */
  private refund(cost: number): void {
    this.workingSet = Math.max(0, this.workingSet - cost)
    if (!this.memoryRefused) this.memory?.set(this.workingSet + this.fillScratch)
  }
  /** Sets the chart fill's live scratch to `bytes` (its whole current total, not an increment); a refusal keeps
   * the previous reservation. 0 releases it. */
  reserveFillScratch(bytes: number): boolean {
    if (this.memory && !this.memory.set(this.workingSet + bytes)) return false
    this.fillScratch = bytes
    return true
  }
  endFill(): void { this.fillScratch = 0 }
  /** Grows the pool share by `cost`; on refusal the read is refused as a whole and
   * the parse winds down: no further part reads and no further drawing nodes. */
  private admit(cost: number): boolean {
    if (this.memoryRefused) return false
    if (!this.memory || this.memory.set(this.workingSet + this.fillScratch + cost)) { this.workingSet += cost; return true }
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

/** Resolve one sheet's clone-safe pivot/control descriptors through the same bounded package reader as drawings. */
export async function readPagedPivotDisplayParts(source: { src: RangeSource; inventory: ZipInventory }, input: {
  sheetId: string; sheetName: string; sheetPath: string; pivotRegions: ReadonlyMap<string, readonly PivotRegionRef[]>
  slicerIds: readonly string[]; timelineIds: readonly string[]; theme?: ThemeContext; styles?: TableStyleSource
}, profile: ProfileId, signal: AbortSignal, memory?: DrawingMemory): Promise<Pick<RemotePivotDisplay, 'pivotTables' | 'slicers' | 'timelines' | 'diagnostics'>> {
  const limits = pagedDrawingLimits(profile)
  const pkg = new RangePackage(source.src, source.inventory, limits, signal, memory, profile)
  const diagnostics: RemotePivotDisplay['diagnostics'][number][] = []
  let pivotTables: PivotTableDefinition[] = [], slicers: RemotePivotDisplay['slicers'][number][] = [], timelines: RemotePivotDisplay['timelines'][number][] = []
  try {
    if (input.pivotRegions.size) {
      const parsed = await (await import('../pivots/load')).loadPivotWorkbook(pkg, input.pivotRegions)
      pivotTables = parsed.tables.filter(table => table.sheet === input.sheetName)
      diagnostics.push(...parsed.diagnostics)
    }
    if (input.slicerIds.length || input.timelineIds.length) {
      const controls = await resolveSheetControls(pkg, { sheetPath: input.sheetPath, slicerIds: input.slicerIds, timelineIds: input.timelineIds, theme: input.theme ?? parseThemeContext(), ...(input.styles ? { styles: input.styles } : {}) })
      if (controls) {
        slicers = [...controls.slicers.values()]
        timelines = [...controls.timelines.values()]
        diagnostics.push(...controls.diagnostics)
      } else diagnostics.push({ kind: 'unsupported', feature: 'slicer-engine-unavailable', message: 'Slicers and timelines could not be loaded; their fallback shapes are shown' })
    }
  } catch (cause) {
    if (signal.aborted) throw aborted(signal)
    diagnostics.push({ kind: 'unsupported', feature: 'pivot-display-read-failed', message: cause instanceof Error ? cause.message.slice(0, 256) : 'Pivot display metadata could not be read' })
  }
  diagnostics.push(...pkg.finish().map(item => ({ kind: 'unsupported' as const, feature: item.code, message: item.message, partPath: item.partId })))
  return { pivotTables, slicers, timelines, diagnostics }
}

/**
 * A chart read's transient decode work (charts/read.ts: the crossing reservation, one chunk's text) on a
 * part this package already admitted with its full UTF-16 text (`read`: 3x the inflated size): those bytes
 * are already in the pool, so the workspace counts them toward its sublimit only. Tree nodes are new
 * allocations and are charged. (Suffixes as CHART_CROSSING_TAG / CHART_CHUNK_TAG; this module stays
 * independent of the chart chunk.)
 */
const transientChartTag = (tag: string) => tag.endsWith(':crossing') || tag.endsWith(':chunk')
const escapeAttr = (value: string) => value.replace(/&/g, '&amp;').replace(/</g, '&lt;').replace(/"/g, '&quot;')
const aborted = (signal: AbortSignal) => signal.reason ?? new DOMException('Aborted', 'AbortError')

/** A paged read's scoped chart workspace (charts/workspace `scopeChartWorkspace`), as this module sees it. */
export interface PagedChartScope { readonly shareBytes: number; handOff(): void; releaseAll(): void }
export interface PagedDrawingOptions {
  memory?: DrawingMemory
  startNodes?: number
  profile?: ProfileId
  /** The document's chart facet holder (one per paged document; created by the chart engine on first use). */
  chartDocument?: { workspace?: unknown }
  fillCharts?: (drawings: XlsxDrawing[], ownerSheet: string, names: DefinedNameMetadata[], admit: (bytes: number) => boolean, workspace: unknown) => Promise<{ pendingSheets: string[]; uncacheable?: boolean; refused?: boolean; diagnostics?: WorkerDiagnostic[] }>
}
export async function loadPagedSheetDrawings(
  source: { src: RangeSource; inventory: ZipInventory },
  sheet: PagedDrawingSheet,
  geometry: SheetGeometry,
  limits: PagedDrawingLimits,
  signal: AbortSignal,
  options: PagedDrawingOptions = {},
): Promise<PagedSheetDrawings> {
  const legacy = sheet.legacyDrawingIds ?? []
  if (!sheet.drawingIds.length && !legacy.length) return { drawings: sheetDrawingPayload([], undefined, [], limits.payloadBytes), nodes: 0, workingSet: 0 }
  const pkg = new RangePackage(source.src, source.inventory, limits, signal, options.memory, options.profile, options.chartDocument)
  const scope = () => drawingPartContext(pkg).chart?.workspace as unknown as PagedChartScope | undefined
  // A read that fails, is cancelled or is refused refunds every chart charge it holds (document facet and share).
  try {
    const result = await readPagedSheetDrawings(pkg, sheet, geometry, limits, signal, options)
    const chartScope = scope()
    if (result.drawings.completeness !== 'complete' && !result.payloadBytes) { chartScope?.releaseAll(); return result }
    return chartScope ? { ...result, chartScope } : result
  } catch (error) {
    scope()?.releaseAll()
    throw error
  }
}
async function readPagedSheetDrawings(pkg: RangePackage, sheet: PagedDrawingSheet, geometry: SheetGeometry, limits: PagedDrawingLimits, signal: AbortSignal, options: PagedDrawingOptions): Promise<PagedSheetDrawings> {
  const legacy = sheet.legacyDrawingIds ?? []
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
  let controls: SheetControls | undefined
  if (sheet.slicerIds?.length || sheet.timelineIds?.length) {
    controls = await resolveSheetControls(pkg, { sheetPath: sheet.partPath, slicerIds: sheet.slicerIds ?? [], timelineIds: sheet.timelineIds ?? [], theme })
  }
  const parsed = parseWorksheetDrawings(pkg, shell, sheet.partPath, worksheet, theme, { col: index => axes.cols.start(index), row: index => axes.rows.start(index) }, controls)
  // A pool refusal ends the read at once; the parse it leaves behind reads no parts and reserves no nodes.
  parsed.catch(() => {})
  await Promise.race([parsed, pkg.refusal])
  if (signal.aborted) throw aborted(signal)
  const nodes = (pkg.memoryRefused ? pkg.nodesAtRefusal : context.nodes) - startNodes
  if (pkg.memoryRefused) {
    return { drawings: unavailableDrawings('This sheet\'s drawings need more worker memory than the document has free and are not shown', { code: 'resource-limit', retryable: false, message: 'worker memory budget' }, DRAWING_MEMORY_REFUSED), nodes, workingSet: pkg.workingSet }
  }
  let pendingSheets: string[] | undefined, uncacheable = false
  let chartDiagnostics: WorkerDiagnostic[] = []
  // Charts without cached values read their series from the paged cells.
  if (options.fillCharts && hasXlsxCharts(shell.drawings)) {
    // Defined names come from the workbook part, read only when a chart names one.
    let names: DefinedNameMetadata[] = []
    // A failed engine load leaves the charts as cards (spec 6.3); the drawings still load.
    if ((await loadChartModule().catch(() => undefined))?.drawingsNeedNames(shell.drawings)) {
      try { names = parseDefinedNames(await pkg.xml('xl/workbook.xml')) } catch { if (signal.aborted) throw aborted(signal) }
    }
    // The chart cell workspace is admitted against the drawing memory before any read.
    // The fill's scratch stays in every chart-ledger admission (model storage built during the fill) until it ends.
    let filled: Awaited<ReturnType<NonNullable<PagedDrawingOptions['fillCharts']>>>
    try { filled = await options.fillCharts(shell.drawings!, sheet.name, names, bytes => pkg.reserveFillScratch(bytes), context.chart?.workspace) } finally { pkg.endFill() }
    if (signal.aborted) throw aborted(signal)
    if (filled.pendingSheets.length) pendingSheets = filled.pendingSheets
    if (filled.refused) {
      return { drawings: unavailableDrawings('This sheet\'s drawings need more worker memory than the document has free and are not shown', { code: 'resource-limit', retryable: false, message: 'worker memory budget' }, DRAWING_MEMORY_REFUSED), nodes: context.nodes - startNodes, workingSet: pkg.workingSet }
    }
    if (filled.uncacheable) uncacheable = true
    chartDiagnostics = filled.diagnostics ?? []
    // A chart Choice refused once its data was read gives way to its fallback (spec 6.3).
    await settleChartFallbacks(shell)
    if (signal.aborted) throw aborted(signal)
  }
  const diagnostics = [...pkg.finish(), ...chartDiagnostics]
  if (context.diagnostics.some(issue => issue.kind === 'node-budget' && issue.reason === 'document-budget')) {
    diagnostics.push({ code: DRAWING_NODE_BUDGET, message: `Some drawings exceed the document's ${DOCUMENT_DRAWING_NODE_LIMIT}-node drawing budget and are not shown` })
  }
  const payload = sheetDrawingPayload(shell.drawings ?? [], theme, diagnostics, limits.payloadBytes)
  const drawings = pendingSheets && payload.completeness === 'complete' ? { ...payload, pendingSheets } : payload
  // Chart model storage is already in the share (charged through the read's chart scope before it was built):
  // the payload adds only the rest of its footprint, so each byte counts once.
  const modelBytes = (context.chart?.workspace as unknown as PagedChartScope | undefined)?.shareBytes ?? 0
  const payloadBytes = Math.max(0, drawingFootprint(drawings.drawings) - modelBytes)
  const workingSet = pkg.workingSet + payloadBytes
  if (options.memory && !options.memory.set(workingSet)) {
    return { drawings: unavailableDrawings('This sheet\'s drawings need more worker memory than the document has free and are not shown', { code: 'resource-limit', retryable: false, message: 'worker memory budget' }, DRAWING_MEMORY_REFUSED), nodes, workingSet: pkg.workingSet, payloadBytes: 0 }
  }
  return { drawings, nodes, workingSet, payloadBytes, modelBytes, ...(uncacheable ? { uncacheable: true } : {}) }
}
export interface PagedSheetDrawings {
  drawings: SheetDrawings
  /** Drawing nodes this parse used (for the document ledger). */
  nodes: number
  /** Charged working-set estimate at the end of the parse (the payload included). */
  workingSet: number
  /** The part of `workingSet` that is the payload itself, chart model storage excepted. */
  payloadBytes?: number
  /** The part of `workingSet` that is chart model storage (charged through the read's chart scope). */
  modelBytes?: number
  /** The read's chart scope: the caller hands it off once the payload has its owner (retained entry or reply). */
  chartScope?: PagedChartScope
  /** A chart read a failed sheet: answer, but never cache this payload. */
  uncacheable?: boolean
}
