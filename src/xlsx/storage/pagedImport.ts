/**
 * Paged importer (S1.2, design note §4.4). Kept out of importSession.ts so the
 * eager bundle (which imports the limits and ImportRefusal from there) never
 * pulls in the ZIP/XML readers: only the lazy large-source chunk loads this.
 */
import type { WorkbookFormatting } from '../view/contracts'
import { DEFAULT_FORMATTING, workbookFormattingWithColorContext } from './pagedFormatting'
import type { Completeness, ProfileId, SourceFailure, SourceHandle, SourceStamp, WorkerDiagnostic } from '../view/contracts'
import { PackageError, isPackageError, throwIfAborted } from '../../core/package-error'
import { blobRangeSource, openPart, readPartBytes, readPartPrefixBytes, readZipInventory, type RangeSource, type ZipInventory } from '../../core/zip-stream'
import { tokenizeBytes } from '../../core/xml-stream'
import { attrs, getChildren, parseXml } from '../../core/xml'
import { resolvePartTarget } from '../../drawing/parts'
import { dynamicMetadataIndices } from '../metadata'
import { isTablePartSource, parseTableMetadata } from '../formula/tables'
import type { TableMetadata } from '../formula/types'
import type { XlsxDrawingAnchors } from '../types'
import { PIVOT_PART_MAX_BYTES, parsePivotRegion, type ControlPartRef } from '../pivots/stub'
import { anchorId, anchorRecord } from '../anchorRecords'
import type { SheetColorContext } from '../sheetColors'
import type { ThemeContext } from '../../drawing/style'
import type { TableStyleDef } from '../tableStyles/model'
import type { StylesExtras } from '../tableStyles/model'
import {
  collectDeclaredTable, createDeclaredBudget, declaredContainerBuilder, declaredSheetBuilder, hasDeclaredFiltering, type DeclaredSheetFilters,
} from '../view/declaredFilters'
import { MAX_COLS, MAX_ROWS, parseA1, type SourceRect } from '../view/coordinates'
import { isLoadEvaluated } from '../view/values'
import { createQuotaLedger, quotaLimitsFor, type QuotaLedger } from './quota'
import { WORKER_BUDGETS } from '../view/tileCache'
import { classifyPart, inventoryLimits } from './admission'
import { anchorRecordEvents, newSheetMetadata, normalizeRowMetadata, sharedStringEvents, sheetRecordEvents, type SheetMetadata } from './sheetRecords'
import { PageBuilder, SourcePageStore, browserPageStorageEnv, choosePageBackend, pageOwner, payloadBytes, type PageStorageEnv, type StoredPage } from './sourcePages'
import { PENDING_STRING_BYTES, TEXT_STRING_BYTES } from './constructionMemory'
import { importLimitsFor, type ImportLimits, type ImportLimitOverrides, type ImportOutcome, type ImportProgress, type ImportSession, type ImportedPage } from './importSession'

const MiB = 1024 * 1024

export interface PagedSheet {
  readonly sheetId: string
  readonly name: string
  readonly workbookIndex: number
  readonly partPath: string
  readonly kind: 'worksheet' | 'other'
  state: 'pending' | 'loading' | 'complete' | 'incomplete'
  /** Every row at or before this one is committed (pages cover contiguous rows from 0). */
  importedThroughRow: number
  metadata: SheetMetadata
  /** undefined: the table inventory is unavailable (fail-closed, as parse.ts). */
  tables: TableMetadata[] | undefined
  declared?: DeclaredSheetFilters
  used?: { r0: number; r1: number; c0: number; c1: number }
  /** Array formulas with their saved extent parsed once (undefined: no/invalid ref). */
  arrayCells: { row: number; col: number; saved?: SourceRect; hasRef: boolean; dynamic: boolean }[]
  /** Placement-only drawing anchors (S3.3), read once the sheet's relationships are known. */
  drawingAnchors: XlsxDrawingAnchors
  /** Merge ranges parsed once at the end of the sheet; unreadable refs are counted, not kept. */
  mergeRects: SourceRect[]
  invalidMerges: number
  failure?: SourceFailure
}
export interface PagedWorkbook {
  readonly docId: string
  readonly store: SourcePageStore
  readonly sheets: readonly PagedSheet[]
  readonly diagnostics: WorkerDiagnostic[]
  completeness: Completeness
  failure?: SourceFailure
  /** Formulas the engine would have evaluated at load (values.ts `isLoadEvaluated`)
   * seen so far; final once `completeness` is 'complete'. Feeds the saved-values status. */
  loadEvaluatedFormulas: number
  /** Resolves on every committed page and on completion (adapters wait on it). */
  changed(): Promise<void>
  /** Style table and date system (S5.4), read before the first sheet. */
  formatting: WorkbookFormatting
  /** Only the custom definitions named by this workbook's tables are retained. */
  tableStyleDefs?: ReadonlyMap<string, TableStyleDef>
  /** Referenced table/pivot and slicer styles survive the same bounded retention pass. */
  tableStyleExtras?: StylesExtras
  tableStyleTheme?: ThemeContext
  /** The package behind the import, for on-demand bounded part reads (S13.2:
   * a sheet's drawing content is read when the viewer asks for it). */
  readonly parts: { readonly src: RangeSource; readonly inventory: ZipInventory; readonly signal: AbortSignal }
  /** This session's per-part cap for pivot, cache, slicer and timeline parts (Ruling AP). */
  readonly pivotPartBytes: number
}
export interface PagedImportOptions {
  docId: string
  profile: ProfileId
  signal?: AbortSignal
  limits?: ImportLimitOverrides
  onProgress?(p: ImportProgress): void
  /** Where pages go; default: a store on the backend picked by choosePageBackend. */
  store?: SourcePageStore
  storage?: PageStorageEnv
  ledger?: QuotaLedger
  /** Rows per page cap (default 65,536); the byte target still applies. */
  pageRows?: number
  /** Called once per committed page (tests observe progressive commits). */
  onCommit?(page: ImportedPage): void
  /** Background priority (spec §10.2): awaited before each source read, so the
   * host can park this import while foreground commands (viewport, cell
   * reads) run. */
  beforeRead?(signal: AbortSignal): Promise<void> | void
}
export interface PagedImportSession extends ImportSession {
  readonly workbook: PagedWorkbook
  readonly ledger: QuotaLedger
}

const SST_CHUNK = 4096
let swept = false
/** Page bytes per compressed package byte for the quota precheck (F2 measured 4.9; margin to 6). */
const STORED_PER_COMPRESSED_BYTE = 6
const SMALL_PART_BYTES = 4 * MiB
const utf8 = new TextDecoder()

function controlStyleNames(text: string | undefined, containerName: 'slicers' | 'timelines', itemName: 'slicer' | 'timeline'): string[] {
  if (!text) return []
  try {
    const root = parseXml(text), container = getChildren(root, containerName)[0] ?? root
    return [...new Set(getChildren(container, itemName).map(node => attrs(node).style).filter((name): name is string => Boolean(name)))].slice(0, 256)
  } catch { return [] }
}

function asFailure(error: unknown): SourceFailure {
  if (error instanceof PackageError) return { code: error.code, retryable: error.code === 'storage-error' || error.code === 'quota' || error.code === 'cancelled', message: error.message }
  return { code: 'parse-error', retryable: false, message: error instanceof Error ? error.message : String(error) }
}

/**
 * Opens a paged import. Resolves once the workbook's sheet list is known; the
 * rest of the import runs in the background, committing pages while
 * `completeness` is 'loading'. `done` settles complete, incomplete (typed
 * failure, never a complete blank source) or cancelled.
 */
export async function openImportSession(source: Extract<SourceHandle, { kind: 'blob' }>, options: PagedImportOptions): Promise<PagedImportSession> {
  const limits = importLimitsFor(options.profile, options.limits)
  const controller = new AbortController()
  const forward = () => controller.abort()
  if (options.signal?.aborted) controller.abort()
  else options.signal?.addEventListener('abort', forward, { once: true })
  const signal = controller.signal
  const raw = blobRangeSource(source.blob), beforeRead = options.beforeRead
  const src: typeof raw = beforeRead ? { size: raw.size, async read(start, end, readSignal) {
    const parked = beforeRead(signal)
    if (parked) await parked
    return raw.read(start, end, readSignal)
  } } : raw
  // One ledger per import: options.ledger, else the injected store's own, else a
  // fresh one. Two different ledgers would split the charges, so that is refused.
  if (options.ledger && options.store?.ledger && options.store.ledger !== options.ledger) {
    throw new RangeError('openImportSession: options.ledger differs from the injected store ledger')
  }
  const ledger = options.ledger ?? options.store?.ledger ?? createQuotaLedger(quotaLimitsFor(options.profile), { maxTotalExpandedBytes: limits.paged.expandedBytes })
  // An injected store without a ledger is charged on the session ledger.
  options.store?.attachLedger(ledger)
  const partOpts = (name: string, cap: number) => ({ signal, budget: ledger.partBudget(name, cap), rangeReadBytes: limits.rangeReadBytes, outputBufferBytes: limits.inflateOutputBytes })
  const inventory = await readZipInventory(src, inventoryLimits(limits), signal)
  const smallPart = async (name: string): Promise<string | undefined> => {
    const entry = inventory.find(name)
    if (!entry) return undefined
    try {
      return utf8.decode(await readPartBytes(src, entry, SMALL_PART_BYTES, partOpts(name, SMALL_PART_BYTES)))
    } finally {
      ledger.drainQueue()
    }
  }
  const workbookText = await smallPart('xl/workbook.xml')
  if (workbookText === undefined) throw new PackageError('parse-error', 'xl/workbook.xml missing — not a valid xlsx?')
  const relsText = await smallPart('xl/_rels/workbook.xml.rels')
  const relMap = new Map<string, string>(), relTypes = new Map<string, string>()
  let metadataPath: string | undefined, themePath: string | undefined
  if (relsText !== undefined) {
    for (const rel of getChildren(parseXml(relsText), 'Relationship')) {
      const a = attrs(rel)
      if (a.Id && a.Target && a.TargetMode !== 'External') relMap.set(a.Id, a.Target)
      if (a.Id && a.Type) relTypes.set(a.Id, a.Type)
      if (a.Type === 'http://schemas.openxmlformats.org/officeDocument/2006/relationships/sheetMetadata' && a.TargetMode !== 'External' && a.Target) metadataPath = resolvePartTarget('xl/workbook.xml', a.Target)
      if (a.Type?.endsWith('/theme') && a.TargetMode !== 'External' && a.Target) themePath = resolvePartTarget('xl/workbook.xml', a.Target)
    }
  }
  // parse.ts reads xl/sharedStrings.xml by name, not through the relationship.
  const sstPath = inventory.find('xl/sharedStrings.xml') ? 'xl/sharedStrings.xml' : undefined
  const sheets: PagedSheet[] = []
  const workbookRoot = parseXml(workbookText)
  const calcPrNode = getChildren(workbookRoot, 'calcPr')[0]
  const calcPr = calcPrNode ? attrs(calcPrNode) : {}
  const fullCalcOnLoad = calcPr.fullCalcOnLoad === '1' || calcPr.fullCalcOnLoad === 'true'
  let workbookIndex = 0
  for (const sheetNode of getChildren(getChildren(workbookRoot, 'sheets')[0], 'sheet')) {
    const a = attrs(sheetNode)
    const name = (a.name as string) ?? 'Sheet'
    const sheetId = typeof a.sheetId === 'string' && a.sheetId !== '' ? a.sheetId : String(workbookIndex + 1)
    const relId = (a['r:id'] ?? a.id ?? '') as string
    const target = relMap.get(relId) ?? ''
    const relType = relTypes.get(relId) ?? ''
    const kind = relType.endsWith('/worksheet') ? 'worksheet' : relType !== '' ? 'other' : 'worksheet'
    const partPath = target.startsWith('/') ? target.slice(1) : `xl/${target.replace(/^\.\.\//, '')}`
    sheets.push({ sheetId, name, workbookIndex, partPath, kind, state: 'pending', importedThroughRow: -1, metadata: newSheetMetadata(), tables: [], arrayCells: [], drawingAnchors: { records: [], unresolved: 0 }, mergeRects: [], invalidMerges: 0 })
    workbookIndex++
  }
  const storageEnv = options.storage ?? browserPageStorageEnv()
  let store = options.store
  if (!store) {
    const owner = pageOwner(options.docId)
    const backend = await choosePageBackend(storageEnv, owner, source.blob.size * STORED_PER_COMPRESSED_BYTE, limits.spool.memoryLimitBytes)
    // Browser storage is shared by every tab and worker: hold the owner's Web Lock
    // for the store's lifetime so no other realm's sweep can take it.
    const ownerLock = backend.kind !== 'memory' && storageEnv.locks ? await storageEnv.locks.hold(owner) : undefined
    store = new SourcePageStore({
      docId: options.docId,
      backend,
      ledger,
      // Construction is charged against the profile's worker budget until the viewer binds the document pool.
      memoryBudgetBytes: WORKER_BUDGETS[options.profile],
      ownerLock,
    })
  }
  // Once per process: page storage left by crashed sessions (OPFS/IndexedDB owners older than 24 h).
  // Detached, but its working set is construction memory of this store (§12.4);
  // closing the store stops the sweep, whose charge lasts until it returned.
  if (!swept) {
    swept = true
    void store.sweepLeftovers(storageEnv)
  }
  let notify: () => void = () => {}
  let changedPromise = new Promise<void>((resolve) => { notify = resolve })
  const pulse = () => { const fire = notify; changedPromise = new Promise<void>((resolve) => { notify = resolve }); fire() }
  const workbook: PagedWorkbook = {
    docId: options.docId, store, sheets, diagnostics: [], completeness: 'loading', loadEvaluatedFormulas: 0,
    changed: () => changedPromise, formatting: DEFAULT_FORMATTING,
    parts: { src, inventory, signal }, pivotPartBytes: limits.paged.pivotPartBytes,
  }
  const stamp: SourceStamp = { docId: options.docId, sheetId: sheets[0]?.sheetId ?? '', valueGen: 0 }
  const plannedParts = ['xl/workbook.xml', 'xl/_rels/workbook.xml.rels', sstPath, metadataPath, ...sheets.map((s) => s.partPath)]
    .filter((p): p is string => p !== undefined && inventory.find(p) !== undefined)
  let expandedTotal = plannedParts.reduce((sum, p) => sum + inventory.find(p)!.uncompressedSize, 0)
  // Parts discovered later (worksheet rels, tables, drawings) join the total
  // before they are read, so expandedBytes never exceeds the reported total.
  const planned = new Set(plannedParts)
  const plan = (name: string) => {
    const entry = inventory.find(name)
    if (entry && !planned.has(name)) { planned.add(name); expandedTotal += entry.uncompressedSize }
  }
  const partTexts = new Map<string, string | undefined>()
  /** Drawing parts tokenized once and shared by every sheet that references them. */
  const drawingParts = new Map<string, XlsxDrawingAnchors>()
  const readPlanned = async (name: string): Promise<string | undefined> => {
    // Each discovered part is inflated once, however often it is referenced.
    if (partTexts.has(name)) return partTexts.get(name)
    plan(name)
    const text = await smallPart(name)
    partTexts.set(name, text)
    return text
  }
  let compressedRead = 0
  let populatedCells = 0
  let cellsSeen = 0
  const progress = (phase: ImportProgress['phase']) => options.onProgress?.({
    phase, compressedRead, compressedTotal: source.blob.size, expandedBytes: ledger.snapshot().totalExpanded, expandedTotal, populatedCells,
  })
  for (const p of ['xl/workbook.xml', 'xl/_rels/workbook.xml.rels']) compressedRead += inventory.find(p)?.compressedSize ?? 0
  progress('inventory')

  const run = async (): Promise<ImportOutcome> => {
    // Store open: records this owner holds but its index does not reference
    // (left by a crashed session that reused the owner) are removed first.
    const sweep = await store.sweepUnreferenced(signal)
    // Best-effort janitor: a failed sweep is a diagnostic, never an import failure.
    // (A store without its owner lock skips by design and reports nothing.)
    if (sweep.skipped && !sweep.skipped.startsWith('the store does not hold')) {
      workbook.diagnostics.push({ code: 'storage-cleanup-skipped', message: `Leftover page records were not cleaned up: ${sweep.skipped}` })
    }
    // Shared strings, chunked into the store.
    let sstCount = 0
    if (sstPath) {
      const entry = inventory.find(sstPath)!
      let chunk: string[] = []
      let chunkBytes = 0
      let chunkIndex = 0
      const pending: Array<{ index: number; strings: string[]; bytes: number }> = []
      // Strings waiting for their chunk commit are construction memory (§12.4).
      const share = store.construction.share('pending shared strings')
      let pendingBytes = 0
      const events = sharedStringEvents(sstPath, (text) => {
        const cost = 2 * text.length + PENDING_STRING_BYTES
        share.need(pendingBytes + cost)
        pendingBytes += cost
        chunkBytes += cost
        chunk.push(text)
        sstCount++
        if (sstCount > limits.paged.uniqueStrings) throw new PackageError('resource-limit', `more than ${limits.paged.uniqueStrings} shared strings`, sstPath)
        if (chunk.length === SST_CHUNK) { pending.push({ index: chunkIndex++, strings: chunk, bytes: chunkBytes }); chunk = []; chunkBytes = 0 }
      }, limits.maxTokenChars)
      const flush = async () => {
        for (const c of pending.splice(0)) {
          await store.commitStrings(c.index, c.strings, SST_CHUNK, options.docId)
          pendingBytes -= c.bytes
          share.need(pendingBytes)
        }
      }
      try {
        try {
          await tokenizeBytes(openPart(src, entry, partOpts(sstPath, limits.paged.sharedStringsBytes)), events, { signal, limits: { maxTokenChars: limits.maxTokenChars }, afterChunk: flush })
        } finally {
          ledger.drainQueue()
        }
        if (chunk.length) pending.push({ index: chunkIndex++, strings: chunk, bytes: chunkBytes })
        await flush()
      } finally {
        share.release()
      }
      compressedRead += entry.compressedSize
      progress('sharedStrings')
    }
    let dynamicIndices: ReadonlySet<number> = new Set()
    if (metadataPath) {
      try {
        const text = await smallPart(metadataPath)
        if (text) dynamicIndices = dynamicMetadataIndices(text)
        compressedRead += inventory.find(metadataPath)?.compressedSize ?? 0
      } catch (error) {
        if (isPackageError(error) && error.code === 'cancelled') throw error
        /* Unverified optional metadata never upgrades a fixed array (parse.ts). */
      }
    }
    let sparklineColorContext: SheetColorContext = {}
    let tableStylesTheme: ThemeContext | undefined
    let stylesTextForPriority: string | undefined, themeTextForPriority: string | undefined
    const prioritizedTableStyleNames = new Set<string>(), prioritizedSlicerStyleNames = new Set<string>()
    try {
      const stylesText = inventory.find('xl/styles.xml') ? await readPlanned('xl/styles.xml') : undefined
      stylesTextForPriority = stylesText
      // The theme colours both cell styles and sparkline groups; an unreadable theme keeps the Office default.
      let themeText: string | undefined
      if (themePath && inventory.find(themePath)) {
        try { themeText = await readPlanned(themePath) } catch (error) { if (isPackageError(error) && error.code === 'cancelled') throw error }
      }
      themeTextForPriority = themeText
      const parsedFormatting = await workbookFormattingWithColorContext(workbookRoot, stylesText, themeText)
      workbook.formatting = parsedFormatting.formatting
      sparklineColorContext = parsedFormatting.colorContext
      workbook.tableStyleDefs = parsedFormatting.tableStyleDefs
      workbook.tableStyleExtras = parsedFormatting.tableStyleExtras
      if (themeText !== undefined) {
        const { parseThemeContext } = await import('../../drawing/style')
        tableStylesTheme = parseThemeContext(themeText)
      }
      workbook.tableStyleTheme = tableStylesTheme
    } catch (error) {
      if (isPackageError(error) && error.code === 'cancelled') throw error
      // parse.ts tolerates an unreadable styles part: cells keep General format.
      workbook.diagnostics.push({ code: 'styles-unreadable', message: `xl/styles.xml could not be read: ${(error as Error).message}` })
    }
    progress('styles')
    const declaredBudget = createDeclaredBudget()
    const pageRows = options.pageRows ?? 65_536
    for (const sheet of sheets) {
      throwIfAborted(signal)
      const entry = inventory.find(sheet.partPath)
      sheet.state = 'loading'
      if (!entry) {
        // parse.ts: an unreadable part is an empty sheet, but a worksheet's table inventory is unknown.
        if (sheet.kind === 'worksheet') sheet.tables = undefined
        sheet.importedThroughRow = MAX_ROWS - 1
        sheet.state = 'complete'
        pulse()
        continue
      }
      const declaredEntry: DeclaredSheetFilters = { sheetId: sheet.sheetId, name: sheet.name, workbookIndex: sheet.workbookIndex, partPath: sheet.partPath, tableAutoFilters: [] }
      const container = declaredContainerBuilder(declaredBudget)
      const newBuilder = () => new PageBuilder({ inlineTextBytes: limits.inlineStringSoftBytes })
      let builder = newBuilder()
      let pageStart = sheet.importedThroughRow + 1
      let rowsInPage = 0
      let pageSeq = 0
      let r0 = MAX_ROWS, r1 = -1, c0 = MAX_COLS, c1 = -1
      const pending: StoredPage[] = []
      // Builder buffers and built pages waiting for their commit are construction
      // memory (§12.4), charged before they grow; the store charges encode scratch.
      const share = store.construction.share('page construction')
      let pendingBytes = 0
      // payloadBytes carries the page's formula-metadata ownership from the builder.
      const pageCost = (page: StoredPage) => payloadBytes(page.payload) + (page.blobs ?? []).reduce((n, b) => n + TEXT_STRING_BYTES + 2 * b.length, 0)
      const account = (extra = 0) => share.need(builder.memoryBytes + pendingBytes + extra)
      const cut = (lastRow: number) => {
        // build() copies the builder into typed lanes of (about) the encoded size.
        account(builder.bytes)
        const payload = builder.build()
        const page: StoredPage = { sheetId: sheet.sheetId, pageId: `p${pageSeq++}`, r0: pageStart, r1: lastRow, cellCount: payload.rows.length, bytes: builder.bytes, payload, blobs: builder.blobs }
        pending.push(page)
        pendingBytes += pageCost(page)
        builder = newBuilder()
        account()
        pageStart = lastRow + 1
        rowsInPage = 0
      }
      const commitPending = async () => {
        for (const page of pending.splice(0)) {
          throwIfAborted(signal)
          await store.commit(page, options.docId)
          pendingBytes -= pageCost(page)
          account()
          sheet.importedThroughRow = page.r1
          populatedCells += page.cellCount
          options.onCommit?.(page)
          progress('sheet')
          pulse()
        }
      }
      const events = sheetRecordEvents(sheet.metadata, {
        partId: sheet.partPath,
        sharedStringCount: sstCount,
        dynamicIndices,
        maxTextChars: limits.maxTokenChars,
        maxStructuralRecords: limits.paged.structuralRecords,
        sparklineColorContext,
        onSparklineDiagnostic(diagnostic) {
          workbook.diagnostics.push({ code: diagnostic.feature, message: diagnostic.message, partId: sheet.partPath })
        },
        declared: container,
        onCell(record) {
          // Charged before the builder retains the record's text and formula metadata.
          account(builder.costOf(record.text, record.meta))
          builder.add(record.row, record.col, record.kind, record.styleIndex, record.num, record.text, record.cached, record.meta, record.valueIsError)
          const meta = record.meta
          if ((meta?.formula !== undefined || meta?.sharedFormula !== undefined)
            && isLoadEvaluated({ cached: record.cached, calcAlways: meta.ca === true, text: meta.formula }, fullCalcOnLoad)) workbook.loadEvaluatedFormulas++
          if (record.row < r0) r0 = record.row
          if (record.row > r1) r1 = record.row
          if (record.col < c0) c0 = record.col
          if (record.col > c1) c1 = record.col
          if (record.meta?.arrayRef !== undefined || record.meta?.dynamicArray) {
            if (sheet.arrayCells.length >= limits.paged.structuralRecords) throw new PackageError('resource-limit', `more than ${limits.paged.structuralRecords} array formulas in one sheet`, sheet.partPath)
            const ref = record.meta.arrayRef
            sheet.arrayCells.push({ row: record.row, col: record.col, saved: ref === undefined ? undefined : rangeOf(ref) ?? undefined, hasRef: ref !== undefined, dynamic: record.meta.dynamicArray === true })
          }
          if (++cellsSeen > limits.paged.cells) throw new PackageError('resource-limit', `more than ${limits.paged.cells} cells`, sheet.partPath)
        },
        onRowEnd(row) {
          rowsInPage++
          if (row >= MAX_ROWS - 1) return
          // Spilled values count toward the cut too, so oversized values cannot pile up behind one page.
          if (builder.bytes + builder.spilledBytes >= limits.targetPageBytes || rowsInPage >= pageRows) cut(row)
        },
      })
      try {
        try {
          await tokenizeBytes(openPart(src, entry, partOpts(sheet.partPath, limits.paged.expandedBytes)), events, {
            signal, limits: { maxTokenChars: limits.maxTokenChars }, afterChunk: commitPending,
          })
        } catch (error) {
          sheet.state = 'incomplete'
          sheet.failure = asFailure(error)
          throw error
        } finally {
          ledger.drainQueue()
        }
        // The last page covers the tail: every row after it is known blank.
        cut(MAX_ROWS - 1)
        await commitPending()
      } finally {
        // Completion, failure and cancellation all return the construction share.
        share.release()
      }
      if (r1 >= 0) sheet.used = { r0, r1, c0, c1 }
      normalizeRowMetadata(sheet.metadata)
      for (const ref of sheet.metadata.merges) {
        const rect = rangeOf(ref)
        if (rect) sheet.mergeRects.push(rect)
        else sheet.invalidMerges++
      }
      sheet.metadata.merges = []
      compressedRead += entry.compressedSize
      // Declared worksheet filters, then tables through the worksheet's own rels (parse.ts order).
      const state = container.finish()
      if (state.autoFilter) declaredEntry.worksheetAutoFilter = state.autoFilter
      if (state.sortState) declaredEntry.worksheetSortState = state.sortState
      const wsRels = await worksheetRels(sheet, readPlanned)
      sheet.tables = await resolveTables(sheet, wsRels, declaredEntry, declaredBudget, readPlanned)
      const resolveControlRefs = (ids: string[], suffix: string): ControlPartRef[] | undefined => {
        const refs = [...new Set(ids)].flatMap(id => {
          const rel = wsRels?.get(id)
          return rel?.type.endsWith(suffix) ? [{ partPath: resolvePartTarget(sheet.partPath, rel.target) }] : []
        })
        return refs.length ? refs : undefined
      }
      sheet.metadata.slicerRefs = resolveControlRefs(sheet.metadata.slicerRelationshipIds, '/slicer')
      sheet.metadata.timelineRefs = resolveControlRefs(sheet.metadata.timelineRelationshipIds, '/timeline')
      const readControlStyles = async (refs: readonly ControlPartRef[] | undefined, containerName: 'slicers' | 'timelines', itemName: 'slicer' | 'timeline') => {
        const out: string[] = []
        for (const ref of refs ?? []) {
          const controlEntry = inventory.find(ref.partPath)
          if (!controlEntry || controlEntry.uncompressedSize > limits.paged.pivotPartBytes) continue
          plan(controlEntry.name)
          try {
            const text = utf8.decode(await readPartPrefixBytes(src, controlEntry, limits.paged.pivotPartBytes, partOpts(controlEntry.name, limits.paged.pivotPartBytes)))
            out.push(...controlStyleNames(text, containerName, itemName))
          } catch (error) {
            if (isPackageError(error) && error.code === 'cancelled') throw error
          } finally { ledger.drainQueue() }
        }
        return [...new Set(out)].slice(0, 256)
      }
      sheet.metadata.slicerStyleNames = await readControlStyles(sheet.metadata.slicerRefs, 'slicers', 'slicer')
      sheet.metadata.timelineStyleNames = await readControlStyles(sheet.metadata.timelineRefs, 'timelines', 'timeline')
      for (const rel of wsRels?.values() ?? []) {
        if (!rel.type.endsWith('/pivotTable')) continue
        const partPath = resolvePartTarget(sheet.partPath, rel.target)
        const pivotEntry = inventory.find(partPath)
        const pivotTooLarge = Boolean(pivotEntry && pivotEntry.uncompressedSize > limits.paged.pivotPartBytes)
        if (pivotTooLarge) {
          ;(sheet.metadata.pivotDiagnostics ??= []).push({ kind: 'unsupported', feature: 'pivot-part-too-large', message: `Pivot part ${partPath} exceeds the ${limits.paged.pivotPartBytes}-byte part limit`, partPath })
        }
        let text: string | undefined
        if (pivotEntry) {
          plan(partPath)
          try {
            // The location scan reads a prefix only, as the resident parser does (stub.ts parsePivotRegion).
            text = utf8.decode(await readPartPrefixBytes(src, pivotEntry, PIVOT_PART_MAX_BYTES, partOpts(partPath, PIVOT_PART_MAX_BYTES)))
          } catch (error) {
            if (isPackageError(error) && error.code === 'cancelled') throw error
          } finally { ledger.drainQueue() }
        }
        const parsed = parsePivotRegion(text, partPath)
        if (parsed.region) (sheet.metadata.pivotRegions ??= []).push(parsed.region)
        else workbook.diagnostics.push({ code: 'pivot-location-unreadable', message: `Pivot location could not be read from ${partPath}`, partId: partPath })
      }
      // Placement-only anchor records (S3.3), as drawing.ts reads them: an
      // unresolvable or unreadable part counts unresolved, never "no anchors".
      const anchoredParts = new Set<string>()
      for (const id of sheet.metadata.drawingIds) {
        const rel = id === undefined ? undefined : wsRels?.get(id)
        const drawing = rel && rel.type.endsWith('/drawing') ? inventory.find(resolvePartTarget(sheet.partPath, rel.target)) : undefined
        if (!drawing) { sheet.drawingAnchors.unresolved++; continue }
        if (anchoredParts.has(drawing.name)) continue
        anchoredParts.add(drawing.name)
        let parsed = drawingParts.get(drawing.name)
        if (!parsed) {
          plan(drawing.name)
          parsed = { records: [], unresolved: 0 }
          let index = 0
          const found = parsed
          const markerShare = store.construction.share('streamed drawing marker fields', { untilReleased: true })
          let markerEvents: ReturnType<typeof anchorRecordEvents> | undefined
          try {
            markerEvents = anchorRecordEvents((name, parts) => {
              const record = found.records.length < limits.paged.structuralRecords ? anchorRecord(name, anchorId(drawing.name, index), parts) : undefined
              index++
              if (record) found.records.push(record)
              else found.unresolved++
            }, () => { found.unresolved++ }, { need: (bytes) => markerShare.need(bytes) })
            await tokenizeBytes(openPart(src, drawing, partOpts(drawing.name, limits.paged.drawingPartEnvelopeBytes)), markerEvents, { signal, limits: { maxTokenChars: limits.maxTokenChars } })
          } catch (error) {
            if (isPackageError(error) && error.code === 'cancelled') throw error
            parsed = { records: [], unresolved: 1 }
            if (isPackageError(error) && error.reason === 'memory-budget') workbook.diagnostics.push({ code: 'import-memory-budget', message: error.message, partId: drawing.name })
          } finally {
            markerEvents?.dispose()
            markerShare.release()
            ledger.drainQueue()
          }
          drawingParts.set(drawing.name, parsed)
        }
        sheet.drawingAnchors.records.push(...parsed.records)
        sheet.drawingAnchors.unresolved += parsed.unresolved
      }
      if (hasDeclaredFiltering(declaredEntry)) sheet.declared = declaredEntry
      for (const domain of sheet.metadata.structuralIncomplete) {
        workbook.diagnostics.push({ code: 'structure-incomplete', message: `${domain} metadata in ${sheet.partPath} is wrapped in mc:AlternateContent and cannot be read`, partId: sheet.partPath })
      }
      if (sheet.metadata.skippedCells) {
        workbook.diagnostics.push({ code: 'cells-outside-sheet', message: `${sheet.metadata.skippedCells} cells outside the sheet grid were ignored`, partId: sheet.partPath })
      }
      // Include this candidate when choosing referenced styles, but keep its
      // public state loading until the priority re-parse and cache replacement
      // have both settled. A concurrent display read must not observe stale
      // style extras under a `complete` state.
      const completeSheets = sheets.filter(candidate => candidate === sheet || candidate.state === 'complete')
      const referencedTableStyles = new Set(completeSheets.flatMap(candidate => [
        ...(candidate.tables?.map(table => table.styleInfo?.name).filter((name): name is string => Boolean(name)) ?? []),
        ...(candidate.metadata.pivotRegions?.map(region => region.styleName).filter((name): name is string => Boolean(name)) ?? []),
        ...(candidate.metadata.slicerStyleNames ?? []), ...(candidate.metadata.timelineStyleNames ?? []),
      ]))
      const referencedSlicerStyles = new Set(completeSheets.flatMap(candidate => [
        ...(candidate.metadata.slicerStyleNames ?? []), ...(candidate.metadata.timelineStyleNames ?? []),
      ]))
      const needsPreferredStyles = [...referencedTableStyles].some(name => !prioritizedTableStyleNames.has(name) && !/^(?:Table|Pivot)Style(?:Light|Medium|Dark)\d+$/.test(name) && !workbook.tableStyleDefs?.has(name))
        || [...referencedSlicerStyles].some(name => !prioritizedSlicerStyleNames.has(name) && !workbook.tableStyleExtras?.slicerStyles.has(name))
      // The styles extras load lazily (spec §6.2), for the first sheet with a pivot, a slicer or a timeline: a pivot's
      // `formats` need the dxfs even under a built-in style, and a slicer without its own style takes the default.
      const needsExtras = !workbook.tableStyleExtras && Boolean(sheet.metadata.pivotRegions?.length || sheet.metadata.slicerRefs?.length || sheet.metadata.timelineRefs?.length)
      if (stylesTextForPriority !== undefined && (needsPreferredStyles || needsExtras)) {
        const prioritized = await workbookFormattingWithColorContext(workbookRoot, stylesTextForPriority, themeTextForPriority, referencedTableStyles, referencedSlicerStyles, 'always')
        workbook.tableStyleDefs = prioritized.tableStyleDefs
        workbook.tableStyleExtras = prioritized.tableStyleExtras
        for (const name of referencedTableStyles) prioritizedTableStyleNames.add(name)
        for (const name of referencedSlicerStyles) prioritizedSlicerStyleNames.add(name)
      }
      sheet.state = 'complete'
      pulse()
    }
    progress('done')
    return { kind: 'complete', stamp }
  }

  let completeness: Completeness = 'loading'
  let released = false
  const done: Promise<ImportOutcome> = (async () => {
    try {
      const outcome = await run()
      completeness = 'complete'
      workbook.completeness = 'complete'
      return outcome
    } catch (error) {
      for (const sheet of sheets) if (sheet.state !== 'complete') sheet.state = 'incomplete'
      if (signal.aborted || (isPackageError(error) && error.code === 'cancelled')) {
        completeness = 'incomplete'
        workbook.completeness = 'incomplete'
        workbook.failure = { code: 'cancelled', retryable: true, message: 'Import cancelled' }
        await store.abandon(options.docId).catch(() => {})
        ledger.releaseAll()
        return { kind: 'cancelled' }
      }
      const failure = asFailure(error)
      completeness = 'incomplete'
      workbook.completeness = 'incomplete'
      workbook.failure = failure
      // A construction peak the worker pool cannot hold has its own typed diagnostic (§12.4).
      const memory = isPackageError(error) && error.reason === 'memory-budget'
      workbook.diagnostics.push({ code: memory ? 'import-memory-budget' : `import-${failure.code}`, message: failure.message })
      // A failed page write leaves nothing trustworthy behind: remove every page.
      if (failure.code === 'storage-error' || failure.code === 'quota') await store.abandon(options.docId).catch(() => {})
      return { kind: 'incomplete', failure, stamp }
    } finally {
      options.signal?.removeEventListener('abort', forward)
      pulse()
    }
  })()
  return {
    stamp,
    workbook,
    ledger,
    get completeness() { return completeness },
    get diagnostics() { return workbook.diagnostics },
    done,
    async cancel() {
      controller.abort()
      await done
    },
    async release() {
      if (released) return
      released = true
      controller.abort()
      await done
      await store.close()
      ledger.releaseAll()
    },
  }
}

type RelTargets = Map<string, { target: string; type: string }>
/** The worksheet's own relationships; undefined when they could not be read. */
async function worksheetRels(sheet: PagedSheet, read: (name: string) => Promise<string | undefined>): Promise<RelTargets | undefined> {
  const slash = sheet.partPath.lastIndexOf('/')
  const relsPath = `${slash < 0 ? '' : `${sheet.partPath.slice(0, slash)}/`}_rels/${slash < 0 ? sheet.partPath : sheet.partPath.slice(slash + 1)}.rels`
  const targets: RelTargets = new Map()
  try {
    const relsText = await read(relsPath)
    if (relsText !== undefined) {
      for (const rel of getChildren(parseXml(relsText), 'Relationship')) {
        const a = attrs(rel)
        if (a.Id && a.Target && a.TargetMode !== 'External') targets.set(a.Id, { target: a.Target, type: a.Type ?? '' })
      }
    }
  } catch (error) {
    if (isPackageError(error) && error.code === 'cancelled') throw error
    return undefined
  }
  return targets
}
async function resolveTables(sheet: PagedSheet, targets: RelTargets | undefined, entry: DeclaredSheetFilters, budget: ReturnType<typeof createDeclaredBudget>, read: (name: string) => Promise<string | undefined>): Promise<TableMetadata[] | undefined> {
  const decls = sheet.metadata.tablePartIds
  if (decls.length === 0) return []
  const sheetDeclared = declaredSheetBuilder(entry, budget)
  let complete = targets !== undefined
  const resolved: TableMetadata[] = []
  try {
    const checks = new Map<string, boolean>()
    for (const rid of decls) {
      if (rid === undefined) { complete = false; continue }
      const rel = targets?.get(rid)
      if (!rel || !rel.type.endsWith('/table')) { complete = false; continue }
      const partPath = resolvePartTarget(sheet.partPath, rel.target)
      try {
        const text = await read(partPath)
        let ok = checks.get(partPath)
        if (ok === undefined) { ok = isTablePartSource(text); checks.set(partPath, ok) }
        if (!ok || text === undefined) { complete = false; continue }
        const root = parseXml(text)
        const node = getChildren(root, 'table')[0] ?? root
        if (node) collectDeclaredTable(sheetDeclared, node, partPath)
        const meta = parseTableMetadata(node, sheet.sheetId, partPath)
        if (!meta) { complete = false; continue }
        resolved.push(meta)
      } catch (error) {
        if (isPackageError(error) && error.code === 'cancelled') throw error
        complete = false
      }
    }
  } catch (error) {
    if (isPackageError(error) && error.code === 'cancelled') throw error
    complete = false
  }
  return complete ? resolved : undefined
}

/** Resolves undefined once a page is committed, or with the outcome if the import ends first. */
export async function firstPageOrOutcome(session: PagedImportSession): Promise<ImportOutcome | undefined> {
  for (;;) {
    if (session.workbook.store.stats.committed > 0) return undefined
    if (session.completeness !== 'loading') return session.done
    const settled = await Promise.race([session.done, session.workbook.changed().then(() => undefined)])
    if (settled) return settled
  }
}

/** "A1" or "A1:B2" as an ordered inclusive rect; null when unreadable (same rule as the object adapter). */
export function rangeOf(ref: string): SourceRect | null {
  try {
    const [first, last = first] = ref.split(':')
    const a = parseA1(first), b = parseA1(last)
    return { r0: Math.min(a.row, b.row), r1: Math.max(a.row, b.row), c0: Math.min(a.col, b.col), c1: Math.max(a.col, b.col) }
  } catch {
    return null
  }
}

/**
 * Whether a file the paged reader cannot serve may fall back to today's
 * legacy parse (worker and internal entry only): the compressed package must
 * fit the object path's buffer cap (G1) AND the declared expanded size of its
 * non-media parts must fit the object-path memory budget. Beyond either the
 * caller surfaces the typed incomplete/refused outcome. An unreadable central
 * directory never allows the fallback.
 */
export async function compatibilityFallbackAllowed(blob: Blob, limits: ImportLimits, signal?: AbortSignal): Promise<boolean> {
  if (blob.size > limits.object.compressedBytes) return false
  try {
    const inventory = await readZipInventory(blobRangeSource(blob), inventoryLimits(limits), signal)
    let declared = 0
    for (const entry of inventory.entries) if (classifyPart(entry.name).cls !== 'media') declared += entry.uncompressedSize
    return declared <= limits.object.budgetBytes
  } catch (error) {
    if (isPackageError(error) && error.code === 'cancelled') throw error
    return false
  }
}
