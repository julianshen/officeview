/**
 * Paged importer (S1.2, design note §4.4). Kept out of importSession.ts so the
 * eager bundle (which imports the limits and ImportRefusal from there) never
 * pulls in the ZIP/XML readers: only the lazy large-source chunk loads this.
 */
import type { Completeness, ProfileId, SourceFailure, SourceHandle, SourceStamp, WorkerDiagnostic } from '../view/contracts'
import { PackageError, isPackageError, throwIfAborted } from '../../core/package-error'
import { blobRangeSource, openPart, readPartBytes, readZipInventory } from '../../core/zip-stream'
import { tokenizeBytes } from '../../core/xml-stream'
import { attrs, getChildren, parseXml } from '../../core/xml'
import { resolvePartTarget } from '../../drawing/parts'
import { dynamicMetadataIndices } from '../metadata'
import { isTablePartSource, parseTableMetadata } from '../formula/tables'
import type { TableMetadata } from '../formula/types'
import {
  collectDeclaredTable, createDeclaredBudget, declaredContainerBuilder, declaredSheetBuilder, hasDeclaredFiltering, type DeclaredSheetFilters,
} from '../view/declaredFilters'
import { MAX_COLS, MAX_ROWS } from '../view/coordinates'
import { createQuotaLedger, quotaLimitsFor, type QuotaLedger } from './quota'
import { inventoryLimits } from './admission'
import { anchorCountEvents, newSheetMetadata, normalizeRowMetadata, sharedStringEvents, sheetRecordEvents, type SheetMetadata } from './sheetRecords'
import { PageBuilder, SourcePageStore, browserPageStorageEnv, choosePageBackend, type PageStorageEnv, type StoredPage } from './sourcePages'
import { importLimitsFor, type ImportLimitOverrides, type ImportOutcome, type ImportProgress, type ImportSession, type ImportedPage } from './importSession'

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
  arrayCells: { row: number; col: number; arrayRef: string; dynamic: boolean }[]
  failure?: SourceFailure
}
export interface PagedWorkbook {
  readonly docId: string
  readonly store: SourcePageStore
  readonly sheets: readonly PagedSheet[]
  readonly diagnostics: WorkerDiagnostic[]
  completeness: Completeness
  failure?: SourceFailure
  /** Resolves on every committed page and on completion (adapters wait on it). */
  changed(): Promise<void>
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
}
export interface PagedImportSession extends ImportSession {
  readonly workbook: PagedWorkbook
  readonly ledger: QuotaLedger
}

const SST_CHUNK = 4096
/** Page bytes per compressed package byte for the quota precheck (F2 measured 4.9; margin to 6). */
const STORED_PER_COMPRESSED_BYTE = 6
const SMALL_PART_BYTES = 4 * MiB
const utf8 = new TextDecoder()

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
  const src = blobRangeSource(source.blob)
  const ledger = options.ledger ?? createQuotaLedger(quotaLimitsFor(options.profile), { maxTotalExpandedBytes: limits.paged.expandedBytes })
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
  let metadataPath: string | undefined
  if (relsText !== undefined) {
    for (const rel of getChildren(parseXml(relsText), 'Relationship')) {
      const a = attrs(rel)
      if (a.Id && a.Target && a.TargetMode !== 'External') relMap.set(a.Id, a.Target)
      if (a.Id && a.Type) relTypes.set(a.Id, a.Type)
      if (a.Type === 'http://schemas.openxmlformats.org/officeDocument/2006/relationships/sheetMetadata' && a.TargetMode !== 'External' && a.Target) metadataPath = resolvePartTarget('xl/workbook.xml', a.Target)
    }
  }
  // parse.ts reads xl/sharedStrings.xml by name, not through the relationship.
  const sstPath = inventory.find('xl/sharedStrings.xml') ? 'xl/sharedStrings.xml' : undefined
  const sheets: PagedSheet[] = []
  const workbookRoot = parseXml(workbookText)
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
    sheets.push({ sheetId, name, workbookIndex, partPath, kind, state: 'pending', importedThroughRow: -1, metadata: newSheetMetadata(), tables: [], arrayCells: [] })
    workbookIndex++
  }
  const store = options.store ?? new SourcePageStore({
    docId: options.docId,
    backend: await choosePageBackend(options.storage ?? browserPageStorageEnv(), `${options.docId}-${Date.now().toString(36)}`, source.blob.size * STORED_PER_COMPRESSED_BYTE, limits.spool.memoryLimitBytes),
    ledger,
  })
  let notify: () => void = () => {}
  let changedPromise = new Promise<void>((resolve) => { notify = resolve })
  const pulse = () => { const fire = notify; changedPromise = new Promise<void>((resolve) => { notify = resolve }); fire() }
  const workbook: PagedWorkbook = {
    docId: options.docId, store, sheets, diagnostics: [], completeness: 'loading',
    changed: () => changedPromise,
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
    // Shared strings, chunked into the store.
    let sstCount = 0
    if (sstPath) {
      const entry = inventory.find(sstPath)!
      let chunk: string[] = []
      let chunkIndex = 0
      const pending: Array<{ index: number; strings: string[] }> = []
      const events = sharedStringEvents(sstPath, (text) => {
        chunk.push(text)
        sstCount++
        if (sstCount > limits.paged.uniqueStrings) throw new PackageError('resource-limit', `more than ${limits.paged.uniqueStrings} shared strings`, sstPath)
        if (chunk.length === SST_CHUNK) { pending.push({ index: chunkIndex++, strings: chunk }); chunk = [] }
      }, limits.maxTokenChars)
      const flush = async () => { for (const c of pending.splice(0)) await store.commitStrings(c.index, c.strings, SST_CHUNK, options.docId) }
      try {
        await tokenizeBytes(openPart(src, entry, partOpts(sstPath, limits.paged.sharedStringsBytes)), events, { signal, limits: { maxTokenChars: limits.maxTokenChars }, afterChunk: flush })
      } finally {
        ledger.drainQueue()
      }
      if (chunk.length) pending.push({ index: chunkIndex++, strings: chunk })
      await flush()
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
      let builder = new PageBuilder()
      let pageStart = sheet.importedThroughRow + 1
      let rowsInPage = 0
      let pageSeq = 0
      let r0 = MAX_ROWS, r1 = -1, c0 = MAX_COLS, c1 = -1
      const pending: StoredPage[] = []
      const cut = (lastRow: number) => {
        const payload = builder.build()
        pending.push({ sheetId: sheet.sheetId, pageId: `p${pageSeq++}`, r0: pageStart, r1: lastRow, cellCount: payload.rows.length, bytes: builder.bytes, payload })
        builder = new PageBuilder()
        pageStart = lastRow + 1
        rowsInPage = 0
      }
      const commitPending = async () => {
        for (const page of pending.splice(0)) {
          throwIfAborted(signal)
          await store.commit(page, options.docId)
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
        declared: container,
        onCell(record) {
          builder.add(record.row, record.col, record.kind, record.styleIndex, record.num, record.text, record.cached, record.meta, record.valueIsError)
          if (record.row < r0) r0 = record.row
          if (record.row > r1) r1 = record.row
          if (record.col < c0) c0 = record.col
          if (record.col > c1) c1 = record.col
          if (record.meta?.arrayRef !== undefined) sheet.arrayCells.push({ row: record.row, col: record.col, arrayRef: record.meta.arrayRef, dynamic: record.meta.dynamicArray === true })
          else if (record.meta?.dynamicArray) sheet.arrayCells.push({ row: record.row, col: record.col, arrayRef: '', dynamic: true })
          if (++cellsSeen > limits.paged.cells) throw new PackageError('resource-limit', `more than ${limits.paged.cells} cells`, sheet.partPath)
        },
        onRowEnd(row) {
          rowsInPage++
          if (row >= MAX_ROWS - 1) return
          if (builder.bytes >= limits.targetPageBytes || rowsInPage >= pageRows) cut(row)
        },
      })
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
      if (r1 >= 0) sheet.used = { r0, r1, c0, c1 }
      normalizeRowMetadata(sheet.metadata)
      compressedRead += entry.compressedSize
      // Declared worksheet filters, then tables through the worksheet's own rels (parse.ts order).
      const state = container.finish()
      if (state.autoFilter) declaredEntry.worksheetAutoFilter = state.autoFilter
      if (state.sortState) declaredEntry.worksheetSortState = state.sortState
      const wsRels = await worksheetRels(sheet, readPlanned)
      sheet.tables = await resolveTables(sheet, wsRels, declaredEntry, declaredBudget, readPlanned)
      for (const id of sheet.metadata.drawingIds) {
        const rel = id === undefined ? undefined : wsRels?.get(id)
        const drawing = rel && rel.type.endsWith('/drawing') ? inventory.find(resolvePartTarget(sheet.partPath, rel.target)) : undefined
        if (!drawing) { sheet.metadata.drawingAnchors++; continue }
        if (planned.has(drawing.name) && partTexts.has(drawing.name)) continue
        plan(drawing.name)
        partTexts.set(drawing.name, undefined)
        try {
          let anchors = 0
          await tokenizeBytes(openPart(src, drawing, partOpts(drawing.name, limits.paged.drawingPartEnvelopeBytes)), anchorCountEvents(() => { anchors++ }), { signal, limits: { maxTokenChars: limits.maxTokenChars } })
          sheet.metadata.drawingAnchors += anchors
        } catch (error) {
          if (isPackageError(error) && error.code === 'cancelled') throw error
          sheet.metadata.drawingAnchors++
        } finally {
          ledger.drainQueue()
        }
      }
      if (hasDeclaredFiltering(declaredEntry)) sheet.declared = declaredEntry
      for (const domain of sheet.metadata.structuralIncomplete) {
        workbook.diagnostics.push({ code: 'structure-incomplete', message: `${domain} metadata in ${sheet.partPath} is wrapped in mc:AlternateContent and cannot be read`, partId: sheet.partPath })
      }
      if (sheet.metadata.skippedCells) {
        workbook.diagnostics.push({ code: 'cells-outside-sheet', message: `${sheet.metadata.skippedCells} cells outside the sheet grid were ignored`, partId: sheet.partPath })
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
      workbook.diagnostics.push({ code: `import-${failure.code}`, message: failure.message })
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
