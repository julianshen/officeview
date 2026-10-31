/**
 * Resolves a workbook's pivot parts into a PivotWorkbook (spec §3, §6.2). Lazy
 * chunk `pivots`: `parseXlsx` awaits it on the resident path, once the sheets
 * locate a pivot (Ruling AD: attachPivotTables), and the worker runs it over
 * `RangePackage`. Records are not read at load: on the resident path `retainPivotRuntime` keeps the parsed caches
 * and each records part's zip entry, streamed only when a session asks (Task 25).
 *
 * Failure policy, per part:
 * - workbook `pivotCaches` unreadable: diagnostic; each table falls back to its own cache relationship;
 * - definition missing, malformed XML, or no usable `location`/`cacheId`: diagnostic, table skipped;
 * - cache missing or malformed: diagnostic, table kept without a cache entry (display needs only the definition);
 * - cache found only through the definition's relationship: `pivot-cache-mapping-recovered`, table kept;
 * - an index outside its field's domain (`validateIndices`): `invalid-index:<where>` in `unsupported` plus
 *   `pivot-definition-invalid`, table kept display-only;
 * - a declared `count` that disagrees with the children: `pivot-count-mismatch`, the children win;
 * - unknown constructs in either part: recorded in `PivotTableDefinition.unsupported`, table kept.
 */
import type JSZip from 'jszip'
import { attrs, getChild, getChildren, type XmlNode } from '../../core/xml'
import { PackageError } from '../../core/package-error'
import { resolvePartTarget } from '../../drawing/parts'
import type { PivotCacheDefinition, PivotDiagnostic, PivotTableDefinition, PivotWorkbook } from './model'
import type { PivotRegionRef } from './stub'
import { parseCacheDefinition, pushDiagnostic } from './parseCacheDefinition'
import { parsePivotTable, validateIndices } from './parsePivotTable'
import { bindPivotRuntime, createPivotRuntime, type PivotRuntime, type RecordsSource } from './session/runtime'

/**
 * The part access both `OfficePackage` and the worker's `RangePackage` provide.
 * `xmlOrdered` is needed for cache definitions: shared items are indexed by
 * document position across mixed element names, which the grouped `xml` tree loses.
 * The loader never calls `has` on a records part: `RangePackage.has` refuses (and
 * reports) an over-cap part as a side effect, and records may be tens of MB.
 */
export interface PackageReader {
  has(path: string): boolean
  xml(path: string): Promise<XmlNode | undefined>
  xmlOrdered(path: string): Promise<XmlNode | undefined>
}

const WORKBOOK = 'xl/workbook.xml'
const relsPathFor = (owner: string): string => {
  const slash = owner.lastIndexOf('/')
  return `${slash < 0 ? '' : owner.slice(0, slash + 1)}_rels/${owner.slice(slash + 1)}.rels`
}

/** Internal relationship targets of `owner`, by id, keeping only `typeSuffix` types. Unreadable rels read as none. */
export async function relationships(pkg: PackageReader, owner: string, typeSuffix: string): Promise<Map<string, string>> {
  const out = new Map<string, string>()
  let rels: XmlNode | undefined
  try { rels = await pkg.xml(relsPathFor(owner)) } catch { return out }
  for (const rel of getChildren(rels, 'Relationship')) {
    const a = attrs(rel)
    if (a.Id && a.Target && a.TargetMode !== 'External' && (a.Type ?? '').endsWith(typeSuffix)) out.set(a.Id, resolvePartTarget(owner, a.Target))
  }
  return out
}

/** Workbook `pivotCaches`: cacheId → cache definition part. */
export async function workbookCaches(pkg: PackageReader, diagnostics: PivotDiagnostic[]): Promise<Map<number, string>> {
  const out = new Map<number, string>()
  let workbook: XmlNode | undefined
  try { workbook = await pkg.xml(WORKBOOK) } catch { workbook = undefined }
  if (!workbook) {
    pushDiagnostic(diagnostics, { kind: 'parse', feature: 'pivot-cache-refs-unreadable', message: `Pivot cache list could not be read from ${WORKBOOK}`, partPath: WORKBOOK })
    return out
  }
  const caches = getChildren(getChild(workbook, 'pivotCaches'), 'pivotCache')
  if (caches.length === 0) return out
  const targets = await relationships(pkg, WORKBOOK, '/pivotCacheDefinition')
  for (const cache of caches) {
    const a = attrs(cache)
    const id = a.cacheId !== undefined && /^\d+$/.test(a.cacheId) ? Number(a.cacheId) : undefined
    const target = a.id === undefined ? undefined : targets.get(a.id)
    if (id !== undefined && target && !out.has(id)) out.set(id, target)
  }
  return out
}

export async function loadPivotWorkbook(pkg: PackageReader, regions: ReadonlyMap<string, readonly PivotRegionRef[]>): Promise<PivotWorkbook> {
  const diagnostics: PivotDiagnostic[] = []
  const caches = new Map<number, PivotCacheDefinition>()
  const tables: PivotTableDefinition[] = []
  if (![...regions.values()].some(list => list.length)) return { caches, tables, diagnostics }
  const cacheParts = await workbookCaches(pkg, diagnostics)
  // One parse per cache part, shared by every table on it (§7.4).
  const parsed = new Map<string, Promise<PivotCacheDefinition | undefined>>()
  const loadCache = (partPath: string, cacheId: number) => {
    let pending = parsed.get(partPath)
    if (!pending) {
      pending = (async () => {
        let root: XmlNode | undefined
        try { root = await pkg.xmlOrdered(partPath) } catch { root = undefined }
        if (!root) {
          pushDiagnostic(diagnostics, { kind: 'parse', feature: 'pivot-cache-unreadable', message: `Pivot cache definition could not be read from ${partPath}`, partPath })
          return undefined
        }
        const cache = parseCacheDefinition(root, cacheId, partPath, diagnostics)
        const relId = attrs(root).id
        const records = relId === undefined ? undefined : (await relationships(pkg, partPath, '/pivotCacheRecords')).get(relId)
        if (records) cache.recordsPartPath = records
        return cache
      })()
      parsed.set(partPath, pending)
    }
    return pending
  }

  for (const [sheet, refs] of regions) {
    for (const region of refs) {
      const partPath = region.partPath
      let root: XmlNode | undefined
      try { root = await pkg.xml(partPath) } catch { root = undefined }
      if (!root) {
        pushDiagnostic(diagnostics, { kind: 'parse', feature: 'pivot-definition-unreadable', message: `Pivot table definition could not be read from ${partPath}`, partPath })
        continue
      }
      const table = parsePivotTable(root, { sheet, region }, diagnostics)
      if (!table) continue
      let cachePath = cacheParts.get(table.cacheId)
      if (!cachePath) {
        cachePath = [...(await relationships(pkg, partPath, '/pivotCacheDefinition')).values()][0]
        if (cachePath) pushDiagnostic(diagnostics, { kind: 'parse', feature: 'pivot-cache-mapping-recovered', message: `Pivot cache ${table.cacheId} is not in the workbook's pivotCaches; recovered from ${partPath}'s relationship`, partPath })
      }
      let cache: PivotCacheDefinition | undefined
      if (cachePath) {
        table.cachePartPath = cachePath
        cache = await loadCache(cachePath, table.cacheId)
        if (cache) {
          if (!caches.has(table.cacheId)) caches.set(table.cacheId, cache)
          if (cache.unsupported.length) table.unsupported = [...new Set([...table.unsupported, ...cache.unsupported.map(entry => `cache:${entry}`)])]
        }
      } else {
        pushDiagnostic(diagnostics, { kind: 'parse', feature: 'pivot-cache-unreadable', message: `Pivot cache ${table.cacheId} has no definition part (${partPath})`, partPath })
      }
      const invalid = validateIndices(table, cache)
      if (invalid.length) {
        table.unsupported = [...new Set([...table.unsupported, ...invalid])]
        pushDiagnostic(diagnostics, { kind: 'parse', feature: 'pivot-definition-invalid', message: `Pivot table ${partPath} has ${invalid.length} index reference(s) outside their field (first: ${invalid[0].slice('invalid-index:'.length)})`, partPath })
      }
      tables.push(table)
    }
  }
  return { caches, tables, diagnostics }
}

/** A resident package (`OfficePackage`): its zip entries, read on demand. The worker's `RangePackage` has none. */
export interface EntryPackage { entry(path: string): JSZip.JSZipObject | null }
const hasEntries = (pkg: PackageReader): pkg is PackageReader & EntryPackage => typeof (pkg as Partial<EntryPackage>).entry === 'function'
type StreamHelper = { on(event: 'data', fn: (chunk: Uint8Array) => void): StreamHelper; on(event: 'end', fn: () => void): StreamHelper; on(event: 'error', fn: (error: Error) => void): StreamHelper; pause(): StreamHelper; resume(): StreamHelper }

/** A zip entry's inflated bytes, chunk by chunk, paused while the consumer works (no whole-part copy, nothing cached). */
export async function* entryChunks(entry: JSZip.JSZipObject, signal: AbortSignal): AsyncIterable<Uint8Array> {
  const stream = (entry as unknown as { internalStream(type: 'uint8array'): StreamHelper }).internalStream('uint8array')
  const queue: Uint8Array[] = []
  let done = false, failure: Error | undefined, wake: (() => void) | undefined
  const notify = () => { const fn = wake; wake = undefined; fn?.() }
  stream.on('data', chunk => { queue.push(chunk); stream.pause(); notify() })
  stream.on('end', () => { done = true; notify() })
  stream.on('error', error => { failure = error; done = true; notify() })
  signal.addEventListener('abort', notify)
  try {
    stream.resume()
    for (;;) {
      if (signal.aborted) throw new PackageError('cancelled', 'operation cancelled')
      const chunk = queue.shift()
      if (chunk) { if (!queue.length && !done) stream.resume(); yield chunk; continue }
      if (failure) throw failure
      if (done) return
      await new Promise<void>(resolve => { wake = resolve })
    }
  } finally {
    signal.removeEventListener('abort', notify)
    if (!done) stream.pause()
  }
}

/**
 * Resident load: keep the workbook's caches (by cacheId and part path) and a records source over the package's zip
 * entries, in the runtime registered under `owners` (the document's sheets array and its sheets). Each records part's
 * entry is looked up now (no read); its bytes are inflated only when `records()` streams them. Undefined for a package
 * without zip entries (the worker's, Task 27).
 *
 * Memory: a JSZip entry keeps a view of the archive it came from, so a workbook with records keeps its compressed
 * file bytes alive for the document's lifetime (report, Task 25).
 */
export function retainPivotRuntime(owners: readonly object[], book: Pick<PivotWorkbook, 'caches'>, pkg: PackageReader, date1904: () => boolean): PivotRuntime | undefined {
  if (!hasEntries(pkg) || !book.caches.size) return undefined
  const entries = new Map<string, JSZip.JSZipObject>()
  for (const cache of book.caches.values()) {
    const path = cache.recordsPartPath
    if (path === undefined || entries.has(path)) continue
    const entry = pkg.entry(path)
    if (entry) entries.set(path, entry)
  }
  const source: RecordsSource = (path, signal) => { const entry = entries.get(path); return entry ? entryChunks(entry, signal) : undefined }
  const runtime = createPivotRuntime(book.caches, source, date1904)
  bindPivotRuntime(owners, runtime)
  return runtime
}
