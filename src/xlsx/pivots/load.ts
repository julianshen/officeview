/**
 * Resolves a workbook's pivot parts into a PivotWorkbook (spec §3, §6.2). Lazy
 * chunk `pivots`: runs after `parseXlsx` on the resident path (never inside it)
 * and over `RangePackage` in the worker. Records are not read in Stage 1.
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
import { attrs, getChild, getChildren, type XmlNode } from '../../core/xml'
import { resolvePartTarget } from '../../drawing/parts'
import type { PivotCacheDefinition, PivotDiagnostic, PivotTableDefinition, PivotWorkbook } from './model'
import type { PivotRegionRef } from './stub'
import { parseCacheDefinition, pushDiagnostic } from './parseCacheDefinition'
import { parsePivotTable, validateIndices } from './parsePivotTable'

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
