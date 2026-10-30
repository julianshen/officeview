/**
 * `pivotCacheDefinition` → PivotCacheDefinition (spec §3). Lazy chunk `pivots`.
 *
 * Never throws on content: an unknown element, attribute, extension or enum
 * value is recorded in `unsupported` and the schema default stands. The tree
 * must come from the order-preserving parse (`xmlOrdered`): mixed shared items
 * (`<s/><m/><n/>`) are indexed by document position, which the grouped tree loses.
 */
import { attrs, elementChildren, getChild, getChildren, orderedChildren, type XmlNode } from '../../core/xml'
import type { PivotCacheDefinition, PivotCacheField, PivotCacheItem, PivotCacheSource, PivotDiagnostic, PivotFieldGroup, PivotSharedItems } from './model'

// ---------------------------------------------------------------- shared attribute readers (also used by parsePivotTable)

/** Reads one attribute; `bad` records a present value the schema type does not admit. */
export type Reader<T> = (raw: string | undefined, bad: () => void) => T
const parseBool = (raw: string): boolean | undefined => raw === '1' || raw === 'true' ? true : raw === '0' || raw === 'false' ? false : undefined
const parseInt32 = (raw: string): number | undefined => /^\s*-?\d+\s*$/.test(raw) && Number.isSafeInteger(Number(raw)) ? Number(raw) : undefined
const parseUint = (raw: string): number | undefined => { const n = parseInt32(raw); return n !== undefined && n >= 0 ? n : undefined }
const parseDouble = (raw: string): number | undefined => raw.trim() !== '' && Number.isFinite(Number(raw)) ? Number(raw) : undefined
const withDefault = <T>(parse: (raw: string) => T | undefined, d: T): Reader<T> => (raw, bad) => {
  if (raw === undefined) return d
  const v = parse(raw)
  if (v === undefined) bad()
  return v ?? d
}
const optional = <T>(parse: (raw: string) => T | undefined): Reader<T | undefined> => (raw, bad) => {
  if (raw === undefined) return undefined
  const v = parse(raw)
  if (v === undefined) bad()
  return v
}
export const bool = (d: boolean): Reader<boolean> => withDefault(parseBool, d)
export const optBool: Reader<boolean | undefined> = optional(parseBool)
export const uint = (d: number): Reader<number> => withDefault(parseUint, d)
export const optUint: Reader<number | undefined> = optional(parseUint)
export const int = (d: number): Reader<number> => withDefault(parseInt32, d)
export const optInt: Reader<number | undefined> = optional(parseInt32)
export const optDouble: Reader<number | undefined> = optional(parseDouble)
export const double = (d: number): Reader<number> => withDefault(parseDouble, d)
export const str: Reader<string | undefined> = raw => raw
export const oneOf = <T extends string>(values: readonly T[], d: T): Reader<T> => withDefault(raw => values.includes(raw as T) ? raw as T : undefined, d)
export const optOneOf = <T extends string>(values: readonly T[]): Reader<T | undefined> => optional(raw => values.includes(raw as T) ? raw as T : undefined)

/** Prefix-stripped bookkeeping attributes every writer may add: `mc:Ignorable`, `xr:uid` and friends. */
const BOOKKEEPING = new Set(['Ignorable', 'uid'])

/** Entries kept per part; a part with more (say, a million malformed items) ends with one truncation marker. */
export const MAX_UNSUPPORTED_ENTRIES = 256
export const UNSUPPORTED_TRUNCATED = 'unsupported(truncated)'

/** Deduplicated, bounded record of unmodelled constructs, plus the part's declared-count mismatches. */
export class Unsupported {
  private readonly seen = new Set<string>()
  /** First mismatch per element name (so bounded by the schema's element names). */
  readonly countMismatches = new Map<string, { declared: number; actual: number }>()
  countMismatch(el: string, declared: number, actual: number): void {
    if (!this.countMismatches.has(el)) this.countMismatches.set(el, { declared, actual })
  }
  add(entry: string): void {
    if (this.seen.size < MAX_UNSUPPORTED_ENTRIES) this.seen.add(entry)
    else if (!this.seen.has(entry)) this.seen.add(UNSUPPORTED_TRUNCATED)
  }
  list(): string[] { return [...this.seen] }
}

type Read<S> = { [K in keyof S]: S[K] extends Reader<infer T> ? T : never }
/** Reads `spec`'s attributes (absent optional ones stay absent) and records any other attribute not in `known`. */
export function readAttrs<S extends Record<string, Reader<unknown>>>(node: XmlNode, el: string, spec: S, u: Unsupported, known: readonly string[] = []): Read<S> {
  const a = attrs(node)
  const out: Record<string, unknown> = {}
  for (const key of Object.keys(spec)) {
    const raw = a[key]
    const value = spec[key](raw, () => u.add(`${el}@${key}=${String(raw).slice(0, 40)}`))
    if (value !== undefined) out[key] = value
  }
  for (const key of Object.keys(a)) if (!Object.prototype.hasOwnProperty.call(spec, key) && !known.includes(key) && !BOOKKEEPING.has(key)) u.add(`${el}@${key}`)
  return out as Read<S>
}

/** A declared `count` that disagrees with the `actual` children read is reported; the children win and the count is not kept. */
export function checkCount(node: XmlNode, el: string, actual: number, u: Unsupported): void {
  const raw = attrs(node).count
  if (raw === undefined) return
  const declared = parseUint(raw)
  if (declared === undefined) u.add(`${el}@count=${raw.slice(0, 40)}`)
  else if (declared !== actual) u.countMismatch(el, declared, actual)
}

/** A list container: its only attribute is `count`. */
export function readList(node: XmlNode, el: string, actual: number, u: Unsupported): void {
  readAttrs(node, el, {}, u, ['count'])
  checkCount(node, el, actual, u)
}

/** Adds a diagnostic unless one with the same feature and message is already there. */
export function pushDiagnostic(out: PivotDiagnostic[], d: PivotDiagnostic): void {
  if (!out.some(item => item.feature === d.feature && item.message === d.message)) out.push(d)
}

/** `pivot-count-mismatch` diagnostics for a part's recorded mismatches. */
export function countDiagnostics(u: Unsupported, partPath: string, out: PivotDiagnostic[]): void {
  for (const [el, { declared, actual }] of u.countMismatches) {
    pushDiagnostic(out, { kind: 'parse', feature: 'pivot-count-mismatch', message: `${el} in ${partPath} declares count ${declared} but has ${actual}`, partPath })
  }
}

/** Records child elements outside `known` (`extLst` is always known; callers read it through `readExtensions`). */
export function checkChildren(node: XmlNode, el: string, known: readonly string[], u: Unsupported): void {
  for (const [name] of elementChildren(node)) if (name !== 'extLst' && !known.includes(name)) u.add(`${el}/${name}`)
}

/** Dispatches `extLst/ext` by `@uri` (prefixes are stripped, so the URI is the only reliable key). Unknown URIs are recorded. */
export function readExtensions(node: XmlNode, handlers: Record<string, (ext: XmlNode) => void>, ignored: readonly string[], u: Unsupported): void {
  for (const ext of getChildren(getChild(node, 'extLst'), 'ext')) {
    const uri = attrs(ext).uri ?? ''
    const handler = handlers[uri]
    if (handler) handler(ext)
    else if (!ignored.includes(uri)) u.add(`ext${uri}`)
  }
}

/**
 * Indices from `<x v=".."/>` children. CT_Index (`discretePr`, pivot-area references) is an unsignedInt;
 * CT_X (`rowItems`/`colItems`) is an xsd:int defaulting to 0 per the schema, and the loader rejects a
 * negative one as an invalid index (`validateIndices`).
 */
export function readIndices(node: XmlNode | undefined, el: string, u: Unsupported, kind: 'index' | 'x' = 'index'): number[] {
  return getChildren(node, 'x').map(x => readAttrs(x, `${el}/x`, { v: kind === 'x' ? int(0) : uint(0) }, u).v)
}

// ---------------------------------------------------------------- cache definition

export const X14_PIVOT_CACHE_EXT = '{725AE2AE-9491-48be-B2B4-4EB974FC3084}'
/** xlpdsc `dataSourceClean`: a refresh hint with no effect on the saved cache. */
const DATA_SOURCE_CLEAN_EXT = '{F4336242-56EE-4DD3-A5F7-E900642BF6E9}'
const OLAP_CHILDREN = ['cacheHierarchies', 'kpis', 'tupleCache', 'calculatedMembers', 'dimensions', 'measureGroups', 'maps'] as const
const ITEM_TYPES = new Set(['m', 'n', 'b', 'e', 's', 'd'])

/** One `m`/`n`/`b`/`e`/`s`/`d` item. A malformed value still occupies its index (as missing) so positions stay aligned. */
function readItem(name: string, node: XmlNode, u: Unsupported): PivotCacheItem | undefined {
  if (!ITEM_TYPES.has(name)) { u.add(`item/${name}`); return undefined }
  const a = readAttrs(node, name, { v: str, u: bool(false), f: bool(false), c: str }, u)
  // `x` (member-property indices) and `tpls` (OLAP tuples) are not modelled.
  checkChildren(node, name, [], u)
  const raw = a.v
  if (name === 'm' ? raw !== undefined : raw === undefined) u.add(`${name}@v`)
  let item: PivotCacheItem = { type: 'missing' }
  if (name === 'n' || name === 'b') {
    const value = raw === undefined ? undefined : name === 'n' ? parseDouble(raw) : parseBool(raw)
    if (raw !== undefined && value === undefined) u.add(`${name}@v=${raw.slice(0, 40)}`)
    if (typeof value === 'number') item = { type: 'number', value }
    else if (typeof value === 'boolean') item = { type: 'boolean', value }
  } else if (name !== 'm') {
    item = { type: name === 'e' ? 'error' : name === 's' ? 'string' : 'date', value: raw ?? '' }
  }
  if (a.u) item.unused = true
  if (a.f) item.calculated = true
  if (a.c !== undefined) item.caption = a.c
  return item
}
function readItems(node: XmlNode, u: Unsupported): PivotCacheItem[] {
  const out: PivotCacheItem[] = []
  for (const [name, child] of orderedChildren(node)) {
    if (name === '#text' || name === 'extLst') continue
    const item = readItem(name, child, u)
    if (item) out.push(item)
  }
  return out
}

function readSharedItems(node: XmlNode, u: Unsupported): PivotSharedItems {
  const a = readAttrs(node, 'sharedItems', {
    containsSemiMixedTypes: bool(true), containsNonDate: bool(true), containsDate: bool(false), containsString: bool(true),
    containsBlank: bool(false), containsMixedTypes: bool(false), containsNumber: bool(false), containsInteger: bool(false),
    minValue: optDouble, maxValue: optDouble, minDate: str, maxDate: str, longText: bool(false),
  }, u, ['count'])
  const items = readItems(node, u)
  checkCount(node, 'sharedItems', items.length, u)
  return { ...a, inline: items.length === 0, items }
}

function readFieldGroup(node: XmlNode, u: Unsupported): PivotFieldGroup {
  const group: PivotFieldGroup = readAttrs(node, 'fieldGroup', { par: optUint, base: optUint }, u)
  checkChildren(node, 'fieldGroup', ['rangePr', 'discretePr', 'groupItems'], u)
  const range = getChild(node, 'rangePr')
  if (range) group.rangePr = readAttrs(range, 'rangePr', {
    autoStart: bool(true), autoEnd: bool(true),
    groupBy: oneOf(['range', 'seconds', 'minutes', 'hours', 'days', 'months', 'quarters', 'years'] as const, 'range'),
    startNum: optDouble, endNum: optDouble, startDate: str, endDate: str, groupInterval: double(1),
  }, u)
  const discrete = getChild(node, 'discretePr')
  if (discrete) {
    group.discretePr = readIndices(discrete, 'discretePr', u)
    readList(discrete, 'discretePr', group.discretePr.length, u)
  }
  const groupItems = getChild(node, 'groupItems')
  if (groupItems) {
    group.groupItems = readItems(groupItems, u)
    readList(groupItems, 'groupItems', group.groupItems.length, u)
  }
  return group
}

function readCacheField(node: XmlNode, u: Unsupported): PivotCacheField {
  const a = readAttrs(node, 'cacheField', {
    name: str, caption: str, propertyName: str, serverField: bool(false), uniqueList: bool(true), numFmtId: optUint,
    formula: str, sqlType: int(0), hierarchy: int(0), level: uint(0), databaseField: bool(true), mappingCount: optUint,
    memberPropertyField: bool(false),
  }, u)
  // mpMap (OLAP member-property maps) is not modelled.
  checkChildren(node, 'cacheField', ['sharedItems', 'fieldGroup'], u)
  readExtensions(node, {}, [], u)
  const field: PivotCacheField = { ...a, name: a.name ?? '' }
  const shared = getChild(node, 'sharedItems')
  if (shared) field.sharedItems = readSharedItems(shared, u)
  const group = getChild(node, 'fieldGroup')
  if (group) field.fieldGroup = readFieldGroup(group, u)
  return field
}

function readSource(node: XmlNode | undefined, u: Unsupported): PivotCacheSource {
  if (!node) {
    u.add('pivotCacheDefinition/cacheSource(missing)')
    return { type: 'worksheet', connectionId: 0 }
  }
  const a = readAttrs(node, 'cacheSource', { type: oneOf(['worksheet', 'external', 'consolidation', 'scenario'] as const, 'worksheet'), connectionId: uint(0) }, u)
  if (attrs(node).type === undefined) u.add('cacheSource@type(missing)')
  checkChildren(node, 'cacheSource', ['worksheetSource', 'consolidation'], u)
  readExtensions(node, {}, [], u)
  const source: PivotCacheSource = { ...a }
  const ws = getChild(node, 'worksheetSource')
  if (ws) {
    const w = readAttrs(ws, 'worksheetSource', { ref: str, name: str, sheet: str }, u, ['id'])
    const relId = attrs(ws).id
    source.worksheet = { ...w, ...(relId !== undefined ? { relId } : {}) }
  }
  // A consolidation source is refused as a whole (§7.6); its page and range sets are not read.
  const consolidation = getChild(node, 'consolidation')
  if (consolidation) source.consolidation = readAttrs(consolidation, 'consolidation', { autoPage: bool(true) }, u)
  return source
}

/**
 * Parses one cache definition. `cacheId` comes from the workbook's `pivotCaches`;
 * `recordsPartPath` is resolved by the loader from the root's `r:id`.
 */
export function parseCacheDefinition(root: XmlNode, cacheId: number, partPath: string, diagnostics: PivotDiagnostic[]): PivotCacheDefinition {
  const u = new Unsupported()
  const a = readAttrs(root, 'pivotCacheDefinition', {
    invalid: bool(false), saveData: bool(true), refreshOnLoad: bool(false), optimizeMemory: bool(false), enableRefresh: bool(true),
    refreshedBy: str, refreshedDate: optDouble, refreshedDateIso: str, backgroundQuery: bool(false), missingItemsLimit: optUint,
    createdVersion: uint(0), refreshedVersion: uint(0), minRefreshableVersion: uint(0), recordCount: optUint,
    upgradeOnRefresh: bool(false), tupleCache: bool(false), supportSubquery: bool(false), supportAdvancedDrill: bool(false),
  }, u, ['id'])
  // calculatedItems (calculated items, §5.6) is deliberately absent: it lands in `unsupported`.
  checkChildren(root, 'pivotCacheDefinition', ['cacheSource', 'cacheFields', ...OLAP_CHILDREN], u)
  const fieldsNode = getChild(root, 'cacheFields')
  if (!fieldsNode) u.add('pivotCacheDefinition/cacheFields(missing)')
  else readList(fieldsNode, 'cacheFields', getChildren(fieldsNode, 'cacheField').length, u)
  const { tupleCache, ...rest } = a
  const cache: PivotCacheDefinition = {
    cacheId, partPath, ...rest,
    source: readSource(getChild(root, 'cacheSource'), u),
    fields: getChildren(fieldsNode, 'cacheField').map(node => readCacheField(node, u)),
    olap: tupleCache || OLAP_CHILDREN.some(name => getChild(root, name) !== undefined),
    unsupported: [],
  }
  readExtensions(root, {
    [X14_PIVOT_CACHE_EXT]: ext => {
      const node = getChild(ext, 'pivotCacheDefinition')
      if (!node) return
      cache.x14 = readAttrs(node, 'x14:pivotCacheDefinition', {
        pivotCacheId: optUint, slicerData: bool(false), supportSubqueryNonVisual: bool(false), supportSubqueryCalcMem: bool(false), supportAddCalcMems: bool(false),
      }, u)
      checkChildren(node, 'x14:pivotCacheDefinition', [], u)
    },
  }, [DATA_SOURCE_CLEAN_EXT], u)
  cache.unsupported = u.list()
  countDiagnostics(u, partPath, diagnostics)
  return cache
}
