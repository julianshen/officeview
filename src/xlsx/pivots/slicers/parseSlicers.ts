/**
 * Slicer parts and slicer caches (MS-XLSX x14 `slicers`, `slicerCacheDefinition`)
 * into a SlicerModel (spec §3, Task 12). Lazy chunk `pivots`.
 *
 * - `xl/slicers/slicerN.xml`: one `<slicer>` per control (name, cache, caption,
 *   startItem, columnCount, showCaption, style, rowHeight).
 * - `xl/slicerCaches/slicerCacheN.xml`: the pivot tables it drives and, for a
 *   pivot slicer, `data/tabular` (pivotCacheId, crossFilter, showMissing, items
 *   `x` / `s` selected / `nd` no data). Item labels are the pivot cache field's
 *   shared items, the field named by `sourceName`, the cache found by its x14
 *   `pivotCacheId` (not the workbook `cacheId`).
 * - OLAP caches (`data/olap`) and table slicers (x15 `tableSlicerCache`, Stage 5)
 *   fail closed: the drawing keeps its Fallback, with a diagnostic.
 *
 * Bounded: at most MAX_SLICER_ITEMS items per cache and MAX_SLICERS_PER_PART
 * slicers per part; never throws (a malformed part reads as absent).
 */
import { attrs, getChild, getChildren, type XmlNode } from '../../../core/xml'
import type { PivotCacheDefinition, PivotCacheItem } from '../model'
import type { ControlDiagnostic, SlicerAppearance, SlicerItem, SlicerModel } from './model'

export const MAX_SLICER_ITEMS = 10_000
export const MAX_SLICERS_PER_PART = 256
/** Excel's own default for a new slicer (recipe R1 writes 209550 EMU = 16.5 pt). */
export const DEFAULT_SLICER_ROW_HEIGHT_EMU = 209_550

const off = (v: string | undefined) => v === '0' || v === 'false'
const on = (v: string | undefined) => v === '1' || v === 'true'
const uint = (v: string | undefined): number | undefined => v !== undefined && /^\d+$/.test(v.trim()) ? Number(v) : undefined
const root = (node: XmlNode | undefined, name: string) => getChild(node, name) ?? node

export interface SlicerPart { name: string; cache: string; caption?: string; startItem: number; columnCount: number; showCaption: boolean; style?: string; rowHeightEmu: number }

export function parseSlicerPart(node: XmlNode | undefined): SlicerPart[] {
  const out: SlicerPart[] = []
  for (const slicer of getChildren(root(node, 'slicers'), 'slicer').slice(0, MAX_SLICERS_PER_PART)) {
    const a = attrs(slicer)
    if (!a.name || !a.cache) continue
    const columns = uint(a.columnCount), rowHeight = uint(a.rowHeight)
    out.push({
      name: a.name, cache: a.cache,
      ...(a.caption !== undefined ? { caption: a.caption } : {}),
      startItem: uint(a.startItem) ?? 0,
      columnCount: Math.min(64, Math.max(1, columns ?? 1)),
      showCaption: !off(a.showCaption),
      ...(a.style ? { style: a.style } : {}),
      rowHeightEmu: rowHeight && rowHeight > 0 ? Math.min(rowHeight, 12_700 * 409) : DEFAULT_SLICER_ROW_HEIGHT_EMU,
    })
  }
  return out
}

export type SlicerCrossFilter = 'none' | 'showItemsWithDataAtTop' | 'showItemsWithNoData'
export interface SlicerCacheItemRef { x: number; selected: boolean; noData: boolean }
export interface SlicerCacheDef {
  name: string
  sourceName: string
  pivotTables: string[]
  kind: 'tabular' | 'olap' | 'table' | 'none'
  tabular?: { pivotCacheId?: number; crossFilter: SlicerCrossFilter; showMissing: boolean; items: SlicerCacheItemRef[]; truncated: boolean; invalid: boolean }
}

export function parseSlicerCache(node: XmlNode | undefined): SlicerCacheDef | undefined {
  const def = root(node, 'slicerCacheDefinition')
  const a = attrs(def)
  if (!def || !a.name) return undefined
  const pivotTables = getChildren(getChild(def, 'pivotTables'), 'pivotTable').map(table => attrs(table).name).filter((name): name is string => !!name)
  const base = { name: a.name, sourceName: a.sourceName ?? '', pivotTables }
  const data = getChild(def, 'data')
  const tabular = getChild(data, 'tabular')
  if (tabular) {
    const t = attrs(tabular)
    const list = getChildren(getChild(tabular, 'items'), 'i')
    let invalid = false
    const items = list.slice(0, MAX_SLICER_ITEMS).map(item => {
      const ia = attrs(item), x = uint(ia.x)
      if (x === undefined) invalid = true
      return { x: x ?? -1, selected: on(ia.s), noData: on(ia.nd) }
    })
    const crossFilter: SlicerCrossFilter = t.crossFilter === 'none' || t.crossFilter === 'showItemsWithNoData' ? t.crossFilter : 'showItemsWithDataAtTop'
    return { ...base, kind: 'tabular', tabular: { ...(uint(t.pivotCacheId) !== undefined ? { pivotCacheId: uint(t.pivotCacheId) } : {}), crossFilter, showMissing: !off(t.showMissing), items, truncated: list.length > MAX_SLICER_ITEMS, invalid } }
  }
  if (getChild(data, 'olap')) return { ...base, kind: 'olap' }
  const table = getChildren(getChild(def, 'extLst'), 'ext').some(ext => getChild(ext, 'tableSlicerCache'))
  return { ...base, kind: table ? 'table' : 'none' }
}

const MONTH_DAY = (iso: string) => {
  const m = /^(\d{4})-(\d{2})-(\d{2})/.exec(iso)
  return m ? `${Number(m[2])}/${Number(m[3])}/${m[1]}` : iso
}
/** An item's label as Excel shows it in a slicer: its caption, else its value (dates as m/d/yyyy). */
export function itemLabel(item: PivotCacheItem): string {
  if (item.caption !== undefined) return item.caption
  switch (item.type) {
    case 'string': case 'error': return item.value
    case 'number': return String(Number(item.value.toPrecision(15)))
    case 'boolean': return item.value ? 'TRUE' : 'FALSE'
    case 'date': return MONTH_DAY(item.value)
    default: return '(blank)'
  }
}

/** The slicer's model, or the diagnostics that keep its drawing on the Fallback. */
export function buildSlicerModel(part: SlicerPart, cache: SlicerCacheDef | undefined, pivotCaches: ReadonlyMap<number, PivotCacheDefinition>, style: { style: string; appearance: SlicerAppearance }): { model?: SlicerModel; diagnostics: ControlDiagnostic[] } {
  const fail = (feature: string, message: string) => ({ diagnostics: [{ kind: 'unsupported' as const, feature, message: `Slicer ${part.name}: ${message}; its fallback shape is shown` }] })
  if (!cache) return fail('slicer-cache-missing', `slicer cache ${part.cache} is missing or unreadable`)
  if (cache.kind === 'olap') return fail('slicer-cache-unsupported', `OLAP slicer cache ${cache.name} is not supported`)
  if (cache.kind === 'table') return fail('slicer-cache-unsupported', `table slicer cache ${cache.name} is not supported yet`)
  if (cache.kind !== 'tabular' || !cache.tabular) return fail('slicer-cache-unsupported', `slicer cache ${cache.name} has no item data`)
  const tabular = cache.tabular
  const pivot = tabular.pivotCacheId === undefined ? undefined : pivotCaches.get(tabular.pivotCacheId)
  if (!pivot) return fail('slicer-pivot-cache-missing', `pivot cache ${tabular.pivotCacheId ?? '(none)'} of slicer cache ${cache.name} is missing`)
  if (pivot.olap) return fail('slicer-cache-unsupported', `slicer cache ${cache.name} reads an OLAP pivot cache`)
  const field = pivot.fields.find(candidate => candidate.name === cache.sourceName)
  const shared = field?.sharedItems?.items.length ? field.sharedItems.items : field?.fieldGroup?.groupItems ?? []
  if (!field || !shared.length) return fail('slicer-item-invalid', `field ${cache.sourceName} has no items in pivot cache ${tabular.pivotCacheId}`)
  if (tabular.invalid || tabular.items.some(item => item.x < 0 || item.x >= shared.length)) return fail('slicer-item-invalid', `an item index lies outside field ${cache.sourceName}`)
  const diagnostics: ControlDiagnostic[] = []
  if (tabular.truncated) diagnostics.push({ kind: 'unsupported', feature: 'slicer-items-truncated', message: `Slicer ${part.name}: only the first ${MAX_SLICER_ITEMS} items are shown` })
  // Cross-filtering (spec §7.2): no-data items are marked unless crossFilter="none", hidden by showMissing="0",
  // and listed after the items with data by the default showItemsWithDataAtTop.
  let items: SlicerItem[] = tabular.items
    .map(item => ({ label: itemLabel(shared[item.x]!), selected: item.selected, noData: item.noData && tabular.crossFilter !== 'none' }))
    .filter(item => tabular.showMissing || !item.noData)
  if (tabular.crossFilter === 'showItemsWithDataAtTop') items = [...items.filter(item => !item.noData), ...items.filter(item => item.noData)]
  return {
    diagnostics,
    model: {
      name: part.name,
      caption: part.caption ?? (cache.sourceName || part.name),
      showCaption: part.showCaption,
      columnCount: part.columnCount,
      rowHeightEmu: part.rowHeightEmu,
      style: style.style,
      startItem: part.startItem,
      items,
      drives: cache.pivotTables.map(pivotTable => ({ pivotTable })),
      appearance: style.appearance,
    },
  }
}
