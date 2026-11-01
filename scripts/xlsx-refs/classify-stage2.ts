/**
 * Classifies the Stage 2 engine cases as Stage 3 & Stage 4 interaction oracles (pivots plan Tasks 24 & 34, step 1):
 * tests/fixtures/xlsx-pivots/interaction/delta-from-stage2.json.
 *
 *   bun scripts/xlsx-refs/classify-stage2.ts            write the classification
 *   bun scripts/xlsx-refs/classify-stage2.ts --check    exit 1 when the file on disk differs from a fresh run
 *
 * 1. Same cache: a built case whose cache definition `<cacheFields>` and records part equal engine/baseline.xlsx's
 *    byte for byte (the seed refreshed by Excel). Only those can be replayed on the seed's records.
 * 2. For each same-cache case, the saved definitions of baseline and case are parsed with the viewer's loaders and
 *    diffed as view state: axis moves (including `PIVOT_VALUES_FIELD`), page selections, item visibility and collapse
 *    (matched by shared index `x`, since a sort reorders the `<items>` list), sortType/autoSortScope, filters, and
 *    data-field additions / value-field settings (`fn`, `showAs`, `baseField`, `baseItem`, `name`, `numFmtId`) are
 *    Stage 3 & 4 UI actions; layout/subtotal/table options and renamed pivot items are not.
 * 3. Cross-check against the case's own steps (engine-cases.ts `lines` / `xmlPrep`): a step that is not a Stage 3 or
 *    Stage 4 UI action must not produce a delta unless an override says why.
 * 4. `OVERRIDES` record the judgment calls, each with its reason.
 */
import { existsSync, readFileSync, writeFileSync } from 'node:fs'
import { join, relative } from 'node:path'
import { ENGINE_CASES, type Target } from './engine-cases'
import type { PivotDelta } from './interaction-cases'
import { CACHE_RE, PIVOT_DIR, RECORDS_RE, ROOT, loadPivot, loadZip, partNames, partText } from './pivot-files'
import {
  PIVOT_VALUES_FIELD, dataFieldSource,
  type PivotArea, type PivotAxis, type PivotCacheField, type PivotDataFunction, type PivotField, type PivotFieldItem,
  type PivotTableDefinition,
} from '../../src/xlsx/pivots/model'
import { ZH_TW_CAPTIONS } from '../../src/xlsx/pivots/session/captions'
import type { PivotValueShowAs } from '../../src/xlsx/pivots/session/delta'

export const OUT = join(PIVOT_DIR, 'interaction/delta-from-stage2.json')
const ENGINE_DIR = join(PIVOT_DIR, 'engine')

export type Classified = { case: string; delta: PivotDelta[]; note?: string } | { case: string; delta: null; reason: string; note?: string }

const S2_AA_VARIANCE_REASON = 'refused by S2-AA variance-total 15th-digit rule when totals are shown; covered on total-less base in Stage 4 manifest'

/** Judgment calls on top of the mechanical diff, each documented. */
const OVERRIDES: Record<string, { note: string; delta?: PivotDelta[] | null; reason?: string }> = {
  baseline: { note: 'identity: the seed refreshed with no change; an empty delta list is the trivial oracle (and the Reset target)' },
  'quirk-off-axis-hidden-item': {
    delta: [
      { kind: 'moveField', field: 6, axis: 'row', index: 2 },
      { kind: 'items', field: 6, hidden: [1] },
      { kind: 'moveField', field: 6, axis: 'none' },
    ],
    note: 'round-trip move: Channel moves onto rows, hides Retail (shared index 1), and moves back off every axis; saved definition equals baseline because moving off-axis clears hidden items',
  },
  'values-fn-stdev': { reason: S2_AA_VARIANCE_REASON, note: 'Stage 2 engine refuses stdDev when subtotals or grand totals are shown (ruling S2-AA)' },
  'values-fn-stdevp': { reason: S2_AA_VARIANCE_REASON, note: 'Stage 2 engine refuses stdDevp when subtotals or grand totals are shown (ruling S2-AA)' },
  'values-fn-var': { reason: S2_AA_VARIANCE_REASON, note: 'Stage 2 engine refuses var when subtotals or grand totals are shown (ruling S2-AA)' },
  'values-fn-varp': { reason: S2_AA_VARIANCE_REASON, note: 'Stage 2 engine refuses varp when subtotals or grand totals are shown (ruling S2-AA)' },
}

/** AppleScript steps that are Stage 3 or Stage 4 UI actions; any other step makes the case a non-UI change. */
const UI_STEP = /^(set show detail of pivot item|set visible of pivot item|set current page of|auto sort |make new pivot filter|auto show |clear (all|label|value) filters|clear manual filter|set pivot field orientation of |set position of |add data field |set function of |set calculation of |set base field of |set base item of |set number format of data field |set name of data field )/
const NON_UI_STEP_REASONS: Array<[RegExp, string]> = [
  [/orient as/, 'field move (orientation)'], [/row axis layout/, 'layout form'], [/subtotal location/, 'subtotal placement'],
  [/set subtotals/, 'named subtotals (field setting)'], [/set function of/, 'value field function'], [/set calculation of|base field|base item/, 'show values as'],
  [/show all items/, 'show items with no data (field setting)'], [/layout blank line/, 'blank rows (layout option)'], [/repeat labels/, 'repeat labels (layout option)'],
  [/merge labels/, 'merge labels (layout option)'], [/grid drop zones/, 'classic layout'], [/^set name of/, 'renamed field or item'],
  [/make new calculated/, 'calculated field or item'], [/^group /, 'grouping'], [/sort using custom lists/, 'custom-list sorting (table option)'],
  [/set (row|column) grand/, 'grand totals (layout option)'], [/set position of/, 'field move (position)'], [/range "|set formula/, 'source data edit'],
]

// ---- Same cache -------------------------------------------------------------------------------------------------

async function cacheKey(file: string): Promise<string> {
  const zip = await loadZip(file)
  const def = await partText(zip, partNames(zip, CACHE_RE)[0]) ?? ''
  const records = await partText(zip, partNames(zip, RECORDS_RE)[0]) ?? ''
  return `${/<cacheFields[\s\S]*<\/cacheFields>/.exec(def)?.[0] ?? ''}\n${records.replace(/^[\s\S]*?<pivotCacheRecords\b[^>]*>/, '')}`
}

// ---- View diff ---------------------------------------------------------------------------------------------------

const json = (v: unknown) => JSON.stringify(v)
/** Definition keys that never decide the classification: identity, the saved layout, versions, and the parts the
 * view diff reads separately. */
const DEF_IGNORED = new Set<keyof PivotTableDefinition>([
  'sheet', 'partPath', 'region', 'cacheId', 'cachePartPath', 'name', 'location', 'fields', 'rowFields', 'colFields',
  'pageFields', 'dataFields', 'dataOnRows', 'dataPosition', 'rowItems', 'colItems', 'filters', 'updatedVersion',
  'minRefreshableVersion', 'createdVersion', 'unsupported', 'chartFormats',
])
/** Field keys the UI diff reads (axis, dataField, page multi-select, items, sort) or that follow a value filter (measureFilter). */
const FIELD_IGNORED = new Set<keyof PivotField>(['axis', 'dataField', 'multipleItemSelectionAllowed', 'items', 'sortType', 'autoSortScope', 'measureFilter'])
const ITEM_UI = new Set<keyof PivotFieldItem>(['hidden', 'showDetails'])

const BASE_FIELD_MODES = new Set<PivotValueShowAs>([
  'difference', 'percent', 'percentDiff', 'runTotal',
  'percentOfParent', 'percentOfRunningTotal', 'rankAscending', 'rankDescending',
])
const BASE_ITEM_MODES = new Set<PivotValueShowAs>(['difference', 'percent', 'percentDiff'])
const PCT_AUTO_FMT_MODES = new Set<PivotValueShowAs>([
  'percent', 'percentDiff', 'percentOfRow', 'percentOfCol', 'percentOfTotal',
  'percentOfParent', 'percentOfParentRow', 'percentOfParentCol', 'percentOfRunningTotal',
])

function byData(scope: PivotArea | undefined): number | undefined {
  const ref = scope?.references.find(r => r.field === PIVOT_VALUES_FIELD || r.field === 4294967294)
  return ref?.x[0]
}

type UiAxis = 'row' | 'col' | 'page' | 'none'
function axisLocation(def: PivotTableDefinition, field: number): { axis: UiAxis; index: number } {
  const r = def.rowFields.indexOf(field)
  if (r >= 0) return { axis: 'row', index: r }
  const c = def.colFields.indexOf(field)
  if (c >= 0) return { axis: 'col', index: c }
  const p = def.pageFields.findIndex(page => page.fld === field)
  if (p >= 0) return { axis: 'page', index: p }
  return { axis: 'none', index: -1 }
}

function expectedFieldAxis(def: PivotTableDefinition, field: number): PivotAxis | undefined {
  const { axis } = axisLocation(def, field)
  return axis === 'row' ? 'axisRow' : axis === 'col' ? 'axisCol' : axis === 'page' ? 'axisPage' : undefined
}

function pageSelection(def: PivotTableDefinition, field: number): readonly number[] | 'all' | undefined {
  const page = def.pageFields.find(p => p.fld === field)
  if (!page) return undefined
  if (page.item !== undefined) return [page.item]
  const f = def.fields[field]
  if (f?.multipleItemSelectionAllowed && f.items?.some(i => i.type === 'data' && i.hidden)) {
    return f.items.flatMap((item, k) => item.type === 'data' && !item.hidden ? [k] : [])
  }
  return 'all'
}

function defaultDataFn(cacheField: PivotCacheField | undefined): PivotDataFunction {
  const s = cacheField?.sharedItems
  return s && s.containsNumber && !s.containsString && !s.containsDate ? 'sum' : 'count'
}

function autoDataName(fn: PivotDataFunction, fieldName: string): string | undefined {
  return ZH_TW_CAPTIONS[`dataField:${fn}`]?.replace('{field}', fieldName)
}

export function diffDefinitions(
  base: PivotTableDefinition,
  next: PivotTableDefinition,
  cacheFields: readonly PivotCacheField[] = [],
): { delta: PivotDelta[]; other: string[] } {
  const other: string[] = []
  for (const key of Object.keys({ ...base, ...next }) as Array<keyof PivotTableDefinition>) {
    if (DEF_IGNORED.has(key)) continue
    if (json(base[key]) !== json(next[key])) other.push(String(key))
  }
  if (base.dataPosition !== next.dataPosition) other.push('dataPosition')
  const expectedDataOnRows = next.rowFields.includes(PIVOT_VALUES_FIELD) ? true : next.colFields.includes(PIVOT_VALUES_FIELD) ? false : base.dataOnRows
  if (next.dataOnRows !== expectedDataOnRows) other.push('dataOnRows')

  const delta: PivotDelta[] = []
  const count = Math.max(base.fields.length, next.fields.length)

  // 1. Axis & Values moves (rowFields, colFields, pageFields).
  const candidates: number[] = []
  for (let f = 0; f < count; f++) candidates.push(f)
  if (base.rowFields.includes(PIVOT_VALUES_FIELD) || base.colFields.includes(PIVOT_VALUES_FIELD)
    || next.rowFields.includes(PIVOT_VALUES_FIELD) || next.colFields.includes(PIVOT_VALUES_FIELD)) {
    candidates.push(PIVOT_VALUES_FIELD)
  }
  const movedToAxis = new Map<UiAxis, Array<{ field: number; index: number }>>()
  for (const f of candidates) {
    const before = axisLocation(base, f), after = axisLocation(next, f)
    if (before.axis === after.axis) continue
    // Values (-2) appearing/disappearing solely from 1 <-> 2 data fields is driven by add/removeDataField, not moveField.
    if (f === PIVOT_VALUES_FIELD && (before.axis === 'none' || after.axis === 'none')) continue
    if (after.axis === 'none') delta.push({ kind: 'moveField', field: f, axis: 'none' })
    else {
      const list = movedToAxis.get(after.axis) ?? []
      list.push({ field: f, index: after.index })
      movedToAxis.set(after.axis, list)
    }
  }
  for (const [axis, list] of movedToAxis) {
    for (const entry of list.sort((a, b) => a.index - b.index)) {
      delta.push({ kind: 'moveField', field: entry.field, axis: axis as 'row' | 'col' | 'page', index: entry.index })
    }
  }
  for (const [axisName, aList, bList] of [
    ['rowFields', base.rowFields, next.rowFields],
    ['colFields', base.colFields, next.colFields],
    ['pageFields', base.pageFields.map(p => p.fld), next.pageFields.map(p => p.fld)],
  ] as const) {
    const keptA = aList.filter(x => bList.includes(x))
    const keptB = bList.filter(x => aList.includes(x))
    if (json(keptA) !== json(keptB)) other.push(`${axisName} (reordered)`)
  }

  // 2. Page selections.
  const multiPageFields = new Set<number>()
  for (const page of next.pageFields) {
    const beforeSel = pageSelection(base, page.fld) ?? 'all'
    const afterSel = pageSelection(next, page.fld) ?? 'all'
    if (next.fields[page.fld]?.multipleItemSelectionAllowed && page.item === undefined) multiPageFields.add(page.fld)
    if (json(beforeSel) !== json(afterSel)) delta.push({ kind: 'page', field: page.fld, selected: afterSel })
  }

  // 3. Per-field flags, item visibility, collapse, and sort.
  const nextDataSources = new Set(next.dataFields.map(dataFieldSource))
  for (let f = 0; f < count; f++) {
    const a = base.fields[f], b = next.fields[f]
    if (!a || !b) { other.push(`fields[${f}] added or removed`); continue }
    for (const key of Object.keys({ ...a, ...b }) as Array<keyof PivotField>) {
      if (FIELD_IGNORED.has(key)) continue
      if (json(a[key]) !== json(b[key])) other.push(`fields[${f}].${key}`)
    }
    if (a.axis !== b.axis && b.axis !== expectedFieldAxis(next, f)) other.push(`fields[${f}].axis`)
    if (a.dataField !== b.dataField && b.dataField !== nextDataSources.has(f)) other.push(`fields[${f}].dataField`)
    if (a.multipleItemSelectionAllowed !== b.multipleItemSelectionAllowed
      && !(b.multipleItemSelectionAllowed && next.pageFields.some(p => p.fld === f))) {
      other.push(`fields[${f}].multipleItemSelectionAllowed`)
    }
    const aItems = a.items ?? [], bItems = b.items ?? []
    const dataX = (items: readonly PivotFieldItem[]) => items.filter(i => i.type === 'data').map(i => i.x)
    // Items match by shared index; anything but hidden/sd differing (captions, type, the set of items) is not UI.
    const byX = new Map(bItems.filter(i => i.type === 'data').map(i => [i.x, i]))
    if (json([...dataX(aItems)].sort()) !== json([...dataX(bItems)].sort()) || json(aItems.filter(i => i.type !== 'data').map(i => i.type)) !== json(bItems.filter(i => i.type !== 'data').map(i => i.type))) {
      other.push(`fields[${f}].items (set)`)
      continue
    }
    for (const item of aItems) {
      if (item.type !== 'data') continue
      const twin = byX.get(item.x)!
      for (const key of Object.keys({ ...item, ...twin }) as Array<keyof PivotFieldItem>) {
        if (!ITEM_UI.has(key) && json(item[key]) !== json(twin[key])) other.push(`fields[${f}].items[x=${item.x}].${key}`)
      }
    }
    const sortChanged = a.sortType !== b.sortType || json(a.autoSortScope) !== json(b.autoSortScope)
    // Without a sort change, the list order is the manual order: a reorder is a manual move (not Stage 3).
    if (!sortChanged && json(dataX(aItems)) !== json(dataX(bItems))) other.push(`fields[${f}].items (order)`)
    if (!multiPageFields.has(f)) {
      const hiddenBefore = aItems.flatMap((item, k) => item.hidden ? [k] : [])
      const hiddenAfter = aItems.flatMap((item, k) => item.type === 'data' && byX.get(item.x)!.hidden ? [k] : [])
      if (json(hiddenBefore) !== json(hiddenAfter)) delta.push({ kind: 'items', field: f, hidden: hiddenAfter })
    }
    const changed = aItems.flatMap((item, k) => item.type === 'data' && item.showDetails !== byX.get(item.x)!.showDetails ? [k] : [])
    const dataCount = aItems.filter(i => i.type === 'data').length
    if (changed.length && changed.length === dataCount && changed.every(k => byX.get(aItems[k]!.x)!.showDetails === byX.get(aItems[changed[0]!]!.x)!.showDetails)) {
      delta.push({ kind: 'collapseField', field: f, collapsed: !byX.get(aItems[changed[0]!]!.x)!.showDetails })
    } else {
      for (const k of changed) delta.push({ kind: 'collapse', field: f, item: k, collapsed: !byX.get(aItems[k]!.x)!.showDetails })
    }
    if (sortChanged) {
      if (b.sortType === 'manual') other.push(`fields[${f}].sortType (back to manual)`)
      else {
        const data = byData(b.autoSortScope)
        delta.push({ kind: 'sort', field: f, order: b.sortType, ...(data !== undefined ? { byData: data } : {}) })
      }
    }
  }

  // 4. Filters: removed ones clear their field; new or changed ones apply.
  const key = (filter: PivotTableDefinition['filters'][number]) => json({ ...filter, id: 0, evalOrder: 0 })
  const before = new Map(base.filters.map(filter => [key(filter), filter])), after = new Set(next.filters.map(key))
  for (const filter of base.filters) if (!after.has(key(filter)) && !next.filters.some(n => n.fld === filter.fld)) delta.push({ kind: 'clearFilters', field: filter.fld })
  for (const filter of next.filters) if (!before.has(key(filter))) delta.push({ kind: 'filter', field: filter.fld, filter })

  // 5. Data fields & value settings.
  const prefixMatches = next.dataFields.length >= base.dataFields.length
    && base.dataFields.every((a, i) => dataFieldSource(a) === dataFieldSource(next.dataFields[i]!))
  if (!prefixMatches) {
    other.push('dataFields')
  } else {
    for (let i = 0; i < next.dataFields.length; i++) {
      const b = next.dataFields[i]!
      const src = dataFieldSource(b)
      const fieldName = cacheFields[src]?.name ?? ''
      let initFn: PivotDataFunction
      let initName: string | undefined
      let initShowAs: PivotValueShowAs
      let initBaseField: number
      let initBaseItem: number
      let initNumFmtId: number | undefined
      if (i < base.dataFields.length) {
        const a = base.dataFields[i]!
        initFn = a.subtotal
        initName = a.name
        initShowAs = a.x14?.pivotShowAs ?? a.showDataAs
        initBaseField = a.baseField
        initBaseItem = a.baseItem
        initNumFmtId = a.numFmtId
      } else {
        delta.push({ kind: 'addDataField', field: src, index: i })
        initFn = defaultDataFn(cacheFields[src])
        initName = autoDataName(initFn, fieldName)
        initShowAs = 'normal'
        initBaseField = 0
        initBaseItem = 0
        initNumFmtId = undefined
      }
      const bShowAs: PivotValueShowAs = b.x14?.pivotShowAs ?? b.showDataAs
      const bAutoName = autoDataName(b.subtotal, fieldName)
      const fnChanged = initFn !== b.subtotal || (initName !== b.name && b.name !== undefined && b.name === bAutoName)
      const showAsChanged = initShowAs !== bShowAs
      const baseFieldChanged = BASE_FIELD_MODES.has(bShowAs) && (showAsChanged || initBaseField !== b.baseField)
      const baseItemChanged = BASE_ITEM_MODES.has(bShowAs) && (showAsChanged || initBaseItem !== b.baseItem)
      const nameChanged = fnChanged ? b.name !== bAutoName : b.name !== initName
      const autoPctFmt = showAsChanged && PCT_AUTO_FMT_MODES.has(bShowAs) && initNumFmtId === undefined && b.numFmtId === 10
      const numFmtChanged = initNumFmtId !== b.numFmtId && !autoPctFmt
      if (fnChanged || showAsChanged || baseFieldChanged || baseItemChanged || nameChanged || numFmtChanged) {
        delta.push({
          kind: 'valueSettings',
          dataIndex: i,
          ...(fnChanged ? { fn: b.subtotal } : {}),
          ...(showAsChanged ? { showAs: bShowAs } : {}),
          ...(baseFieldChanged ? { baseField: b.baseField } : {}),
          ...(baseItemChanged ? { baseItem: b.baseItem } : {}),
          ...(nameChanged && b.name !== undefined ? { name: b.name } : {}),
          ...(numFmtChanged ? { numFmtId: b.numFmtId ?? null } : {}),
        })
      }
    }
  }

  return { delta, other }
}

/** The non-UI steps of a case's AppleScript (`lines` and `xmlPrep`), as reasons. */
function nonUiSteps(id: string): string[] {
  const c = ENGINE_CASES.find(e => e.id === id)!
  const t: Target = { p: '<p>', ws: '<ws>', wb: '<wb>', data: '<data>' }
  const steps = [...(c.lines?.(t) ?? []), ...(c.xmlPrep?.(t) ?? [])]
  const reasons = new Set<string>()
  for (const step of steps) {
    if (UI_STEP.test(step)) continue
    const hit = NON_UI_STEP_REASONS.find(([re]) => re.test(step))
    reasons.add(hit ? hit[1] : `step "${step.slice(0, 60)}"`)
  }
  return [...reasons]
}

export async function classify(): Promise<Classified[]> {
  const manifest = JSON.parse(readFileSync(join(ENGINE_DIR, 'manifest.json'), 'utf8')) as Array<{ case: string; status: string; route: string }>
  const baseFile = join(ENGINE_DIR, 'baseline.xlsx')
  const baseKey = await cacheKey(baseFile)
  const { def: base, cache: baseCache } = await loadPivot(baseFile)
  const out: Classified[] = []
  for (const entry of manifest) {
    const file = join(ENGINE_DIR, `${entry.case}.xlsx`)
    if (entry.status !== 'built' || !existsSync(file) || await cacheKey(file) !== baseKey) continue
    const override = OVERRIDES[entry.case]
    const { delta, other } = diffDefinitions(base, (await loadPivot(file)).def, baseCache.fields)
    const steps = nonUiSteps(entry.case)
    if (override && 'reason' in override) { out.push({ case: entry.case, delta: null, reason: override.reason!, note: override.note }); continue }
    if (override && Array.isArray(override.delta) && !other.length && !steps.length) {
      out.push({ case: entry.case, delta: override.delta, note: override.note })
      continue
    }
    const viaXml = entry.route === 'xml'
    if (other.length || steps.length) {
      const reason = [...steps, ...(viaXml ? ['built through the xml route'] : []), ...(other.length ? [`saved definition differs in ${other.join(', ')}`] : [])].join('; ')
      out.push({ case: entry.case, delta: null, reason, ...(override ? { note: override.note } : {}) })
      continue
    }
    const note = override?.note ?? (viaXml ? 'built through the xml route (the AppleScript route was refused or ignored); the change is the UI dialog\'s, and the delta carries the saved setting as Excel refreshed it' : undefined)
    out.push({ case: entry.case, delta, ...(note ? { note } : {}) })
  }
  return out
}

if ((import.meta as { main?: boolean }).main) {
  const result = await classify()
  const text = `${JSON.stringify(result, null, 2)}\n`
  const withDelta = result.filter(r => r.delta !== null).length
  if (process.argv.includes('--check')) {
    const same = existsSync(OUT) && readFileSync(OUT, 'utf8') === text
    console.log(same ? `${relative(ROOT, OUT)} is current` : `${relative(ROOT, OUT)} is stale: rerun without --check`)
    process.exit(same ? 0 : 1)
  }
  writeFileSync(OUT, text)
  console.log(`${result.length} same-cache cases: ${withDelta} with a delta, ${result.length - withDelta} null; wrote ${relative(ROOT, OUT)}`)
}
