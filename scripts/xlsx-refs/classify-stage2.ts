/**
 * Classifies the Stage 2 engine cases as Stage 3 interaction oracles (pivots plan Task 24, step 1):
 * tests/fixtures/xlsx-pivots/interaction/delta-from-stage2.json.
 *
 *   bun scripts/xlsx-refs/classify-stage2.ts            write the classification
 *   bun scripts/xlsx-refs/classify-stage2.ts --check    exit 1 when the file on disk differs from a fresh run
 *
 * 1. Same cache: a built case whose cache definition `<cacheFields>` and records part equal engine/baseline.xlsx's
 *    byte for byte (the seed refreshed by Excel). Only those can be replayed on the seed's records.
 * 2. For each same-cache case, the saved definitions of baseline and case are parsed with the viewer's loaders and
 *    diffed as view state: item visibility and collapse (matched by shared index `x`, since a sort reorders the
 *    `<items>` list), sortType/autoSortScope, and filters are Stage 3 UI actions; everything else (axes, data fields,
 *    field and table settings, item captions, item order without a sort) is not. A case whose diff is confined to
 *    the first group gets `delta` (indices into baseline's lists, as `initialViewState` indexes them); any other
 *    difference gives `delta: null` with the differing properties as the reason.
 * 3. Cross-check against the case's own steps (engine-cases.ts `lines` / `xmlPrep`): a step that is not a Stage 3 UI
 *    action (a field move, a layout or value setting) must not produce a delta unless an override says why.
 * 4. `OVERRIDES` record the judgment calls, each with its reason.
 */
import { existsSync, readFileSync, writeFileSync } from 'node:fs'
import { join, relative } from 'node:path'
import { ENGINE_CASES, type Target } from './engine-cases'
import type { PivotDelta } from './interaction-cases'
import { CACHE_RE, PIVOT_DIR, RECORDS_RE, ROOT, loadPivot, loadZip, partNames, partText } from './pivot-files'
import { PIVOT_VALUES_FIELD, type PivotArea, type PivotField, type PivotFieldItem, type PivotTableDefinition } from '../../src/xlsx/pivots/model'

export const OUT = join(PIVOT_DIR, 'interaction/delta-from-stage2.json')
const ENGINE_DIR = join(PIVOT_DIR, 'engine')

export type Classified = { case: string; delta: PivotDelta[]; note?: string } | { case: string; delta: null; reason: string; note?: string }

/** Judgment calls on top of the mechanical diff, each documented. */
const OVERRIDES: Record<string, { note: string; delta?: null; reason?: string }> = {
  baseline: { note: 'identity: the seed refreshed with no change; an empty delta list is the trivial oracle (and the Reset target)' },
}

/** AppleScript steps that are Stage 3 UI actions; any other step makes the case a non-UI change. */
const UI_STEP = /^(set show detail of pivot item|set visible of pivot item|set current page of|auto sort |make new pivot filter|auto show |clear (all|label|value) filters|clear manual filter)/
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
const DEF_IGNORED = new Set<keyof PivotTableDefinition>(['sheet', 'partPath', 'region', 'cacheId', 'cachePartPath', 'name', 'location', 'fields', 'rowItems', 'colItems',
  'filters', 'updatedVersion', 'minRefreshableVersion', 'createdVersion', 'unsupported', 'chartFormats'])
/** Field keys the UI diff reads (items, sort) or that follow a value filter (measureFilter). */
const FIELD_IGNORED = new Set<keyof PivotField>(['items', 'sortType', 'autoSortScope', 'measureFilter'])
const ITEM_UI = new Set<keyof PivotFieldItem>(['hidden', 'showDetails'])

function byData(scope: PivotArea | undefined): number | undefined {
  const ref = scope?.references.find(r => r.field === PIVOT_VALUES_FIELD || r.field === 4294967294)
  return ref?.x[0]
}

export function diffDefinitions(base: PivotTableDefinition, next: PivotTableDefinition): { delta: PivotDelta[]; other: string[] } {
  const other: string[] = []
  for (const key of Object.keys({ ...base, ...next }) as Array<keyof PivotTableDefinition>) {
    if (DEF_IGNORED.has(key)) continue
    if (json(base[key]) !== json(next[key])) other.push(String(key))
  }
  const delta: PivotDelta[] = []
  const count = Math.max(base.fields.length, next.fields.length)
  for (let f = 0; f < count; f++) {
    const a = base.fields[f], b = next.fields[f]
    if (!a || !b) { other.push(`fields[${f}] added or removed`); continue }
    for (const key of Object.keys({ ...a, ...b }) as Array<keyof PivotField>) {
      if (FIELD_IGNORED.has(key)) continue
      if (json(a[key]) !== json(b[key])) other.push(`fields[${f}].${key}`)
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
    const hiddenBefore = aItems.flatMap((item, k) => item.hidden ? [k] : [])
    const hiddenAfter = aItems.flatMap((item, k) => item.type === 'data' && byX.get(item.x)!.hidden ? [k] : [])
    if (json(hiddenBefore) !== json(hiddenAfter)) delta.push({ kind: 'items', field: f, hidden: hiddenAfter })
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
  // Filters: removed ones clear their field; new or changed ones apply.
  const key = (filter: PivotTableDefinition['filters'][number]) => json({ ...filter, id: 0, evalOrder: 0 })
  const before = new Map(base.filters.map(filter => [key(filter), filter])), after = new Set(next.filters.map(key))
  for (const filter of base.filters) if (!after.has(key(filter)) && !next.filters.some(n => n.fld === filter.fld)) delta.push({ kind: 'clearFilters', field: filter.fld })
  for (const filter of next.filters) if (!before.has(key(filter))) delta.push({ kind: 'filter', field: filter.fld, filter })
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
  const base = (await loadPivot(baseFile)).def
  const out: Classified[] = []
  for (const entry of manifest) {
    const file = join(ENGINE_DIR, `${entry.case}.xlsx`)
    if (entry.status !== 'built' || !existsSync(file) || await cacheKey(file) !== baseKey) continue
    const override = OVERRIDES[entry.case]
    const { delta, other } = diffDefinitions(base, (await loadPivot(file)).def)
    const steps = nonUiSteps(entry.case)
    if (override && 'reason' in override) { out.push({ case: entry.case, delta: null, reason: override.reason!, note: override.note }); continue }
    const viaXml = entry.route === 'xml'
    if (other.length || steps.length) {
      const reason = [...steps, ...(viaXml ? ['built through the xml route'] : []), ...(other.length ? [`saved definition differs in ${other.join(', ')}`] : [])].join('; ')
      out.push({ case: entry.case, delta: null, reason, ...(override ? { note: override.note } : {}) })
      continue
    }
    const note = override?.note ?? (viaXml ? 'built through the xml route (the AppleScript route was refused or ignored); the change is the UI dialog\'s, and the delta carries the filter as Excel saved it' : undefined)
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
