/** Pure Task 37 cache/definition/view transitions for measured Excel grouping. */
import type { PivotCacheDefinition, PivotCacheItem, PivotField, PivotFieldItem } from '../model'
import { VALUE_TAG, isoToSerial } from '../engine/records'
import type { EngineResult } from '../engine/axes'
import type { PivotDelta, PivotDeltaContext, PivotDeltaStateResult, PivotMemoryHold, PivotDateGroupPeriod } from './delta'
import type { PivotMenuGroupInfo } from './menu'
import type { PivotViewState, FieldState } from '../engine/view'
import { decodeSerial } from '../../formula/serial'
import { sortItemOrder } from '../engine/sort'

const refuse = (feature: string, message: string): EngineResult<never> => ({ ok: false, diagnostic: { kind: 'unsupported', feature, message } })
const invalid = (message: string) => refuse('pivot-delta-invalid', message)
const PERIODS: readonly PivotDateGroupPeriod[] = ['seconds', 'minutes', 'hours', 'days', 'months', 'quarters', 'years']
const namePart: Record<PivotDateGroupPeriod, string> = { seconds: '秒', minutes: '分鐘', hours: '小時', days: '天', months: '月', quarters: '季', years: '年' }
function excelDateCaption(iso: string): string {
  const serial = isoToSerial(iso, false), decoded = decodeSerial(serial, '1900')
  if (!('civil' in decoded)) return iso.slice(0, 10).replace(/-/g, '/')
  return `${decoded.civil.year}/${decoded.civil.month}/${decoded.civil.day}`
}
const groupCaptions = (period: PivotDateGroupPeriod, start: string, end: string, interval: number): string[] => {
  const lo = excelDateCaption(start)
  const hi = excelDateCaption(end)
  if (period === 'months') return [`<${lo}`, ...Array.from({ length: 12 }, (_, i) => `${i + 1}月`), `>${hi}`]
  if (period === 'quarters') return [`<${lo}`, '第一季', '第二季', '第三季', '第四季', `>${hi}`]
  if (period === 'years') {
    const first = Number(start.slice(0, 4)), last = Number(end.slice(0, 4))
    if (last - first > 9999) return []
    return [`<${lo}`, ...Array.from({ length: last - first + 1 }, (_, i) => `${first + i}年`), `>${hi}`]
  }
  if (period === 'days') {
    if (interval === 1) {
      const leap = Array.from({ length: 366 }, (_, day) => {
        const d = new Date(Date.UTC(2000, 0, 1 + day))
        return `${d.getUTCMonth() + 1}月${d.getUTCDate()}日`
      })
      return [`<${lo}`, ...leap, `>${hi}`]
    }
    const startDay = Date.parse(`${start.slice(0, 10)}T00:00:00Z`), endDay = Date.parse(`${end.slice(0, 10)}T00:00:00Z`)
    const count = Math.ceil((endDay - startDay) / 86_400_000 / interval) + 1
    if (count > 10000) return []
    const captions = [`<${lo}`]
    for (let i = 0; i < count; i++) {
      const a = new Date(startDay + i * interval * 86_400_000), b = new Date(Math.min(endDay, startDay + ((i + 1) * interval - 1) * 86_400_000))
      const fmt = (d: Date) => `${d.getUTCFullYear()}/${d.getUTCMonth() + 1}/${d.getUTCDate()}`
      captions.push(`${fmt(a)} - ${fmt(b)}`)
    }
    captions.push(`>${hi}`)
    return captions
  }
  const max = period === 'hours' ? 24 : 60, unit = period === 'hours' ? '時' : period === 'minutes' ? '分' : '秒'
  return [`<${lo}`, ...Array.from({ length: max }, (_, i) => `${i}${unit}`), `>${hi}`]
}
function dateRange(cache: PivotCacheDefinition, base: number, date1904: boolean, start?: number, end?: number): { start: string; end: string; autoStart: boolean; autoEnd: boolean } | undefined {
  const shared = cache.fields[base]?.sharedItems
  if (!shared?.containsDate || shared.containsNonDate) return undefined
  const startText = start === undefined ? shared.minDate : serialText(start, date1904)
  const endText = end === undefined ? shared.maxDate : serialText(end, date1904)
  if (!startText || !endText || isoToSerial(startText, date1904) > isoToSerial(endText, date1904)) return undefined
  return { start: startText, end: endText, autoStart: start === undefined, autoEnd: end === undefined }
}
function serialText(serial: number, date1904: boolean): string | undefined {
  const decoded = decodeSerial(serial, date1904 ? '1904' : '1900')
  if (!('civil' in decoded)) return undefined
  const { year, month, day } = decoded.civil
  if (year < 1900 || year > 9999) return undefined
  return `${String(year).padStart(4, '0')}-${String(month).padStart(2, '0')}-${String(day).padStart(2, '0')}T00:00:00`
}
const groupItems = (labels: readonly string[]): PivotCacheItem[] => labels.map(value => ({ type: 'string', value }))
const pivotItems = (count: number): PivotFieldItem[] => [
  ...Array.from({ length: count }, (_, x): PivotFieldItem => ({ type: 'data', x, hidden: false, char: false, showDetails: true, calculated: false, missing: false, childItems: false, expanded: false, drillAcross: false })),
  { type: 'default', hidden: false, char: false, showDetails: true, calculated: false, missing: false, childItems: false, expanded: false, drillAcross: false },
]
const pivotItemsForOrder = (order: readonly number[]): PivotFieldItem[] => [
  ...order.map((x): PivotFieldItem => ({ type: 'data', x, hidden: false, char: false, showDetails: true, calculated: false, missing: false, childItems: false, expanded: false, drillAcross: false })),
  { type: 'default', hidden: false, char: false, showDetails: true, calculated: false, missing: false, childItems: false, expanded: false, drillAcross: false },
]
const makeFieldState = (old: FieldState, count: number): FieldState => ({ ...old, hidden: Array.from({ length: count }, () => false), collapsed: Array.from({ length: count }, () => false), sortType: 'manual', autoSortScope: undefined })
function groupBase(cache: PivotCacheDefinition, field: number): number {
  const group = cache.fields[field]?.fieldGroup
  if (!group) return field
  if (group.base !== undefined) return group.base
  if (group.par !== undefined) return field
  return field
}
function axisOf(view: PivotViewState, field: number): 'row' | 'col' | 'page' | undefined {
  if (view.rows.includes(field)) return 'row'
  if (view.cols.includes(field)) return 'col'
  if (view.pages.some(page => page.field === field)) return 'page'
  return undefined
}
function axisOfGroup(view: PivotViewState, cache: PivotCacheDefinition, base: number): 'row' | 'col' | 'page' | undefined {
  const child = cache.fields.findIndex((field, index) => index !== base && field.fieldGroup?.base === base && axisOf(view, index) !== undefined)
  return child >= 0 ? axisOf(view, child) : axisOf(view, base)
}
function remapView(view: PivotViewState, removed: ReadonlySet<number>): PivotViewState {
  const rows = view.rows.filter(f => !removed.has(f)), cols = view.cols.filter(f => !removed.has(f))
  const pages = view.pages.filter(p => !removed.has(p.field))
  const filters = view.filters.filter(f => !removed.has(f.fld))
  const fields = view.fields.filter((_, i) => !removed.has(i)).map(state => {
    if (!state.autoSortScope) return state
    const references = state.autoSortScope.references.filter(ref => ref.field === undefined || !removed.has(ref.field))
    return { ...state, ...(references.length ? { autoSortScope: { ...state.autoSortScope, references } } : { autoSortScope: undefined }) }
  })
  return { ...view, rows, cols, pages, filters, fields }
}
function removeGroupFields(ctx: PivotDeltaContext, base: number): { cache: PivotCacheDefinition; def: PivotDeltaContext['def']; view: PivotViewState; removed: Set<number> } {
  const removed = new Set<number>()
  for (let i = ctx.cache.fields.length - 1; i >= 0; i--) {
    const fg = ctx.cache.fields[i]?.fieldGroup
    if (i !== base && fg?.base === base) removed.add(i)
  }
  const cache = { ...ctx.cache, fields: ctx.cache.fields.filter((_, i) => !removed.has(i)).map((f, i) => i === base ? { ...f, fieldGroup: undefined } : f) }
  const def = { ...ctx.def, fields: ctx.def.fields.filter((_, i) => !removed.has(i)), rowFields: ctx.def.rowFields.filter(f => !removed.has(f)), colFields: ctx.def.colFields.filter(f => !removed.has(f)), pageFields: ctx.def.pageFields.filter(f => !removed.has(f.fld)) }
  return { cache, def, view: remapView(ctx.view, removed), removed }
}
function allocation(ctx: PivotDeltaContext, bytes: number): { hold?: PivotMemoryHold } | undefined {
  const hold = ctx.admitCandidateBytes?.(bytes)
  return ctx.admitCandidateBytes && !hold ? undefined : { ...(hold ? { hold } : {}) }
}
function groupedAxisSet(cache: PivotCacheDefinition, base: number): Set<number> {
  const out = new Set<number>([base])
  cache.fields.forEach((field, index) => { if (index !== base && field.fieldGroup?.base === base) out.add(index) })
  return out
}
function insertionPoint(axis: readonly number[], grouped: ReadonlySet<number>): number {
  const first = axis.findIndex(field => grouped.has(field))
  return first < 0 ? axis.length : axis.slice(0, first).filter(field => !grouped.has(field)).length
}
function groupDate(ctx: PivotDeltaContext, delta: Extract<PivotDelta, { kind: 'groupDate' }>): EngineResult<PivotDeltaStateResult> {
  if (!Number.isInteger(delta.field) || delta.field < 0 || delta.field >= ctx.cache.fields.length) return invalid('The date field does not exist.')
  const base = groupBase(ctx.cache, delta.field), axis = axisOf(ctx.view, delta.field) ?? axisOfGroup(ctx.view, ctx.cache, base)
  if (!axis || axis === 'page') return invalid('Group a date field on Rows or Columns.')
  const periods = [...new Set(delta.periods)].sort((a, b) => PERIODS.indexOf(a) - PERIODS.indexOf(b))
  if (!periods.length || periods.some(p => !PERIODS.includes(p))) return invalid('Choose at least one date grouping period.')
  if (delta.startDate !== undefined && !Number.isFinite(delta.startDate) || delta.endDate !== undefined && !Number.isFinite(delta.endDate) ||
      delta.startDate !== undefined && delta.endDate !== undefined && delta.startDate > delta.endDate) return invalid('The date grouping bounds are invalid.')
  const daysInterval = delta.daysInterval ?? 1
  if (!Number.isInteger(daysInterval) || daysInterval < 1 || daysInterval > 32767 || (periods.length > 1 && daysInterval > 1)) return invalid('Number of days must be 1 unless Days is the only selected period.')
  if (ctx.captionRule?.locale !== 'zh-TW') return refuse('pivot-caption-locale', 'Date group captions are measured only for Excel zh-TW.')
  const bounds = dateRange(ctx.cache, base, ctx.date1904, delta.startDate, delta.endDate)
  if (!bounds) return refuse('pivot-engine-unmeasured', `Date bounds for field ${base} are not available as measured dates.`)
  const rootColumn = ctx.store.columns[base]
  if (!rootColumn || ctx.cache.fields[base]?.sharedItems?.inline) return refuse('pivot-records-unmeasured', 'Grouping an inline date field is not measured.')
  const nameBase = ctx.def.fields[base]?.name ?? ctx.cache.fields[base]?.name ?? `Field ${base + 1}`
  const sourceAxis = axis === 'row' ? ctx.view.rows : ctx.view.cols
  const insertAt = insertionPoint(sourceAxis, groupedAxisSet(ctx.cache, base))
  const generated = periods.map(period => ({ period, labels: groupCaptions(period, bounds.start, bounds.end, daysInterval) }))
  if (generated.some(entry => !entry.labels.length || entry.labels.length > 10000)) return refuse('pivot-records-limit', 'The date group has too many items.')
  // Reserve an upper bound before any group labels, cache fields, or pivot-item arrays are created.
  const bytes = generated.reduce((sum, g) => sum + g.labels.length * 240 + 512, 0)
  const admitted = allocation(ctx, bytes)
  if (!admitted) return refuse('pivot-records-limit', 'Date grouping exceeds the memory budget.')
  try {
    const prior = removeGroupFields(ctx, base)
    const fields = [...prior.cache.fields], pivotFields = [...prior.def.fields], viewFields = [...prior.view.fields]
    const newIndices = new Map<PivotDateGroupPeriod, number>()
    for (const { period, labels } of generated) {
      const index = fields.length, items = groupItems(labels)
      newIndices.set(period, index)
      fields.push({ ...ctx.cache.fields[base]!, name: `${namePart[period]} (${nameBase})`, databaseField: false, sharedItems: { containsSemiMixedTypes: false, containsNonDate: false, containsDate: false, containsString: true, containsBlank: false, containsMixedTypes: false, containsNumber: false, containsInteger: false, longText: false, inline: false, items: [] }, fieldGroup: { base, rangePr: { autoStart: bounds.autoStart, autoEnd: bounds.autoEnd, groupBy: period, startDate: bounds.start, endDate: bounds.end, groupInterval: period === 'days' ? daysInterval : 1 }, groupItems: items } })
      const source = ctx.def.fields[base]!
      const axisName = axis === 'row' ? 'axisRow' : 'axisCol'
      const pf: PivotField = { ...source, axis: axisName, items: pivotItems(items.length) }
      pivotFields.push(pf)
      viewFields.push(makeFieldState(prior.view.fields[base]!, pf.items!.length))
    }
    const oldIndex = newIndices.get(periods[0]!)!
    fields[base] = { ...fields[base]!, fieldGroup: { par: oldIndex } }
    const newGroupIndices = periods.map(period => newIndices.get(period)!)
    const ordered = [...newGroupIndices].reverse()
    const rows = [...prior.view.rows], cols = [...prior.view.cols]
    const targetAxis = axis === 'row' ? rows : cols
    const baseAt = targetAxis.indexOf(base)
    if (baseAt >= 0) targetAxis.splice(baseAt, 1)
    targetAxis.splice(Math.min(insertAt, targetAxis.length), 0, ...ordered)
    for (const index of ordered) pivotFields[index] = { ...pivotFields[index]!, axis: axis === 'row' ? 'axisRow' : 'axisCol' }
    const def = { ...prior.def, fields: pivotFields, rowFields: rows, colFields: cols }
    const view: PivotViewState = { ...prior.view, rows, cols, fields: viewFields }
    return { ok: true, value: { def, cache: { ...prior.cache, fields }, store: ctx.store, view, extraBytes: bytes, ...(admitted.hold ? { hold: admitted.hold } : {}) } }
  } catch (error) { admitted.hold?.release(); throw error }
}
function numberGroup(ctx: PivotDeltaContext, delta: Extract<PivotDelta, { kind: 'groupNumber' }>): EngineResult<PivotDeltaStateResult> {
  const field = delta.field, cf = ctx.cache.fields[field], shared = cf?.sharedItems, axis = axisOf(ctx.view, field)
  if (!Number.isInteger(field) || field < 0 || !cf || !axis || axis === 'page') return invalid('Group a number field on Rows or Columns.')
  if (ctx.captionRule?.locale !== 'zh-TW') return refuse('pivot-caption-locale', 'Number group captions are measured only for Excel zh-TW.')
  if (!shared?.containsNumber || shared.containsString || shared.containsDate || shared.containsMixedTypes) return refuse('pivot-engine-unmeasured', 'Only numeric fields can be grouped by number.')
  if (!Number.isFinite(delta.by) || delta.by <= 0 || (delta.start !== undefined && !Number.isFinite(delta.start)) || (delta.end !== undefined && !Number.isFinite(delta.end))) return invalid('By must be greater than zero and bounds must be finite.')
  const min = delta.start ?? shared.minValue, max = delta.end ?? shared.maxValue
  if (min === undefined || max === undefined || min > max) return invalid('The numeric grouping bounds are invalid.')
  const count = Math.ceil((max - min) / delta.by)
  if (!Number.isSafeInteger(count) || count < 1 || count > 10000) return refuse('pivot-records-limit', 'The number group has too many items.')
  const bytes = (count + 2) * 240 + 512
  const admitted = allocation(ctx, bytes)
  if (!admitted) return refuse('pivot-records-limit', 'Number grouping exceeds the memory budget.')
  try {
    const labels = [`<${min}${shared.containsBlank ? ' 或 (空白)' : ''}`, ...Array.from({ length: count }, (_, i) => {
      const start = min + i * delta.by, end = min < 0 ? Math.min(start + delta.by - 1, min + count * delta.by) : start + delta.by
      return `${start}-${end}`
    }), `>${min + count * delta.by}`]
    const items = groupItems(labels), pivots = pivotItems(items.length)
    const cache = { ...ctx.cache, fields: ctx.cache.fields.map((field, i) => i === delta.field ? { ...field, fieldGroup: { base: delta.field, rangePr: { autoStart: delta.start === undefined, autoEnd: delta.end === undefined, groupBy: 'range' as const, startNum: min, endNum: max, groupInterval: delta.by }, groupItems: items, savedPivotItems: field.fieldGroup?.savedPivotItems ?? ctx.def.fields[i]?.items, savedFieldState: field.fieldGroup?.savedFieldState ?? ctx.view.fields[i] } } : field) }
    const def = { ...ctx.def, fields: ctx.def.fields.map((field, i) => i === delta.field ? { ...field, items: pivots } : field) }
    const view = { ...ctx.view, fields: ctx.view.fields.map((field, i) => i === delta.field ? makeFieldState(field, pivots.length) : field) }
    return { ok: true, value: { def, cache, store: ctx.store, view, extraBytes: bytes, ...(admitted.hold ? { hold: admitted.hold } : {}) } }
  } catch (error) { admitted.hold?.release(); throw error }
}
function ungroup(ctx: PivotDeltaContext, delta: Extract<PivotDelta, { kind: 'ungroup' }>): EngineResult<PivotDeltaStateResult> {
  const field = delta.field, cf = ctx.cache.fields[field]
  if (!Number.isInteger(field) || field < 0 || !cf) return invalid('The grouped field does not exist.')
  const base = groupBase(ctx.cache, field), group = cf.fieldGroup
  if (!group || (!group.base && !group.discretePr && !cf.fieldGroup?.par)) return invalid('This field is not grouped.')
  const groupSet = groupedAxisSet(ctx.cache, base)
  const groupedField = ctx.cache.fields.findIndex((candidate, index) => index !== base && candidate.fieldGroup?.base === base && axisOf(ctx.view, index) !== undefined)
  const axis = axisOf(ctx.view, field) ?? axisOf(ctx.view, base) ?? (groupedField >= 0 ? axisOf(ctx.view, groupedField) : undefined)
  const sourceAxis = axis === 'row' ? ctx.view.rows : ctx.view.cols
  const insertAt = insertionPoint(sourceAxis, groupSet)
  const prior = removeGroupFields(ctx, base)
  let cache = prior.cache, store = ctx.store, def = prior.def, view = prior.view
  if (axis === 'row' || axis === 'col') {
    const rows = [...view.rows], cols = [...view.cols]
    const target = axis === 'row' ? rows : cols
    if (!target.includes(base)) target.splice(Math.min(insertAt, target.length), 0, base)
    view = { ...view, rows, cols }
    def = { ...def, rowFields: rows, colFields: cols, fields: def.fields.map((f, i) => i === base ? { ...f, axis: axis === 'row' ? 'axisRow' : 'axisCol' } : f) }
  }
  const source = store.columns[base]
  if (group.rangePr && group.base === base && source?.kind === 'shared') {
    let items = group.savedPivotItems
    if (!items) {
      const raw = cache.fields[base]?.sharedItems?.items ?? []
      const order = sortItemOrder(raw.map((_, index) => index), raw, 'ascending', base, ctx.def.customListSort !== false)
      if (!order.ok) return order
      items = pivotItemsForOrder(order.value)
    }
    const state = group.savedFieldState as FieldState | undefined
    def = { ...def, fields: def.fields.map((f, i) => i === base ? { ...f, items } : f) }
    view = { ...view, fields: view.fields.map((f, i) => i === base ? state ?? makeFieldState(f, items.length) : f) }
  }
  if (cache.fields[base]?.sharedItems?.inline && source?.kind === 'inline') {
    const bytes = store.count * 108 + 256, admitted = allocation(ctx, bytes)
    if (!admitted) return refuse('pivot-records-limit', 'Ungrouping this inline field exceeds the memory budget.')
    const items: PivotCacheItem[] = [], indices = new Uint32Array(store.count), byKey = new Map<string, number>()
    for (let row = 0; row < store.count; row++) {
      const tag = source.tag[row]!, text = source.text[source.textIndex[row]!] ?? ''
      const value: PivotCacheItem = tag === VALUE_TAG.missing ? { type: 'missing' } : tag === VALUE_TAG.number || tag === VALUE_TAG.date ? { type: tag === VALUE_TAG.date ? 'date' : 'number', value: source.num[row]! } as PivotCacheItem : tag === VALUE_TAG.boolean ? { type: 'boolean', value: !!source.num[row]! } : tag === VALUE_TAG.error ? { type: 'error', value: text } : { type: 'string', value: text }
      const key = `${value.type}:${'value' in value ? String(value.value) : ''}`
      let index = byKey.get(key); if (index === undefined) { index = items.length; byKey.set(key, index); items.push(value) }
      indices[row] = index
    }
    const old = cache.fields[base]!.sharedItems!
    cache = { ...cache, fields: cache.fields.map((f, i) => i === base ? { ...f, sharedItems: { ...old, inline: false, items } } : f) }
    store = { ...store, columns: store.columns.map((c, i) => i === base ? { kind: 'shared' as const, index: indices } : c), bytes: store.bytes + indices.byteLength }
    const pItems = pivotItems(items.length), pf = { ...def.fields[base]!, items: pItems }
    def = { ...def, fields: def.fields.map((f, i) => i === base ? pf : f) }
    view = { ...view, fields: view.fields.map((f, i) => i === base ? makeFieldState(f, pItems.length) : f) }
    return { ok: true, value: { def, cache, store, view, extraBytes: bytes, ...(admitted.hold ? { hold: admitted.hold } : {}) } }
  }
  return { ok: true, value: { def, cache, store, view, extraBytes: 0 } }
}
export function applyGrouping(ctx: PivotDeltaContext, delta: Extract<PivotDelta, { kind: 'groupDate' | 'groupNumber' | 'ungroup' }>): EngineResult<PivotDeltaStateResult> {
  if (ctx.def.disableFieldList) return refuse('pivot-field-list-disabled', `The field list is disabled for ${ctx.def.name}.`)
  if (delta.kind === 'groupDate') return groupDate(ctx, delta)
  if (delta.kind === 'groupNumber') return numberGroup(ctx, delta)
  return ungroup(ctx, delta)
}

/** The menu capability for a field, computed without walking the record store. */
export function pivotMenuGroupInfo(cache: PivotCacheDefinition, field: number, date1904: boolean): PivotMenuGroupInfo {
  const cf = cache.fields[field], group = cf?.fieldGroup
  const baseField = group?.base ?? field
  const grouped = Boolean(group?.base !== undefined || group?.rangePr || group?.discretePr || group?.par)
  const shared = cache.fields[baseField]?.sharedItems
  const periods = cache.fields.flatMap(candidate => candidate.fieldGroup?.base === baseField && candidate.fieldGroup.rangePr && candidate.fieldGroup.rangePr.groupBy !== 'range' ? [candidate.fieldGroup.rangePr.groupBy as PivotDateGroupPeriod] : [])
  if (shared?.containsDate && !shared.containsNonDate) {
    const bounds = shared.minDate && shared.maxDate ? { min: isoToSerial(shared.minDate, date1904), max: isoToSerial(shared.maxDate, date1904) } : undefined
    const range = group?.rangePr
    return { kind: 'date', grouped, baseField, ...(bounds ? { bounds } : {}), date: { periods: periods.length ? periods : [], ...(range?.startDate ? { startDate: isoToSerial(range.startDate, date1904) } : {}), ...(range?.endDate ? { endDate: isoToSerial(range.endDate, date1904) } : {}), autoStart: range?.autoStart ?? true, autoEnd: range?.autoEnd ?? true, daysInterval: range?.groupInterval ?? 1 } }
  }
  if (shared?.containsNumber && !shared.containsString && !shared.containsDate && !shared.containsMixedTypes) {
    const range = group?.rangePr
    return { kind: 'number', grouped, baseField, ...(shared.minValue !== undefined && shared.maxValue !== undefined ? { bounds: { min: shared.minValue, max: shared.maxValue } } : {}), number: { ...(range?.startNum !== undefined ? { start: range.startNum } : {}), ...(range?.endNum !== undefined ? { end: range.endNum } : {}), by: range?.groupInterval ?? 1, autoStart: range?.autoStart ?? true, autoEnd: range?.autoEnd ?? true } }
  }
  return { grouped, baseField }
}
