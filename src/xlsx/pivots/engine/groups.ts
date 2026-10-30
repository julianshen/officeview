/** Grouped cache fields (spec §5.1). Values are derived per record from the fieldGroup's base column. */
import type { PivotCacheDefinition, PivotDiagnostic, PivotValue } from '../model'
import type { EngineResult } from './axes'
import { VALUE_TAG, type RecordStore } from './records'

export const NO_GROUP = 0xffffffff
const fail = (feature: string, message: string): EngineResult<never> => ({
  ok: false, diagnostic: { kind: 'unsupported', feature, message } satisfies PivotDiagnostic,
})
const invalid = (message: string): EngineResult<never> => ({
  ok: false, diagnostic: { kind: 'parse', feature: 'pivot-definition-invalid', message },
})
const unmeasured = (message: string) => fail('pivot-engine-unmeasured', message)

function sourceValue(store: RecordStore, cache: PivotCacheDefinition, base: number, row: number): PivotValue | undefined {
  const column = store.columns[base]
  if (!column) return undefined
  if (column.kind === 'shared') return cache.fields[base]?.sharedItems?.items[column.index[row]!]
  const tag = column.tag[row]
  switch (tag) {
    case VALUE_TAG.missing: return { type: 'missing' }
    case VALUE_TAG.number: return { type: 'number', value: column.num[row]! }
    case VALUE_TAG.date: return { type: 'number', value: column.num[row]! }
    case VALUE_TAG.boolean: return { type: 'boolean', value: !!column.num[row] }
    case VALUE_TAG.error: return { type: 'error', value: column.text[column.textIndex[row]!] ?? '' }
    case VALUE_TAG.string: return { type: 'string', value: column.text[column.textIndex[row]!] ?? '' }
    default: return undefined
  }
}

function missingGroup(groupItems: readonly PivotValue[]): number {
  return groupItems.findIndex(item => item.type === 'missing' || (item.type === 'string' && item.value.toLocaleLowerCase() === '(blank)'))
}
function caption(item: PivotValue): string | undefined {
  return item.type === 'string' || item.type === 'number' ? String(item.value) : undefined
}
function rangeBucket(groupItems: readonly PivotValue[], start: number, end: number, interval: number, value: number): number {
  if (value < start) return groupItems.findIndex(item => item.type === 'string' && /^\s*</u.test(item.value))
  if (value > end) return groupItems.findIndex(item => item.type === 'string' && /^\s*>/u.test(item.value))
  const buckets: Array<{ index: number; start: number; end: number }> = []
  for (let i = 0; i < groupItems.length; i++) {
    const text = caption(groupItems[i]!)
    if (!text) continue
    const match = /^\s*([+-]?(?:\d+(?:\.\d*)?|\.\d+))\s*[-–]\s*([+-]?(?:\d+(?:\.\d*)?|\.\d+))\s*$/.exec(text)
    if (match) buckets.push({ index: i, start: Number(match[1]), end: Number(match[2]) })
  }
  buckets.sort((a, b) => a.start - b.start)
  const count = Math.ceil((end - start) / interval)
  if (buckets.length !== count) return -1
  const slot = Math.min(Math.floor((value - start) / interval), count - 1)
  const bucket = buckets[slot]
  if (!bucket) return -1
  const expectedStart = start + slot * interval
  const tolerance = Math.max(1, Math.abs(expectedStart)) * Number.EPSILON * 8
  if (Math.abs(bucket.start - expectedStart) > tolerance) return -1
  return bucket.index
}

const MONTHS = ['January', 'February', 'March', 'April', 'May', 'June', 'July', 'August', 'September', 'October', 'November', 'December']
function datePartLabel(groupItems: readonly PivotValue[], value: string, groupBy: string): number {
  const date = new Date(value.endsWith('Z') ? value : `${value}Z`)
  if (!Number.isFinite(date.getTime())) return -1
  const year = date.getUTCFullYear(), month = date.getUTCMonth(), day = date.getUTCDate()
  const hour = date.getUTCHours(), minute = date.getUTCMinutes(), second = date.getUTCSeconds()
  const candidates = new Set<string>()
  if (groupBy === 'years') candidates.add(String(year))
  else if (groupBy === 'quarters') {
    const q = Math.floor(month / 3) + 1
    candidates.add(`Q${q}`); candidates.add(`Qtr${q}`); candidates.add(`Quarter ${q}`); candidates.add(`Quarter${q}`)
  } else if (groupBy === 'months') {
    candidates.add(MONTHS[month]!); candidates.add(MONTHS[month]!.slice(0, 3)); candidates.add(String(month + 1))
  } else if (groupBy === 'days') {
    candidates.add(`${year}-${String(month + 1).padStart(2, '0')}-${String(day).padStart(2, '0')}`)
    candidates.add(String(day))
  } else if (groupBy === 'hours') candidates.add(String(hour))
  else if (groupBy === 'minutes') candidates.add(String(minute))
  else if (groupBy === 'seconds') candidates.add(String(second))
  for (let i = 0; i < groupItems.length; i++) {
    const item = groupItems[i]!
    if (item.type === 'date' && groupBy === 'days' && item.value.slice(0, 10) === `${year}-${String(month + 1).padStart(2, '0')}-${String(day).padStart(2, '0')}`) return i
    const label = caption(item)
    if (label !== undefined && candidates.has(label)) return i
  }
  return -1
}

/** Returns one groupItems index per source record. The sentinel represents an unmapped discrete item. */
export function groupColumn(store: RecordStore, cache: PivotCacheDefinition, field: number): EngineResult<Uint32Array> {
  const fieldGroup = cache.fields[field]?.fieldGroup
  const groupItems = fieldGroup?.groupItems
  if (!fieldGroup || !groupItems?.length) return invalid(`field ${field} has no usable groupItems`)
  const base = fieldGroup.base ?? field
  const source = store.columns[base]
  if (!source) return invalid(`group field ${field} has no base record column ${base}`)
  const output = new Uint32Array(store.count).fill(NO_GROUP)
  const blank = missingGroup(groupItems)

  if (fieldGroup.discretePr) {
    const map = fieldGroup.discretePr
    for (let row = 0; row < store.count; row++) {
      const value = sourceValue(store, cache, base, row)
      if (!value) return invalid(`group field ${field} has an invalid base value at record ${row}`)
      const sourceItem = source.kind === 'shared' ? source.index[row]! : -1
      const group = map[sourceItem]
      if (group === undefined) {
        if (value.type === 'missing' && blank >= 0) { output[row] = blank; continue }
        return invalid(`group field ${field} has no discrete mapping for base item ${sourceItem}`)
      }
      if (group >= groupItems.length) return invalid(`group field ${field} maps outside groupItems`)
      output[row] = group
    }
    return { ok: true, value: output }
  }

  const range = fieldGroup.rangePr
  if (!range) return invalid(`group field ${field} has neither rangePr nor discretePr`)
  for (let row = 0; row < store.count; row++) {
    const value = sourceValue(store, cache, base, row)
    if (!value) return invalid(`group field ${field} has an invalid base value at record ${row}`)
    if (value.type === 'missing') {
      if (blank >= 0) output[row] = blank
      continue
    }
    if (range.groupBy === 'range') {
      if (value.type !== 'number') return unmeasured(`number group field ${field} contains a non-numeric base item`)
      const start = range.startNum, end = range.endNum
      if (start === undefined || end === undefined || !Number.isFinite(start) || !Number.isFinite(end) || end < start ||
        !Number.isFinite(range.groupInterval) || range.groupInterval <= 0)
        return invalid(`number group field ${field} has invalid range bounds`)
      const item = rangeBucket(groupItems, start, end, range.groupInterval, value.value)
      if (item < 0) return unmeasured(`number group label for ${value.value} in field ${field} is not represented`)
      output[row] = item
      continue
    }
    if (value.type !== 'date') return unmeasured(`date group field ${field} contains a non-date base item`)
    const iso = value.value.slice(0, 10)
    if (range.startDate && iso < range.startDate.slice(0, 10) || range.endDate && iso > range.endDate.slice(0, 10))
      return unmeasured(`out-of-range date group labels for field ${field} are not measured`)
    if (source.kind === 'inline') return unmeasured(`inline date grouping for field ${field} needs the workbook date-system context`)
    const item = datePartLabel(groupItems, value.value, range.groupBy)
    if (item < 0) return unmeasured(`date group label for ${iso} in field ${field} is not represented`)
    output[row] = item
  }
  return { ok: true, value: output }
}
