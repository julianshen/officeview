/**
 * The refreshOnLoad refresh path (plan Task 32, spec §7.7): Excel refreshes a `refreshOnLoad="1"` pivot from its
 * source when the file opens, so its saved records and cells may both be out of date. Display always shows the saved
 * cells; on the first interaction a session rebuilds the cache and records from the source range, as Excel's
 * refresh does, then applies the change. Lazy chunk `pivots`.
 *
 * Measured against Task 24's refresh oracle (recipes/refresh-seed*.xlsx refreshed and saved by Excel 16.114 as
 * interaction/refresh-on-load*.xlsx; docs/validation/xlsx-pivots.md "refreshOnLoad"), plus the Stage 2 caches:
 * - a field keeps whether it lists shared items (axis fields) or stores its values inline (the others);
 * - a listed field's shared items are rebuilt in first-appearance order over the source rows: after Data!A2 became
 *   North, Region's items are North, East, South, West, Central (they were East, North, …). The definition's pivot
 *   items keep their order and are re-pointed at the new indices (`refreshDefinition`);
 * - the type flags are recomputed from the values: `containsString` (text and errors, engine/quirk-error-value),
 *   `containsBlank`, `containsNumber`, `containsInteger` (every number whole), `containsDate`, `containsNonDate` (any
 *   text or number), `containsSemiMixedTypes` (any text or blank), `containsMixedTypes` (more than one of text,
 *   number and date), `longText` (text over 255 characters), `minValue`/`maxValue` over the numbers and
 *   `minDate`/`maxDate` over the dates. `maxDate` is the latest date's next day when that date has no time of day
 *   (Date: latest 2024-12-20, saved 2024-12-21T00:00:00), else the latest date-time itself (Stamp: 2023-03-15T06:00:00);
 * - the cell values carry no number format here, so a field whose cache held dates only reads its numbers as dates and
 *   any other field reads them as numbers;
 * - text is one value whatever its ASCII case, spelled as first seen: Code's "AB" records are saved as "ab", which
 *   comes first (an inline field). Case folding beyond ASCII is not measured: texts that only it would merge refuse.
 * Refused as `pivot-refresh-on-load` (not measured, or not readable): a source that is a defined name or table, a
 * header row or column count that no longer matches the cache's database fields, grouped fields, booleans, texts
 * that only non-ASCII case folding would merge, a non-date value in a dates-only field, a field mixing dates and
 * numbers, an incomplete source read, and a definition whose listed fields gained or lost items. The record limits
 * refuse as `pivot-records-limit`.
 */
import type { SheetRect } from '../../tableStyles/model'
import type { PivotCacheDefinition, PivotCacheField, PivotCacheItem, PivotDiagnostic, PivotSharedItems, PivotTableDefinition, PivotValue } from '../model'
import type { EngineResult } from '../engine/axes'
import type { CellReader } from '../engine/faithful'
import { recordEvents, type RecordLimits, type RecordStore } from '../engine/records'
import { placeCells, type PivotLayout } from '../engine/place'
import { initialViewState } from '../engine/view'
import { savedPivotLayout } from '../savedLayout'
import { parseA1Range } from '../../view/coordinates'

const refuse = (message: string): EngineResult<never> => ({ ok: false, diagnostic: { kind: 'unsupported', feature: 'pivot-refresh-on-load', message } satisfies PivotDiagnostic })
const limit = (message: string): EngineResult<never> => ({ ok: false, diagnostic: { kind: 'unsupported', feature: 'pivot-records-limit', message } satisfies PivotDiagnostic })

/** A worksheet range source's sheet and rectangle; undefined for a defined name or table, another workbook or no ref. */
export function worksheetSourceRange(cache: PivotCacheDefinition): { sheet: string; ref: SheetRect } | undefined {
  const source = cache.source.worksheet
  if (cache.source.type !== 'worksheet' || !source || source.relId !== undefined || source.name !== undefined || !source.ref || !source.sheet) return undefined
  const rect = parseA1Range(source.ref)
  return rect ? { sheet: source.sheet, ref: { minRow: rect.r0, minCol: rect.c0, maxRow: rect.r1, maxCol: rect.c1 } } : undefined
}

const DAY_MS = 86_400_000
/** Excel's last date serial, 9999-12-31 (1900 system; the 1904 system ends 1,462 days earlier). */
const MAX_SERIAL = 2_958_465
/** Whether a serial is a date the cache can write: finite, not negative, at most 9999-12-31 23:59:59.999. */
const isDateSerial = (serial: number, date1904: boolean) => Number.isFinite(serial) && serial >= 0 && serial < (date1904 ? MAX_SERIAL - 1462 : MAX_SERIAL) + 1
/** Excel serial → the ISO dateTime the cache writes (uniform epoch, Ruling S2-H; records.ts `isoToSerial` inverts it).
 * Callers check `isDateSerial` first. */
function serialToIso(serial: number, date1904: boolean): string {
  const epoch = date1904 ? Date.UTC(1904, 0, 1) : Date.UTC(1899, 11, 30)
  const iso = new Date(epoch + Math.round(serial * DAY_MS)).toISOString()
  return iso.endsWith('.000Z') ? iso.slice(0, 19) : iso.slice(0, 23)
}
const isDatesOnly = (field: PivotCacheField) => Boolean(field.sharedItems?.containsDate && !field.sharedItems.containsNonDate)
const isMixedDates = (field: PivotCacheField) => Boolean(field.sharedItems?.containsDate && field.sharedItems.containsNonDate)
const asciiLower = (text: string) => text.replace(/[A-Z]+/g, run => run.toLowerCase())
/** Each text's first-seen spelling under ASCII case folding; undefined when full case folding would merge it with
 * another text that ASCII folding keeps apart (not measured). */
function spellings() {
  const first = new Map<string, string>(), full = new Map<string, string>()
  return {
    get size() { return first.size },
    of(text: string): string | undefined {
      const key = asciiLower(text), wide = text.toLowerCase()
      const known = full.get(wide)
      if (known !== undefined && known !== key) return undefined
      full.set(wide, key)
      let spelled = first.get(key)
      if (spelled === undefined) first.set(key, spelled = text)
      return spelled
    },
  }
}
const itemKey = (value: PivotValue) => value.type === 'missing' ? 'm' : `${value.type}:${String(value.value)}`

/** Bytes the build keeps besides the store, per distinct text (its spellings and keys) and per listed item; a listed
 * item stays in the refreshed cache for the session's life (`refreshedCacheBytes`). */
const TEXT_BYTES = (text: string) => 3 * (48 + 2 * text.length)
const ITEM_BYTES = 96

/** One field's rebuilt shared items, accumulated value by value (no column of values is kept). */
function sharedItemsAccumulator(old: PivotSharedItems, listed: boolean, date1904: boolean, name: string) {
  let hasString = false, hasBlank = false, hasNumber = false, hasDate = false, longText = false, integers = true
  let minValue = Infinity, maxValue = -Infinity, minDate = Infinity, maxDate = -Infinity
  const items: PivotCacheItem[] = []
  const index = new Map<string, number>()
  return {
    index,
    /** Adds one value; returns the bytes it added to the build (a new listed item), or a refusal. */
    add(value: PivotValue): number | PivotDiagnostic {
      switch (value.type) {
        case 'missing': hasBlank = true; break
        case 'string': case 'error':
          hasString = true
          if (value.value.length > 255) longText = true
          break
        case 'number':
          hasNumber = true
          if (!Number.isInteger(value.value)) integers = false
          minValue = Math.min(minValue, value.value); maxValue = Math.max(maxValue, value.value)
          break
        case 'date': {
          hasDate = true
          const serial = dateSerial(value.value, date1904)
          minDate = Math.min(minDate, serial); maxDate = Math.max(maxDate, serial)
          break
        }
        case 'boolean': return { kind: 'unsupported', feature: 'pivot-refresh-on-load', message: `field ${name} has a boolean value, whose refresh is not measured` }
      }
      if (!listed) return 0
      const key = itemKey(value)
      if (index.has(key)) return 0
      index.set(key, items.length)
      items.push({ ...value })
      return ITEM_BYTES
    },
    finish(): PivotSharedItems {
      const types = Number(hasString) + Number(hasNumber) + Number(hasDate)
      const { minValue: _a, maxValue: _b, minDate: _c, maxDate: _d, ...rest } = old
      return {
        ...rest,
        containsSemiMixedTypes: hasString || hasBlank,
        containsNonDate: hasString || hasNumber,
        containsDate: hasDate,
        containsString: hasString,
        containsBlank: hasBlank,
        containsMixedTypes: types > 1,
        containsNumber: hasNumber,
        containsInteger: hasNumber && integers,
        ...(hasNumber ? { minValue, maxValue } : {}),
        ...(hasDate ? { minDate: serialToIso(minDate, date1904), maxDate: serialToIso(Number.isInteger(maxDate) ? maxDate + 1 : maxDate, date1904) } : {}),
        longText,
        inline: !listed,
        items,
      }
    },
  }
}
function dateSerial(iso: string, date1904: boolean): number {
  const epoch = date1904 ? Date.UTC(1904, 0, 1) : Date.UTC(1899, 11, 30)
  return (Date.parse(`${iso}Z`) - epoch) / DAY_MS
}

/**
 * The cache and record store Excel's refresh builds from `ref` (header row, then one record per row) of the source
 * sheet that `cells` reads. Pure: `cache` is not modified.
 */
export function recordsFromSource(cache: PivotCacheDefinition, cells: CellReader, ref: SheetRect, date1904: boolean, limits: RecordLimits): EngineResult<{ cache: PivotCacheDefinition; store: RecordStore }> {
  if (cells.incomplete) return refuse(`the source range could not be read (${cells.incomplete})`)
  const database = cache.fields.flatMap((field, index) => field.databaseField ? [index] : [])
  if (ref.maxCol - ref.minCol + 1 !== database.length) return refuse(`the source range has ${ref.maxCol - ref.minCol + 1} columns for ${database.length} cache fields`)
  if (cache.fields.some(field => field.fieldGroup)) return refuse('refreshing grouped fields is not measured')
  for (let k = 0; k < database.length; k++) {
    const header = cells(ref.minRow, ref.minCol + k), field = cache.fields[database[k]!]!
    if (header?.kind !== 'text' || header.value !== field.name) return refuse(`source column ${k + 1} is headed ${header ? JSON.stringify(header.value) : 'blank'}, not ${JSON.stringify(field.name)}`)
    if (!field.sharedItems) return refuse(`field ${field.name} has no shared items description`)
    if (isMixedDates(field)) return refuse(`field ${field.name} mixes dates and numbers; their cells cannot be told apart without formats`)
  }
  const count = ref.maxRow - ref.minRow
  if (count > limits.maxRecords) return limit(`the source range has ${count} records, more than the ${limits.maxRecords} allowed`)
  // The store's least size follows from the shape alone (4 bytes per listed value, 13 per inline one): refuse before
  // reading a row when even that exceeds the limit.
  const rowBytes = database.reduce((sum, index) => sum + (cache.fields[index]!.sharedItems!.inline === false ? 4 : 13), 0)
  if (count * rowBytes > limits.maxBytes) return limit(`the source range's ${count} records need at least ${count * rowBytes} bytes, more than the ${limits.maxBytes} allowed`)

  // Two passes over the cells, keeping no column of values (codex review T32-1): the first gathers each field's
  // shared items and type flags, the second streams the records into the decoder. What the build keeps besides the
  // store (texts' first spellings, listed items) is counted, admitted as it grows, and left out of the decoder's share.
  const spelling = database.map(() => spellings())
  let kept = 0, inlineTexts = 0
  /** Admits `kept + bytes` (a reservation, or a keep when `commit`): refused past `maxBytes` or by the pool. */
  const admit = (bytes: number, commit: boolean): EngineResult<never> | undefined => {
    if (!bytes) return undefined
    const total = kept + bytes
    if (total > limits.maxBytes) return limit(`refreshing needs more than the ${limits.maxBytes} bytes allowed`)
    if (limits.admitBytes && !limits.admitBytes(total)) return { ok: false, diagnostic: { kind: 'unsupported', feature: 'pivot-worker-quota', message: 'Refreshing the pivot exceeds the memory budget.' } }
    if (commit) kept = total
    return undefined
  }
  const keep = (bytes: number) => admit(bytes, true)
  const valueAt = (row: number, k: number): PivotValue | EngineResult<never> => {
    const field = cache.fields[database[k]!]!, cell = cells(row, ref.minCol + k)
    const datesOnly = isDatesOnly(field)
    if (!cell) return { type: 'missing' }
    if (cell.kind === 'number') {
      if (!datesOnly) return { type: 'number', value: cell.value }
      if (!isDateSerial(cell.value, date1904)) return refuse(`field ${field.name} has the number ${cell.value}, which is not a date Excel can hold`)
      return { type: 'date', value: serialToIso(cell.value, date1904) }
    }
    if (datesOnly) return refuse(`field ${field.name} held dates only and now has ${cell.kind} values`)
    if (cell.kind === 'error') return { type: 'error', value: cell.value }
    if (cell.kind === 'bool') return { type: 'boolean', value: cell.value }
    // Case folding copies the text twice and may keep both keys: reserve that before folding (codex re-review T32-R2-1).
    const reserve = TEXT_BYTES(cell.value)
    const refusedReserve = admit(reserve, false)
    if (refusedReserve) return refusedReserve
    const before = spelling[k]!.size
    const spelled = spelling[k]!.of(cell.value)
    if (spelled === undefined) return refuse(`field ${field.name} has texts that differ only in non-ASCII case, whose item identity is not measured`)
    if (spelling[k]!.size !== before) {
      const refused = keep(TEXT_BYTES(cell.value))
      if (refused) return refused
      if (cache.fields[database[k]!]!.sharedItems!.inline !== false && ++inlineTexts > limits.maxInlineStrings) return limit(`the source has more than ${limits.maxInlineStrings} distinct texts`)
    }
    return { type: 'string', value: spelled }
  }
  const isRefusal = (value: PivotValue | EngineResult<never>): value is EngineResult<never> => 'ok' in value
  // Pass 1: fields.
  const accumulators = database.map(index => sharedItemsAccumulator(cache.fields[index]!.sharedItems!, cache.fields[index]!.sharedItems!.inline === false, date1904, cache.fields[index]!.name))
  for (let row = ref.minRow + 1; row <= ref.maxRow; row++) {
    for (let k = 0; k < database.length; k++) {
      const value = valueAt(row, k)
      if (isRefusal(value)) return value
      const added = accumulators[k]!.add(value)
      if (typeof added !== 'number') return { ok: false, diagnostic: added }
      const refused = keep(added)
      if (refused) return refused
    }
  }
  const fields = [...cache.fields]
  database.forEach((index, k) => { fields[index] = { ...cache.fields[index]!, sharedItems: accumulators[k]!.finish() } })
  const rebuilt: PivotCacheDefinition = { ...cache, fields, recordCount: count }

  // Pass 2: records, through the records part's own decoder (same layout, limits and byte accounting), its admission
  // offset by what the build keeps.
  const decoderLimits: RecordLimits = { ...limits, maxBytes: limits.maxBytes - kept, ...(limits.admitBytes ? { admitBytes: (bytes: number) => limits.admitBytes!(kept + bytes) } : {}) }
  const decoder = recordEvents(rebuilt, decoderLimits, date1904)
  const lookup = database.map((index, k) => fields[index]!.sharedItems!.inline ? undefined : accumulators[k]!.index)
  const NONE: readonly string[] = []
  try {
    decoder.events.open('pivotCacheRecords', ['count'], [String(count)], 1, false)
    for (let row = ref.minRow + 1; row <= ref.maxRow; row++) {
      decoder.events.open('r', NONE, NONE, 0, false)
      for (let k = 0; k < database.length; k++) {
        const value = valueAt(row, k)
        if (isRefusal(value)) return value
        const shared = lookup[k]
        if (shared) decoder.events.open('x', ['v'], [String(shared.get(itemKey(value))!)], 1, true)
        else if (value.type === 'missing') decoder.events.open('m', NONE, NONE, 0, true)
        else {
          const tag = value.type === 'number' ? 'n' : value.type === 'string' ? 's' : value.type === 'error' ? 'e' : value.type === 'date' ? 'd' : 'b'
          const raw = value.type === 'boolean' ? (value.value ? '1' : '0') : String(value.value)
          decoder.events.open(tag, ['v'], [raw], 1, true)
        }
      }
      decoder.events.close('r')
    }
    decoder.events.close('pivotCacheRecords')
  } catch { /* the decoder stopped on a refusal: finish() reports it */ }
  const result = decoder.finish()
  if (!result.ok) return result
  return { ok: true, value: { cache: rebuilt, store: result.store } }
}

/**
 * The definition with each listed field's pivot items re-pointed at the rebuilt shared items, keeping their order
 * (refresh-on-load: Region's items x=0,1 became x=1,0). A field whose values gained or lost an item refuses: where
 * Excel puts a new item, and what it does with a vanished one, is not measured.
 */
export function refreshDefinition(def: PivotTableDefinition, before: PivotCacheDefinition, after: PivotCacheDefinition): EngineResult<PivotTableDefinition> {
  const fields = def.fields.map(field => field)
  for (let index = 0; index < def.fields.length; index++) {
    const field = def.fields[index]!, old = before.fields[index]?.sharedItems, next = after.fields[index]?.sharedItems
    if (!old || !next || old.inline || next.inline) continue
    const at = new Map(next.items.map((item, k) => [itemKey(item), k]))
    const oldKeys = new Set(old.items.map(itemKey))
    if (old.items.length !== next.items.length || next.items.some(item => !oldKeys.has(itemKey(item))))
      return refuse(`field ${after.fields[index]!.name} gained or lost items on refresh, whose placement is not measured`)
    if (!field.items) {
      if (old.items.some((item, k) => at.get(itemKey(item)) !== k)) return refuse(`field ${after.fields[index]!.name} has no item list and its shared items moved`)
      continue
    }
    fields[index] = { ...field, items: field.items.map(item => item.x === undefined ? item : { ...item, x: at.get(itemKey(old.items[item.x]!))! }) }
  }
  return { ok: true, value: { ...def, fields } }
}

/**
 * The saved state's layout laid out from the saved lines alone (`rowItems`/`colItems`), with no records: its labels and
 * implicit captions sit where Excel saved them, so the caption rule can read the saved texts of a pivot whose records
 * are stale or absent. Values are left empty. Undefined when the saved lines cannot be laid out.
 */
export function savedCaptionLayout(def: PivotTableDefinition, cache: PivotCacheDefinition, date1904: boolean): PivotLayout | undefined {
  const saved = savedPivotLayout(def)
  if (!saved) return undefined
  const axes = { rows: saved.rows, cols: saved.cols, rowKey: new Uint32Array(0), colKey: new Uint32Array(0) }
  return placeCells(def, initialViewState(def, cache), cache, axes, { get: () => ({ kind: 'empty' }) }, { date1904 })
}

/**
 * The retained bytes of a refreshed cache's rebuilt metadata, which the record store's `bytes` does not count: each
 * listed field's item objects and arrays (their text counted too: on the paged path no sheet holds it), plus a fixed
 * share for the field and cache objects. Charged with the store for the session's life (codex re-review T32-R2-2).
 */
export function refreshedCacheBytes(cache: PivotCacheDefinition): number {
  let bytes = 512
  for (const field of cache.fields) {
    bytes += 256
    for (const item of field.sharedItems?.items ?? []) bytes += ITEM_BYTES + (item.type === 'string' || item.type === 'error' || item.type === 'date' ? 2 * String(item.value).length : 0)
  }
  return bytes
}
