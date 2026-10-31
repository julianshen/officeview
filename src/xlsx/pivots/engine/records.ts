/** Columnar pivot-cache record decoder (Stage 2, kept in the lazy pivots chunk). */
import { tokenizeBytes, type XmlEvents } from '../../../core/xml-stream'
import type { PivotCacheDefinition, PivotDiagnostic, PivotValue } from '../model'
import type { ProfileId } from '../../view/contracts'
import { daysFromCivil } from '../../formula/serial'

export interface RecordLimits { maxBytes: number; maxRecords: number; maxInlineStrings: number }
export const RECORD_LIMITS: Readonly<Record<ProfileId, RecordLimits>> = Object.freeze({
  D: Object.freeze({ maxBytes: 64 * 1024 * 1024, maxRecords: 4_000_000, maxInlineStrings: 1_048_576 }),
  M: Object.freeze({ maxBytes: 16 * 1024 * 1024, maxRecords: 1_000_000, maxInlineStrings: 262_144 }),
})

export type RecordColumn =
  | { kind: 'shared'; index: Uint32Array }
  | { kind: 'inline'; tag: Uint8Array; num: Float64Array; text: string[]; textIndex: Uint32Array }
export const VALUE_TAG = Object.freeze({ missing: 0, number: 1, boolean: 2, error: 3, string: 4, date: 5 } as const)
export interface RecordStore { readonly count: number; readonly columns: ReadonlyArray<RecordColumn | undefined>; readonly bytes: number }
export type RecordResult = { ok: true; store: RecordStore } | { ok: false; diagnostic: PivotDiagnostic }

const STRING_BYTES = 48 // retained JS string entry plus UTF-16 payload; matches the source-page estimate.
const localName = (name: string) => name.slice(name.lastIndexOf(':') + 1)
const parseError = (message: string): PivotDiagnostic => ({ kind: 'parse', feature: 'pivot-records', message })
const limitError = (message: string): PivotDiagnostic => ({ kind: 'unsupported', feature: 'pivot-records-limit', message })
class StopRecords extends Error {}
type InlineValue = { tag: number; num: number; text?: string }

function valueFrom(kind: string, raw: string | undefined, date1904: boolean, fail: (message: string) => never): PivotValue {
  if (kind === 'm') return { type: 'missing' }
  if (raw === undefined) fail('record value is missing its v attribute')
  if (kind === 'n') {
    if (!/^[+-]?(?:(?:\d+(?:\.\d*)?)|(?:\.\d+))(?:[Ee][+-]?\d+)?$/.test(raw)) fail('record number is invalid')
    const value = Number(raw)
    if (!Number.isFinite(value)) fail('record number is invalid')
    return { type: 'number', value }
  }
  if (kind === 'b') {
    if (raw === '1' || raw === 'true') return { type: 'boolean', value: true }
    if (raw === '0' || raw === 'false') return { type: 'boolean', value: false }
    fail('record boolean is invalid')
  }
  if (kind === 'e') return { type: 'error', value: raw }
  if (kind === 's') return { type: 'string', value: raw }
  if (kind === 'd') {
    if (!Number.isFinite(isoToSerial(raw, date1904))) fail('record date is invalid')
    return { type: 'date', value: raw }
  }
  fail(`unknown record value element ${kind}`)
}

/** Convert SpreadsheetML ISO dateTime to an Excel serial, preserving the time fraction. */
export function isoToSerial(iso: string, date1904: boolean): number {
  const m = /^(\d{4})-(\d{2})-(\d{2})(?:T(\d{2}):(\d{2}):(\d{2})(?:\.(\d+))?)?(?:Z|([+-])(\d{2}):(\d{2}))?$/.exec(iso)
  if (!m) return Number.NaN
  const year = Number(m[1]), month = Number(m[2]), day = Number(m[3])
  const hour = Number(m[4] ?? 0), minute = Number(m[5] ?? 0), second = Number(m[6] ?? 0)
  if (hour > 23 || minute > 59 || second > 59 || (m[7] && !/^\d+$/.test(m[7]))) return Number.NaN
  if (month < 1 || month > 12) return Number.NaN
  const leap = year % 4 === 0 && (year % 100 !== 0 || year % 400 === 0)
  const monthDays = [31, leap ? 29 : 28, 31, 30, 31, 30, 31, 31, 30, 31, 30, 31]
  if (day < 1 || day > monthDays[month - 1]!) return Number.NaN
  const epoch = date1904 ? daysFromCivil(1904, 1, 1) : daysFromCivil(1899, 12, 30)
  const dayOffset = daysFromCivil(year, month, day) - epoch
  if (dayOffset < 0) return Number.NaN
  // Task 15 Excel oracle: saved ISO cache dates use the uniform 1899-12-30
  // epoch. Thus 1900-02-28 is serial 60; do not insert Excel's phantom day.
  const daySerial = dayOffset
  let fraction = (hour * 3600 + minute * 60 + second) / 86400
  if (m[7]) fraction += Number(`0.${m[7]}`) / 86400
  if (m[8]) {
    const zoneHours = Number(m[9]), zoneMinutes = Number(m[10])
    if (zoneHours > 14 || zoneMinutes > 59 || (zoneHours === 14 && zoneMinutes !== 0)) return Number.NaN
    const offset = (zoneHours * 60 + zoneMinutes) * (m[8] === '+' ? 1 : -1)
    fraction -= offset / 1440
  }
  let serial = daySerial + fraction
  if (serial < 0) return Number.NaN
  // Time-zone rollover across a civil day uses the same serial arithmetic.
  return serial
}

interface MutableColumn {
  kind: 'shared' | 'inline'
  index?: Uint32Array
  tag?: Uint8Array
  num?: Float64Array
  textIndex?: Uint32Array
  text?: string[]
  textLookup?: Map<string, number>
}

/** A one-pass SAX event sink. It stores only typed columns and bounded inline text. */
export function recordEvents(cache: PivotCacheDefinition, limits: RecordLimits, date1904: boolean): { events: XmlEvents; finish(): RecordResult } {
  let diagnostic: PivotDiagnostic | undefined
  const dbFields = cache.fields.flatMap((field, index) => field.databaseField ? [index] : [])
  const columns: Array<MutableColumn | undefined> = cache.fields.map(field => !field.databaseField ? undefined : field.sharedItems?.inline === false
    ? { kind: 'shared', index: new Uint32Array(16) }
    : { kind: 'inline', tag: new Uint8Array(16), num: new Float64Array(16), textIndex: new Uint32Array(16), text: [], textLookup: new Map() })
  let count = 0, capacity = 16, rootSeen = false, rootClosed = false, declaredRecordCount: number | undefined, row: Array<number | InlineValue> | undefined
  let rowChildCount = 0
  const stack: string[] = []
  let usedTextBytes = 0
  const fail = (message: string, isLimit = false): never => {
    diagnostic ??= isLimit ? limitError(message) : parseError(message)
    throw new StopRecords(message)
  }
  const attr = (names: readonly string[], values: readonly string[], n: string): string | undefined => {
    for (let i = 0; i < names.length; i++) if (localName(names[i]!) === n) return values[i]
    return undefined
  }
  const checkInputState = () => { if (diagnostic) throw new StopRecords(diagnostic.message) }
  const ensureCapacity = (needed: number) => {
    if (needed <= capacity) return
    let next = capacity
    while (next < needed) next = Math.min(Math.max(next * 2, needed), limits.maxRecords + 1)
    if (next < needed || !Number.isSafeInteger(next)) fail('pivot record count exceeds its limit', true)
    for (const col of columns) {
      if (!col) continue
      if (col.kind === 'shared') { const a = new Uint32Array(next); a.set(col.index!.subarray(0, count)); col.index = a }
      else {
        const t = new Uint8Array(next), n = new Float64Array(next), ix = new Uint32Array(next)
        t.set(col.tag!.subarray(0, count)); n.set(col.num!.subarray(0, count)); ix.set(col.textIndex!.subarray(0, count))
        col.tag = t; col.num = n; col.textIndex = ix
      }
    }
    capacity = next
  }
  const storeBytesAfterRow = (nextCount: number, newTextBytes: number) => {
    let n = newTextBytes
    for (const col of columns) if (col) n += nextCount * (col.kind === 'shared' ? 4 : 13)
    return n
  }
  const appendRow = () => {
    if (!row) return
    if (rowChildCount !== dbFields.length) fail(`record has ${rowChildCount} values for ${dbFields.length} database fields`)
    if (count >= limits.maxRecords) fail('pivot record count exceeds its limit', true)
    const pendingText = new Map<number, Set<string>>()
    let addedTextBytes = 0
    for (let i = 0; i < dbFields.length; i++) {
      const cacheIndex = dbFields[i]!, col = columns[cacheIndex]!, value = row[i]!
      if (col.kind !== 'inline' || typeof value === 'number' || value.tag === VALUE_TAG.missing || value.text === undefined) continue
      if (col.textLookup!.has(value.text)) continue
      let pending = pendingText.get(cacheIndex)
      if (!pending) pendingText.set(cacheIndex, pending = new Set())
      if (!pending.has(value.text)) { pending.add(value.text); addedTextBytes += STRING_BYTES + value.text.length * 2 }
    }
    for (const [fieldIndex, strings] of pendingText) {
      const col = columns[fieldIndex]!
      if (col.kind === 'inline' && col.text!.length + strings.size > limits.maxInlineStrings) fail(`pivot inline string count exceeds its limit in field ${fieldIndex}`, true)
    }
    if (storeBytesAfterRow(count + 1, usedTextBytes + addedTextBytes) > limits.maxBytes) fail('pivot record storage exceeds its byte limit', true)
    ensureCapacity(count + 1)
    for (let i = 0; i < dbFields.length; i++) {
      const cacheIndex = dbFields[i]!, col = columns[cacheIndex]!, value = row[i]!
      if (col.kind === 'shared') col.index![count] = value as number
      else {
        const v = value as InlineValue
        col.tag![count] = v.tag; col.num![count] = v.num
        if (v.text !== undefined) {
          let index = col.textLookup!.get(v.text)
          if (index === undefined) {
            index = col.text!.length; col.textLookup!.set(v.text, index); col.text!.push(v.text)
            usedTextBytes += STRING_BYTES + v.text.length * 2
          }
          col.textIndex![count] = index
        }
      }
    }
    count++
    row = undefined; rowChildCount = 0
  }
  const events: XmlEvents = {
    open(qname, names, values, _count, selfClosing) {
      checkInputState()
      const tag = localName(qname), parent = stack[stack.length - 1]
      if (!rootSeen) {
        if (tag !== 'pivotCacheRecords') fail('pivot records root is missing')
        rootSeen = true
        const declared = attr(names, values, 'count')
        if (declared !== undefined && !/^\d+$/.test(declared)) fail('pivot records count is invalid')
        const declaredNumber = declared === undefined ? undefined : Number(declared)
        if (declaredNumber !== undefined && !Number.isSafeInteger(declaredNumber)) fail('pivot records count is invalid')
        if (declaredNumber !== undefined && declaredNumber > limits.maxRecords) fail('pivot record count exceeds its limit', true)
        if (declaredNumber !== undefined) declaredRecordCount = declaredNumber
      } else if (parent === 'pivotCacheRecords' && tag === 'r') {
        if (row) fail('nested pivot record')
        row = []; rowChildCount = 0
      } else if (parent === 'r' && row) {
        if (rowChildCount >= dbFields.length) fail('record has more values than database fields')
        const fieldIndex = dbFields[rowChildCount]!, col = columns[fieldIndex]!, rawKind = tag
        const raw = attr(names, values, 'v')
        if (tag === 'x') {
          if (col.kind !== 'shared') fail('shared index appears in an inline field')
          if (raw === undefined || !/^\d+$/.test(raw)) fail('shared index is invalid')
          const index = Number(raw), length = cache.fields[fieldIndex]!.sharedItems?.items.length ?? 0
          if (!Number.isSafeInteger(index) || index >= length) fail(`shared index ${raw} is outside field ${fieldIndex}`)
          row.push(index)
        } else {
          if (col.kind !== 'inline') fail('inline value appears in a shared field')
          const value: PivotValue = valueFrom(rawKind, raw, date1904, fail)
          const tagValue = value.type === 'missing' ? VALUE_TAG.missing : value.type === 'number' ? VALUE_TAG.number : value.type === 'boolean' ? VALUE_TAG.boolean : value.type === 'error' ? VALUE_TAG.error : value.type === 'string' ? VALUE_TAG.string : VALUE_TAG.date
          const inline: InlineValue = { tag: tagValue, num: value.type === 'number' ? value.value : value.type === 'boolean' ? Number(value.value) : value.type === 'date' ? isoToSerial(value.value, date1904) : 0 }
          if (value.type === 'string' || value.type === 'error') inline.text = value.value
          row.push(inline)
        }
        rowChildCount++
      } else if (parent === 'pivotCacheRecords') fail(`unexpected element ${tag} in pivot records`)
      else if (parent === 'r' || parent) fail(`unexpected nested element ${tag} in pivot records`)
      if (!selfClosing) stack.push(tag)
      else if (tag === 'r') appendRow()
      else if (tag === 'pivotCacheRecords') rootClosed = true
    },
    close(qname) {
      checkInputState()
      const tag = localName(qname)
      if (stack.pop() !== tag) fail('pivot records XML nesting is invalid')
      if (tag === 'r') appendRow()
      if (tag === 'pivotCacheRecords') rootClosed = true
    },
    text(text) { checkInputState(); if (text.trim() !== '') fail('pivot records contain unexpected text') },
  }
  return {
    events,
    finish(): RecordResult {
      if (diagnostic) return { ok: false, diagnostic }
      if (!rootSeen || !rootClosed || row || stack.length) return { ok: false, diagnostic: parseError('pivot records XML is incomplete') }
      if (cache.recordCount !== undefined && cache.recordCount !== count) return { ok: false, diagnostic: parseError(`pivot record count ${count} does not match cache count ${cache.recordCount}`) }
      if (declaredRecordCount !== undefined && declaredRecordCount !== count) return { ok: false, diagnostic: parseError(`pivot records count ${count} does not match declared count ${declaredRecordCount}`) }
      const resultColumns = columns.map(col => {
        if (!col) return undefined
        if (col.kind === 'shared') return { kind: 'shared' as const, index: col.index!.slice(0, count) }
        return { kind: 'inline' as const, tag: col.tag!.slice(0, count), num: col.num!.slice(0, count), text: col.text!, textIndex: col.textIndex!.slice(0, count) }
      })
      const bytes = resultColumns.reduce((sum, col) => sum + (!col ? 0 : col.kind === 'shared' ? col.index.byteLength : col.tag.byteLength + col.num.byteLength + col.textIndex.byteLength + col.text.reduce((n, s) => n + STRING_BYTES + s.length * 2, 0)), 0)
      return { ok: true, store: { count, columns: resultColumns, bytes } }
    },
  }
}

/** Streams records without reading ahead beyond the chunk containing a refusal. */
export async function decodeRecords(source: AsyncIterable<Uint8Array>, cache: PivotCacheDefinition, limits: RecordLimits, date1904: boolean, signal?: AbortSignal): Promise<RecordResult> {
  const decoder = recordEvents(cache, limits, date1904)
  try { await tokenizeBytes(source, decoder.events, { signal }) }
  catch (error) {
    if (!(error instanceof StopRecords)) return { ok: false, diagnostic: parseError(error instanceof Error ? error.message : String(error)) }
  }
  return decoder.finish()
}
