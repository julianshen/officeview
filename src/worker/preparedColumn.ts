import type { CooperativeSlice } from '../xlsx/view/scheduler'

export const F_CACHED_SET = 1, F_CACHED = 2, F_ERROR_SET = 4, F_ERROR = 8, F_PROV_ENGINE = 16, F_PROV_CACHED = 32, F_UNAVAILABLE = 64
export const TAG_NULL = 0, TAG_NUMBER = 1, TAG_STRING = 2, TAG_BOOLEAN = 3
export const PREPARED_PAGE_ENTRIES = 4096
export const PREPARED_PAGE_OVERHEAD = 256
/** Scratch descriptors and a bounded sixteen-entry metadata dictionary. */
export const PREPARED_ENCODING_OVERHEAD = 1024
export type NumericLane = Uint8Array | Uint16Array | Uint32Array | Int32Array | Float32Array | Float64Array
export interface RawPreparedPage {
  count: number
  rows: Uint32Array
  tags: Uint8Array
  numbers: Float64Array
  stringIds: Uint32Array
  styles: Uint32Array
  flags: Uint8Array
}
export interface PreparedPage {
  count: number
  firstRow: number
  rows?: Uint32Array
  values: NumericLane
  styles: Uint32Array
  codes: Uint16Array
  metadataIds?: Uint8Array
  flatMetadata: boolean
  plainText: boolean
  laneBytes: number
}
const valueAt = (raw: RawPreparedPage, at: number) => raw.tags[at] === TAG_STRING ? raw.stringIds[at] : raw.numbers[at]
const codeAt = (raw: RawPreparedPage, at: number) => raw.flags[at] * 4 + raw.tags[at]

/** Bounded lossless encoder. Its caller reserves old lanes, output (at most
 * 18 bytes/entry), page overhead and scratch BEFORE entering this boundary.
 * `undefined` means ownership/cancellation changed at a cooperative pause. */
export async function compactPreparedPage(raw: RawPreparedPage, slice: CooperativeSlice<string>, halted: () => boolean): Promise<PreparedPage | undefined> {
  const n = raw.count, firstRow = raw.rows[0]
  let denseRows = true, unsigned = true, signed = true, float32 = true, max = 0, plainText = false
  const styleKeys: number[] = [], codeKeys: number[] = []
  let flatMetadata = false
  for (let at = 0; at < n; at++) {
    const value = valueAt(raw, at)
    if (raw.tags[at] === TAG_STRING && !(raw.flags[at] & (F_ERROR | F_UNAVAILABLE))) plainText = true
    if (raw.rows[at] !== firstRow + at) denseRows = false
    if (!Number.isInteger(value) || value < 0 || value > 4294967295 || Object.is(value, -0)) unsigned = false
    if (!Number.isInteger(value) || value < -2147483648 || value > 2147483647 || Object.is(value, -0)) signed = false
    if (!Object.is(Math.fround(value), value)) float32 = false
    max = Math.max(max, value)
    if (!flatMetadata) {
      const style = raw.styles[at], code = codeAt(raw, at)
      let id = 0
      while (id < styleKeys.length && (styleKeys[id] !== style || codeKeys[id] !== code)) id++
      if (id === styleKeys.length) {
        if (id === 16) flatMetadata = true
        else { styleKeys.push(style); codeKeys.push(code) }
      }
    }
    if (slice.due()) { await slice.pause('prepare'); if (halted()) return undefined }
  }
  const values: NumericLane = unsigned ? max <= 255 ? new Uint8Array(n) : max <= 65535 ? new Uint16Array(n) : new Uint32Array(n)
    : signed ? new Int32Array(n) : float32 ? new Float32Array(n) : new Float64Array(n)
  const rows = denseRows ? undefined : raw.rows.slice(0, n)
  const styles = flatMetadata ? new Uint32Array(n) : Uint32Array.from(styleKeys)
  const codes = flatMetadata ? new Uint16Array(n) : Uint16Array.from(codeKeys)
  const metadataIds = !flatMetadata && styleKeys.length > 1 ? new Uint8Array(n) : undefined
  for (let at = 0; at < n; at++) {
    values[at] = valueAt(raw, at)
    const style = raw.styles[at], code = codeAt(raw, at)
    if (flatMetadata) { styles[at] = style; codes[at] = code }
    else if (metadataIds) {
      let id = 0
      while (styleKeys[id] !== style || codeKeys[id] !== code) id++
      metadataIds[at] = id
    }
    if (slice.due()) { await slice.pause('prepare'); if (halted()) return undefined }
  }
  const laneBytes = (rows?.byteLength ?? 0) + values.byteLength + styles.byteLength + codes.byteLength + (metadataIds?.byteLength ?? 0)
  return Object.freeze({ count: n, firstRow, rows, values, styles, codes, metadataIds, flatMetadata, plainText, laneBytes })
}
export function preparedRowAt(page: PreparedPage, at: number): number { return page.rows ? page.rows[at] : page.firstRow + at }
export function preparedMetadataAt(page: PreparedPage, at: number): number { return page.flatMetadata ? at : page.metadataIds?.[at] ?? 0 }

/** Visits a bounded page of populated ordinals and blank spans. */
export function visitPreparedRows(
  count: number, rowAt: (ordinal: number) => number,
  intervals: readonly { r0: number; r1: number }[], limit: number,
  cell: (ordinal: number) => void, blank: (r0: number, r1: number) => void,
): void {
  let emitted = 0
  for (const { r0, r1 } of intervals) {
    let low = 0, high = count
    while (low < high) {
      const mid = Math.floor((low + high) / 2)
      if (rowAt(mid) < r0) low = mid + 1
      else high = mid
    }
    let row = r0
    while (row <= r1 && emitted < limit) {
      const next = low < count ? rowAt(low) : Infinity
      if (next > r1) { blank(row, r1); emitted++; break }
      if (next > row) { blank(row, next - 1); row = next }
      else { cell(low++); row++ }
      emitted++
    }
    if (emitted === limit) break
  }
}
