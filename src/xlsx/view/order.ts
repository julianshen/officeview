import type { RowOrder, RowPermutationBand, RowVisibility } from './contracts'
import { MAX_ROWS, validateCoordinate, type DisplaySlot, type SourceRow } from './coordinates'

export interface PermutationInput { r0: SourceRow; r1: SourceRow; sourceAtOffset: Uint32Array }
/** Runtime row order. Adapters to/from the `ViewCommitDelta` wire encoding are owned by
 * S4.1; nothing here serializes. */
export interface CompactRowOrder extends RowOrder {
  readonly diagnostics: { typedArrayBytes: number; bandCount: number }
}
function rowIndex(row: number): number { validateCoordinate(row, 0); return row }
function validateBand(r0: number, r1: number): void {
  rowIndex(r0)
  rowIndex(r1)
  if (r0 > r1) throw new RangeError('Reversed permutation band')
}
const identity: CompactRowOrder = Object.freeze({
  kind: 'identity', bands: Object.freeze([]),
  displayToSource: rowIndex, sourceToDisplay: rowIndex,
  diagnostics: Object.freeze({ typedArrayBytes: 0, bandCount: 0 }),
})
/** Lazy full-axis identity has no row storage (§5.1). */
export function createIdentityOrder(): CompactRowOrder { return identity }

function makeBand(input: PermutationInput): RowPermutationBand {
  validateBand(input.r0, input.r1)
  const length = input.r1 - input.r0 + 1
  if (!(input.sourceAtOffset instanceof Uint32Array) || input.sourceAtOffset.length !== length) {
    throw new RangeError('Permutation length must match its inclusive band')
  }
  const sourceAtOffset = input.sourceAtOffset.slice()
  const displayOffsetOfSource = new Uint32Array(length).fill(MAX_ROWS)
  for (let displayOffset = 0; displayOffset < length; displayOffset++) {
    const sourceRow = sourceAtOffset[displayOffset]
    if (sourceRow < input.r0 || sourceRow > input.r1 || displayOffsetOfSource[sourceRow - input.r0] !== MAX_ROWS) {
      throw new RangeError('Permutation must be a bijection of its band')
    }
    displayOffsetOfSource[sourceRow - input.r0] = displayOffset
  }
  return Object.freeze({ r0: input.r0, r1: input.r1, sourceAtOffset, displayOffsetOfSource })
}
/** Validate and copy every band. Module-level on purpose: a closure created in the same
 * scope as `inputs`/`sorted` would share their context and keep the caller's buffers
 * alive for the lifetime of the order (review item 2). */
function buildBands(inputs: readonly PermutationInput[]): readonly RowPermutationBand[] {
  const sorted = [...inputs].sort((a, b) => a.r0 - b.r0)
  for (let index = 0; index < sorted.length; index++) {
    validateBand(sorted[index].r0, sorted[index].r1)
    if (index && sorted[index - 1].r1 >= sorted[index].r0) throw new RangeError('Overlapping permutation bands')
  }
  return Object.freeze(sorted.map(makeBand))
}
function bandAt(bands: readonly RowPermutationBand[], row: number): RowPermutationBand | undefined {
  rowIndex(row)
  let lo = 0, hi = bands.length
  while (lo < hi) {
    const mid = Math.floor((lo + hi) / 2)
    if (bands[mid].r1 < row) lo = mid + 1
    else hi = mid
  }
  return lo < bands.length && bands[lo].r0 <= row ? bands[lo] : undefined
}
function orderOf(bands: readonly RowPermutationBand[]): CompactRowOrder {
  return Object.freeze({
    kind: 'bands', bands,
    displayToSource(displaySlot: DisplaySlot) {
      const band = bandAt(bands, displaySlot)
      return band ? band.sourceAtOffset[displaySlot - band.r0] : displaySlot
    },
    sourceToDisplay(sourceRow: SourceRow) {
      const band = bandAt(bands, sourceRow)
      return band ? band.r0 + band.displayOffsetOfSource[sourceRow - band.r0] : sourceRow
    },
    diagnostics: Object.freeze({ bandCount: bands.length, typedArrayBytes: bands.reduce((n, band) => n + band.sourceAtOffset.byteLength + band.displayOffsetOfSource.byteLength, 0) }),
  })
}
/** Input buffers are copied once and never retained; published buffers belong to this
 * immutable view generation and must not be mutated by consumers (§§5.1,10.1). */
export function createRowOrder(inputs: readonly PermutationInput[]): CompactRowOrder {
  return inputs.length ? orderOf(buildBands(inputs)) : identity
}
// Intrinsic getters read internal slots: realm-independent (wire bands are
// structured clones) and immune to own-property or Symbol.toStringTag overrides.
const typedArrayProto = Object.getPrototypeOf(Uint8Array.prototype) as object
const intrinsic = (proto: object, key: PropertyKey) => Object.getOwnPropertyDescriptor(proto, key)?.get
const typedArrayTag = intrinsic(typedArrayProto, Symbol.toStringTag)!
const typedArrayBuffer = intrinsic(typedArrayProto, 'buffer')!
const typedArrayByteOffset = intrinsic(typedArrayProto, 'byteOffset')!
const typedArrayByteLength = intrinsic(typedArrayProto, 'byteLength')!
const typedArrayLength = intrinsic(typedArrayProto, 'length')!
const arrayBufferByteLength = intrinsic(ArrayBuffer.prototype, 'byteLength')!
const arrayBufferResizable = intrinsic(ArrayBuffer.prototype, 'resizable')
/** A Uint32Array of any realm, by its internal [[TypedArrayName]]. */
export const isUint32 = (value: unknown): value is Uint32Array => ArrayBuffer.isView(value) && Reflect.apply(typedArrayTag, value, []) === 'Uint32Array'
/** Re-views a Uint32Array of any realm as this realm's, over the same bytes. */
export const localUint32 = (array: Uint32Array): Uint32Array => array instanceof Uint32Array ? array : new Uint32Array(Reflect.apply(typedArrayBuffer, array, []) as ArrayBufferLike, Reflect.apply(typedArrayByteOffset, array, []) as number, Reflect.apply(typedArrayLength, array, []) as number)
/** The view's real backing store when it is a fixed-length, non-shared
 * ArrayBuffer that the view spans exactly; undefined otherwise. */
function ownedBuffer(array: Uint32Array): ArrayBuffer | undefined {
  const buffer = Reflect.apply(typedArrayBuffer, array, []) as ArrayBuffer
  let byteLength: number
  // The ArrayBuffer getters throw on a SharedArrayBuffer.
  try { byteLength = Reflect.apply(arrayBufferByteLength, buffer, []) as number } catch { return undefined }
  // A resizable buffer lets any holder change a length-tracking view later.
  if (arrayBufferResizable && Reflect.apply(arrayBufferResizable, buffer, []) === true) return undefined
  return Reflect.apply(typedArrayByteOffset, array, []) === 0 && Reflect.apply(typedArrayByteLength, array, []) === byteLength ? buffer : undefined
}
/**
 * Zero-copy order over wire bands that already carry their inverse (a worker
 * view delta): the bands' own buffers become this order's, after validating
 * that each band covers its inclusive range and that displayOffsetOfSource is
 * the exact inverse of sourceAtOffset (which also proves a bijection). The
 * caller hands the buffers over and must not keep using them. Throws a
 * RangeError on any invalid band (callers fall back to createRowOrder).
 */
export function adoptRowOrder(bands: readonly RowPermutationBand[]): CompactRowOrder {
  if (!bands.length) return identity
  const sorted = [...bands].sort((a, b) => a.r0 - b.r0)
  const out: RowPermutationBand[] = []
  const buffers = new Set<ArrayBufferLike>()
  for (let index = 0; index < sorted.length; index++) {
    const band = sorted[index]
    validateBand(band.r0, band.r1)
    if (index && sorted[index - 1].r1 >= band.r0) throw new RangeError('Overlapping permutation bands')
    const length = band.r1 - band.r0 + 1
    if (!isUint32(band.sourceAtOffset) || !isUint32(band.displayOffsetOfSource) || Reflect.apply(typedArrayLength, band.sourceAtOffset, []) !== length || Reflect.apply(typedArrayLength, band.displayOffsetOfSource, []) !== length) {
      throw new RangeError('Permutation length must match its inclusive band')
    }
    // Shared storage cannot be handed over exclusively, even across bands.
    // A partial view would keep unaccounted backing-store padding alive.
    for (const array of [band.sourceAtOffset, band.displayOffsetOfSource]) {
      const buffer = ownedBuffer(array)
      if (!buffer || buffers.has(buffer)) throw new RangeError('Permutation buffers must be distinct, non-shared and fully owned')
      buffers.add(buffer)
    }
    // A structured clone from another realm is re-viewed (same buffer, no copy) as
    // this realm's Uint32Array, which projection and its consumers require.
    const sourceAtOffset = localUint32(band.sourceAtOffset), displayOffsetOfSource = localUint32(band.displayOffsetOfSource)
    for (let offset = 0; offset < length; offset++) {
      const sourceRow = sourceAtOffset[offset]
      if (sourceRow < band.r0 || sourceRow > band.r1 || displayOffsetOfSource[sourceRow - band.r0] !== offset) throw new RangeError('Permutation must be a bijection of its band')
    }
    out.push(Object.freeze({ r0: band.r0, r1: band.r1, sourceAtOffset, displayOffsetOfSource }))
  }
  return orderOf(Object.freeze(out))
}
/** True when a DISPLAY slot must keep its source-corresponding row (header/totals rows,
 * hidden rows). This is deliberately a callback and not a RowVisibility: a visibility mask
 * answers "visible?", the opposite question, so passing one by mistake must not compile. */
export type IsPinnedSlot = (displaySlot: DisplaySlot) => boolean
/** Pin exactly the hidden slots of a display-visibility mask (visible rows are movable). */
export function hiddenSlotPins(display: RowVisibility): IsPinnedSlot {
  return displaySlot => !display.visible(displaySlot)
}
/** Pinned slots keep their own row; `candidates` lists only the movable (unpinned) source
 * rows, in their new display order, and fills the remaining slots left to right (§5.1). */
export function pinRowPermutation(r0: SourceRow, r1: SourceRow, isPinned: IsPinnedSlot, candidates: Uint32Array): RowPermutationBand {
  validateBand(r0, r1)
  if (typeof isPinned !== 'function') throw new TypeError('isPinned must be a callback; use hiddenSlotPins(visibility) for a mask')
  if (!(candidates instanceof Uint32Array)) throw new RangeError('Candidates must be Uint32 source rows')
  const length = r1 - r0 + 1
  let pinnedCount = 0
  for (let displaySlot = r0; displaySlot <= r1; displaySlot++) if (isPinned(displaySlot)) pinnedCount++
  if (candidates.length !== length - pinnedCount) throw new RangeError('Wrong movable candidate count')
  const rows = new Uint32Array(length)
  let candidateOffset = 0
  for (let displaySlot = r0; displaySlot <= r1; displaySlot++) {
    const pinned = isPinned(displaySlot)
    const sourceRow = pinned ? displaySlot : candidates[candidateOffset++]
    if (sourceRow < r0 || sourceRow > r1 || (sourceRow !== displaySlot && isPinned(sourceRow))) throw new RangeError('A pinned/outside source row cannot move')
    rows[displaySlot - r0] = sourceRow
  }
  return makeBand({ r0, r1, sourceAtOffset: rows })
}
