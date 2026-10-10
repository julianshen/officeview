import type { RowOrder, RowPermutationBand, RowVisibility } from './contracts'
import { MAX_ROWS, validateCoordinate, type DisplaySlot, type SourceRow } from './coordinates'

export interface PermutationInput { r0: SourceRow; r1: SourceRow; sourceAtOffset: Uint32Array }
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
/** Input buffers are copied once; published buffers belong to this immutable view
 * generation and must not be mutated by consumers (§§5.1,10.1). */
export function createRowOrder(inputs: readonly PermutationInput[]): CompactRowOrder {
  if (!inputs.length) return identity
  const sorted = [...inputs].sort((a, b) => a.r0 - b.r0)
  sorted.forEach((band, index) => {
    validateBand(band.r0, band.r1)
    if (index && sorted[index - 1].r1 >= band.r0) throw new RangeError('Overlapping permutation bands')
  })
  const bands = Object.freeze(sorted.map(makeBand))
  const lookup = (row: number): RowPermutationBand | undefined => {
    rowIndex(row)
    let lo = 0, hi = bands.length
    while (lo < hi) {
      const mid = Math.floor((lo + hi) / 2)
      if (bands[mid].r1 < row) lo = mid + 1
      else hi = mid
    }
    return lo < bands.length && bands[lo].r0 <= row ? bands[lo] : undefined
  }
  return Object.freeze({
    kind: 'bands', bands,
    displayToSource(displaySlot: DisplaySlot) {
      const band = lookup(displaySlot)
      return band ? band.sourceAtOffset[displaySlot - band.r0] : displaySlot
    },
    sourceToDisplay(sourceRow: SourceRow) {
      const band = lookup(sourceRow)
      return band ? band.r0 + band.displayOffsetOfSource[sourceRow - band.r0] : sourceRow
    },
    diagnostics: Object.freeze({ bandCount: bands.length, typedArrayBytes: bands.reduce((n, band) => n + band.sourceAtOffset.byteLength + band.displayOffsetOfSource.byteLength, 0) }),
  })
}
/** A true pin retains its source-corresponding display slot. Include protected
 * header/totals and hidden rows in pins; candidates contain only movable rows (§5.1). */
export function pinRowPermutation(r0: SourceRow, r1: SourceRow, pins: RowVisibility, candidates: Uint32Array): RowPermutationBand {
  validateBand(r0, r1)
  if (!(candidates instanceof Uint32Array)) throw new RangeError('Candidates must be Uint32 source rows')
  const length = r1 - r0 + 1
  if (candidates.length !== length - (pins.rank(r1 + 1) - pins.rank(r0))) throw new RangeError('Wrong visible candidate count')
  const rows = new Uint32Array(length)
  let candidateOffset = 0
  for (let displaySlot = r0; displaySlot <= r1; displaySlot++) {
    const sourceRow = pins.visible(displaySlot) ? displaySlot : candidates[candidateOffset++]
    if (sourceRow < r0 || sourceRow > r1 || (sourceRow !== displaySlot && pins.visible(sourceRow))) throw new RangeError('A pinned/outside source row cannot move')
    rows[displaySlot - r0] = sourceRow
  }
  return makeBand({ r0, r1, sourceAtOffset: rows })
}
