import { checkBitmap, type Bitmap, type Chip, type MarkParams } from './types'
import { checkParams } from './params'
export interface MarkedBitmap extends Bitmap {
  chipDelta: Float32Array
  chipPedestal: Float32Array // NaN denotes a masked chip
  chipSupport: Chip[]
  mark: { variant: 'white-pedestal' | 'direct'; pedestalCells: number; maskedCells: number; params: MarkParams }
}
/** Equal-channel shift admitting BOTH signs in every channel. Null masks the
 * cell. Shift magnitude is bounded by the explicit pedestal parameter. */
export function pedestalFor(referenceCell: Bitmap, params: MarkParams): number | null {
  checkBitmap(referenceCell); checkParams(params)
  let min = 255, max = 0
  for (let i = 0; i < referenceCell.data.length; i += 4) {
    if (referenceCell.data[i + 3] !== 255) return null
    for (let c = 0; c < 3; c++) { min = Math.min(min, referenceCell.data[i + c]!); max = Math.max(max, referenceCell.data[i + c]!) }
  }
  const lo = params.amplitude - min, hi = 255 - params.amplitude - max
  if (lo > hi) return null
  const shift = Math.max(lo, Math.min(hi, 0))
  return Math.abs(shift) <= params.pedestal ? shift : null
}
export function referenceCell(bmp: Bitmap, chip: Chip): Bitmap {
  const data = new Uint8ClampedArray(chip.cell * chip.cell * 4)
  for (let y = 0; y < chip.cell; y++) data.set(bmp.data.subarray(((chip.y + y) * bmp.width + chip.x) * 4, ((chip.y + y) * bmp.width + chip.x + chip.cell) * 4), y * chip.cell * 4)
  return { width: chip.cell, height: chip.cell, data }
}
/** Code signs are multiplied by the page-keyed carrier signs. Longer bases
 * cycle the extended 512-symbol word (not repetition of the 40-bit payload).
 * Transparent cells are masked; alpha is preserved, input is never mutated. */
export function embed(bmp: Bitmap, chips: Chip[], code: Int8Array, params: MarkParams): MarkedBitmap {
  checkBitmap(bmp); checkParams(params)
  if (chips.length !== params.nChips || code.length !== 512 || code.some(x => x !== -1 && x !== 1)) throw new Error('Expected matching chips and 512 signed symbols')
  // Validate complete support before copying/painting; disjoint grid cells.
  const occupied = new Set<string>()
  for (const chip of chips) {
    if (!Number.isInteger(chip.x) || !Number.isInteger(chip.y) || chip.x < 0 || chip.y < 0 || chip.cell !== params.cellPx || chip.x + chip.cell > bmp.width || chip.y + chip.cell > bmp.height || (chip.sign0 !== -1 && chip.sign0 !== 1)) throw new Error('Invalid chip')
    for (let y = chip.y; y < chip.y + chip.cell; y++) for (let x = chip.x; x < chip.x + chip.cell; x++) {
      const id = `${x},${y}`; if (occupied.has(id)) throw new Error('Overlapping chips'); occupied.add(id)
    }
  }
  const data = bmp.data.slice()
  let pedestalCells = 0, maskedCells = 0
  const chipDelta = new Float32Array(chips.length), chipPedestal = new Float32Array(chips.length).fill(NaN)
  chips.forEach((chip, n) => {
    const pedestal = pedestalFor(referenceCell(bmp, chip), params)
    if (pedestal === null) { maskedCells++; return }
    chipPedestal[n] = pedestal
    if (pedestal !== 0) pedestalCells++
    const shift = pedestal + chip.sign0 * code[n % 512]! * params.amplitude
    let realised = 0
    for (let y = chip.y; y < chip.y + chip.cell; y++) for (let x = chip.x; x < chip.x + chip.cell; x++) {
      const i = (y * bmp.width + x) * 4
      for (let ch = 0; ch < 3; ch++) data[i + ch] = bmp.data[i + ch]! + shift
      realised += 0.2126 * (data[i]! - bmp.data[i]!) + 0.7152 * (data[i + 1]! - bmp.data[i + 1]!) + 0.0722 * (data[i + 2]! - bmp.data[i + 2]!)
    }
    chipDelta[n] = realised / chip.cell ** 2
  })
  return { width: bmp.width, height: bmp.height, data, chipDelta, chipPedestal, chipSupport: chips.map(c => ({ ...c })), mark: { variant: pedestalCells ? 'white-pedestal' : 'direct', pedestalCells, maskedCells, params: { ...params } } }
}
