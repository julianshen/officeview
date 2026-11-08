import { checkBitmap, type Bitmap, type Chip, type MarkParams } from './types'
import { checkParams } from './params'
export interface MarkedBitmap extends Bitmap {
  mark: { variant: 'white-pedestal' | 'direct'; pedestalCells: number; maskedCells: number; params: MarkParams }
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
  chips.forEach((chip, n) => {
    let mean = 0, opaque = true
    for (let y = chip.y; y < chip.y + chip.cell; y++) for (let x = chip.x; x < chip.x + chip.cell; x++) {
      const i = (y * bmp.width + x) * 4
      mean += (bmp.data[i]! + bmp.data[i + 1]! + bmp.data[i + 2]!) / 3
      if (bmp.data[i + 3] !== 255) opaque = false
    }
    if (!opaque) { maskedCells++; return }
    mean /= chip.cell ** 2
    const pedestal = mean >= 254 && params.pedestal > 0
    if (pedestal) pedestalCells++
    const shift = (pedestal ? 255 - params.pedestal - mean : 0) + chip.sign0 * code[n % 512]! * params.amplitude
    for (let y = chip.y; y < chip.y + chip.cell; y++) for (let x = chip.x; x < chip.x + chip.cell; x++) {
      const i = (y * bmp.width + x) * 4
      for (let ch = 0; ch < 3; ch++) data[i + ch] = bmp.data[i + ch]! + shift
    }
  })
  return { width: bmp.width, height: bmp.height, data, mark: { variant: pedestalCells ? 'white-pedestal' : 'direct', pedestalCells, maskedCells, params: { ...params } } }
}
