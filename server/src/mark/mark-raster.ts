/**
 * The per-session Layer-1 encode job body (§9.4 step 4–5): decode the
 * canonical PNG, embed the session's 512-sign code over the page-keyed
 * basis, re-encode losslessly. A unit that cannot carry the mark ships
 * unmarked with its input bytes; the caller reports it (never a silent claim).
 */
import { basis } from './basis'
import { embed, pedestalFor } from './embed'
import { DEFAULT_PARAMS } from './params'
import { decodePng, encodePng } from './png'
import type { Chip, MarkParams } from './types'

export interface MarkRasterInput { png: Uint8Array; w: number; h: number; kBasis: Uint8Array; code: Int8Array; params?: MarkParams }
export type MarkRasterResult =
  | { png: Uint8Array; carriers: number; masked: number; pedestalCells: number }
  | { unmarked: 'capacity' | 'all-masked'; png: Uint8Array }

export function markRaster(input: MarkRasterInput): MarkRasterResult {
  const params = input.params ?? DEFAULT_PARAMS
  let chips: Chip[]
  try { chips = basis(input.kBasis, input.w, input.h, params) } catch (e) {
    if ((e as Error).message === 'Insufficient carrier capacity') return { unmarked: 'capacity', png: input.png }
    throw e
  }
  const d = decodePng(input.png)
  if (d.w !== input.w || d.h !== input.h) throw new Error(`markRaster: PNG is ${d.w}x${d.h}, unit is ${input.w}x${input.h}`)
  const marked = embed({ width: d.w, height: d.h, data: d.rgba }, chips, input.code, params)
  const masked = marked.mark.maskedCells
  if (masked === params.nChips) return { unmarked: 'all-masked', png: input.png }
  return { png: encodePng(d.w, d.h, marked.data), carriers: params.nChips - masked, masked, pedestalCells: marked.mark.pedestalCells }
}

/** Admission estimate for one encode job (§14): decoded RGBA, its marked copy,
 * filtered scanlines and deflate output (each ≈ 4·w·h; deflate bounded by its
 * input for real pages), plus the cloned input PNG and the final PNG. Conservative
 * for rendered pages, not incompressible worst case. */
export const encodeEstimate = (w: number, h: number, pngBytes: number) => w * h * 4 * 4 + 2 * pngBytes

/**
 * Where a show-plan layer sits on its unit's page at rest (piece 4 D7): `x`, `y` are the layer's top-left in
 * page device px at the bucket (may be negative: a layer can hang off the slide), `pageW` × `pageH` the unit
 * page the basis is keyed to. `scale` (default 1) is the layer raster's px per page px: a growth layer (D12) is
 * rasterised above the bucket, so each carrier cell covers `scale` times as many raster px.
 */
export interface LayerCrop { x: number; y: number; pageW: number; pageH: number; scale?: number }
export type MarkLayerResult =
  | { png: Uint8Array; carriers: number; masked: number; pedestalCells: number }
  | { unmarked: 'layer-capacity'; png: Uint8Array }

/**
 * The Layer-1 encode of one show-plan layer (D7): the unit's page-keyed basis, cropped to the layer's rest
 * bounds, with each carrier cell translated into layer-local px. A chip keeps its ORIGINAL index into the
 * 512-sign code (detection scores `code[n % 512]` by the page index, so renumbering a cropped subset would
 * silently break attribution); `embed` cannot be used, as it requires the full chip set inside the bitmap.
 * Cells not wholly inside the layer are skipped; cells that are not fully opaque are masked, as `embed` masks
 * transparent cells. A layer left with zero carriers ships unmarked (`layer-capacity`) with its input bytes.
 */
export function markLayerRaster(input: MarkRasterInput & { crop: LayerCrop }): MarkLayerResult {
  const params = input.params ?? DEFAULT_PARAMS
  const { crop } = input
  const k = crop.scale ?? 1
  if (!(k > 0) || !Number.isFinite(k) || ![crop.x, crop.y].every(Number.isFinite)) throw new Error('markLayerRaster: invalid crop')
  let chips: Chip[]
  try { chips = basis(input.kBasis, crop.pageW, crop.pageH, params) } catch (e) {
    if ((e as Error).message === 'Insufficient carrier capacity') return { unmarked: 'layer-capacity', png: input.png }
    throw e
  }
  if (input.code.length !== 512 || input.code.some((x) => x !== -1 && x !== 1)) throw new Error('markLayerRaster: expected 512 signed symbols')
  const d = decodePng(input.png)
  if (d.w !== input.w || d.h !== input.h) throw new Error(`markLayerRaster: PNG is ${d.w}x${d.h}, layer is ${input.w}x${input.h}`)
  const src = d.rgba, data = src.slice()
  let carriers = 0, masked = 0, pedestalCells = 0
  for (let n = 0; n < chips.length; n++) {
    const chip = chips[n]!
    const u0 = Math.round((chip.x - crop.x) * k), v0 = Math.round((chip.y - crop.y) * k)
    const u1 = Math.round((chip.x + chip.cell - crop.x) * k), v1 = Math.round((chip.y + chip.cell - crop.y) * k)
    if (u0 < 0 || v0 < 0 || u1 > d.w || v1 > d.h || u1 <= u0 || v1 <= v0) continue
    const cw = u1 - u0, ch = v1 - v0
    const cell = new Uint8ClampedArray(cw * ch * 4)
    for (let y = 0; y < ch; y++) cell.set(src.subarray(((v0 + y) * d.w + u0) * 4, ((v0 + y) * d.w + u1) * 4), y * cw * 4)
    const pedestal = pedestalFor({ width: cw, height: ch, data: cell }, params)
    if (pedestal === null) { masked++; continue }
    if (pedestal !== 0) pedestalCells++
    const shift = pedestal + chip.sign0 * input.code[n % 512]! * params.amplitude
    for (let y = v0; y < v1; y++) for (let x = u0; x < u1; x++) {
      const i = (y * d.w + x) * 4
      for (let c = 0; c < 3; c++) data[i + c] = src[i + c]! + shift
    }
    carriers++
  }
  if (carriers === 0) return { unmarked: 'layer-capacity', png: input.png }
  return { png: encodePng(d.w, d.h, data), carriers, masked, pedestalCells }
}
