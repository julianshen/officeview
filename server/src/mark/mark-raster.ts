/**
 * The per-session Layer-1 encode job body (§9.4 step 4–5): decode the
 * canonical PNG, embed the session's 512-sign code over the page-keyed
 * basis, re-encode losslessly. A unit that cannot carry the mark ships
 * unmarked with its input bytes; the caller reports it (never a silent claim).
 */
import { basis } from './basis'
import { embed } from './embed'
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
