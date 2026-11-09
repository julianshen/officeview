import type { MarkParams } from './types'
export const DEFAULT_PARAMS: Readonly<MarkParams> = Object.freeze({ nChips: 512, cellPx: 4, tilePx: 64, amplitude: 2, pedestal: 2 })
export function* sweepParams(): Generator<MarkParams> {
  for (const amplitude of [1, 2, 3, 4]) for (const cellPx of [2, 4, 6]) for (const pedestal of [0, 2]) yield { ...DEFAULT_PARAMS, amplitude, cellPx, pedestal }
}
export function checkParams(p: MarkParams): void {
  if (!Number.isSafeInteger(p.nChips) || p.nChips < 512 || !Number.isSafeInteger(p.cellPx) || p.cellPx < 2 || !Number.isSafeInteger(p.tilePx) || p.tilePx < p.cellPx || !Number.isFinite(p.amplitude) || p.amplitude < 0 || p.amplitude > 255 || !Number.isFinite(p.pedestal) || p.pedestal < 0 || p.pedestal > 255) throw new Error('Invalid mark parameters')
}
/** §13.2 Layer-1 codec identity recorded in audit rows: extended BCH(511,40) → 512 signs. */
export const CODEC_ID = 'l1-bch512x40@1' as const
