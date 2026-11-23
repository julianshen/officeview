/**
 * Stage 4a (S4 §6.5): the tile job, its bounds, its palettes and its Layer-1 basis label. Light on purpose (the wire
 * constants and the theme tokens only, the rest as types): the pool's thread imports it for `TILE_BYTES_ESTIMATE`, and
 * the paint itself runs in the worker (`DocumentHost.renderTile`).
 */
import { TILE_DPRS, TILE_PALETTES, TILE_PX, TILE_ZOOM_STEPS, type TilePalette } from '../../../src/stream/sheet-wire'
import { THEME_DEFAULTS, type CanvasToken } from '../../../src/styles/tokens'
import type { MarkSpec } from '../mark/visible'
import type { SheetClock } from './clock'
import type { WireSizes } from './view-state'

/** One tile of sheet `sheetIndex` (an index into the visible sheets, R12) at `(TILE_ZOOM_STEPS[zoomStep], dpr)`: the
 * 512² device-px square at `(col·512, row·512)` of the sheet's global device grid (`tileRangeFor`), painted at the
 * session's `sizes` (Stage 4b, canonical `WireSizes`). A job without `sizes` is the 4a identity job: its `viewGen`
 * must be 0. With `sizes` (even `[]`, a reset view) `viewGen` is any u32: the paint reads the sizes, never the number. */
export interface TileJob { sheetIndex: number; viewGen: number; zoomStep: number; dpr: 1 | 2; palette: TilePalette; col: number; row: number; mark: MarkSpec; clock: SheetClock; sizes?: WireSizes }
/** What a tile job returns: the canonical (pre-forensic) PNG, and the palette it was painted under (F24: audited). */
export interface RenderedTile { png: Uint8Array; w: typeof TILE_PX; h: typeof TILE_PX; palette: TilePalette }

/** Admission charge of one tile: its RGBA surface and the encode buffers. */
export const TILE_BYTES_ESTIMATE = TILE_PX * TILE_PX * 4 * 2

/** R12: the two fixed server palettes, from the default theme tokens. Only `pending`/`pending-text` reach a tile
 * (`ViewportPaintOptions.colors`); host-customised canvas tokens never do. */
export const TILE_PALETTE_COLORS: Readonly<Record<TilePalette, Readonly<Pick<Record<CanvasToken, string>, 'pending' | 'pending-text'>>>> = Object.freeze({
  light: Object.freeze({ pending: THEME_DEFAULTS.light.pending, 'pending-text': THEME_DEFAULTS.light['pending-text'] }),
  dark: Object.freeze({ pending: THEME_DEFAULTS.dark.pending, 'pending-text': THEME_DEFAULTS.dark['pending-text'] }),
})

/**
 * F8: the number of tiles along an axis of `extentCss` logical px at `(zoom, dpr)`. A valid index is `< tileLimit`,
 * which is `tileRangeFor`'s last index + 1 (`max(0, ceil(extent·s/512) − 1) + 1`), so an empty axis still has tile 0.
 * Tasks 7 (the worker) and 9 (the driver) both bound tiles with it.
 */
export function tileLimit(extentCss: number, zoom: number, dpr: number): number {
  return Math.max(0, Math.ceil((extentCss * zoom * dpr) / TILE_PX) - 1) + 1
}

/** F26: the per-tile Layer-1 basis label, `sheet:<id>@<zoom>x<dpr>:<c>,<r>` with `zoom` the `SHEET_ZOOM_STEPS` value
 * printed by `String()` (e.g. `0.33`), as the Task 1 spike keyed it. */
export const tileBasisLabel = (sheetId: string, zoomStep: number, dpr: number, col: number, row: number) => `sheet:${sheetId}@${String(TILE_ZOOM_STEPS[zoomStep])}x${dpr}:${col},${row}`

const nonNegInt = (v: unknown) => Number.isSafeInteger(v) && (v as number) >= 0
const validMark = (m: unknown) => {
  if (!m || typeof m !== 'object') return false
  const k = (m as { kind?: unknown }).kind
  return k === 'fixed' || k === 'off' || (k === 'viewer' && typeof (m as { text?: unknown }).text === 'string')
}

/** At most this many session sizes per axis (`MAX_SESSION_SIZES`, duplicated: this module stays light). */
const SIZES_PER_AXIS = 4096
const AXIS_COUNT = { col: 16_384, row: 1_048_576 } as const
/** `sizes` is canonical `WireSizes`: axis `'col'` before `'row'`, each strictly ascending, in range, a finite size > 0,
 * at most 4,096 per axis. */
function sizesProblem(sizes: unknown): boolean {
  if (!Array.isArray(sizes) || sizes.length > 2 * SIZES_PER_AXIS) return true
  let prev: { axis: string; index: number } | undefined, perAxis = 0
  for (const e of sizes as unknown[]) {
    if (!e || typeof e !== 'object') return true
    const { axis, index, px } = e as Record<string, unknown>
    if ((axis !== 'col' && axis !== 'row') || !nonNegInt(index) || (index as number) >= AXIS_COUNT[axis] || typeof px !== 'number' || !Number.isFinite(px) || px <= 0) return true
    if (prev && (prev.axis === axis ? prev.index >= (index as number) : prev.axis === 'row')) return true
    perAxis = prev?.axis === axis ? perAxis + 1 : 1
    if (perAxis > SIZES_PER_AXIS) return true
    prev = { axis, index: index as number }
  }
  return false
}

/** Why `job` is malformed (every check that needs no sheet), or undefined. A job without `sizes` has `viewGen` 0 (4a);
 * with them, any u32. The sheet index and the tile bounds are checked against the snapshot. */
export function tileJobProblem(job: unknown): string | undefined {
  if (!job || typeof job !== 'object') return 'job'
  const j = job as Partial<TileJob>
  if (!nonNegInt(j.sheetIndex)) return 'sheetIndex'
  if (j.sizes === undefined ? j.viewGen !== 0 : !nonNegInt(j.viewGen) || j.viewGen! > 0xffffffff) return 'viewGen'
  if (j.sizes !== undefined && sizesProblem(j.sizes)) return 'sizes'
  if (!nonNegInt(j.zoomStep) || j.zoomStep! >= TILE_ZOOM_STEPS.length) return 'zoomStep'
  if (!(TILE_DPRS as readonly unknown[]).includes(j.dpr)) return 'dpr'
  if (!(TILE_PALETTES as readonly unknown[]).includes(j.palette)) return 'palette'
  if (!nonNegInt(j.col) || !nonNegInt(j.row)) return 'col/row'
  if (!validMark(j.mark)) return 'mark'
  if (!j.clock || typeof j.clock !== 'object' || typeof j.clock.day !== 'string' || typeof j.clock.timeZone !== 'string') return 'clock'
  return undefined
}
