/**
 * Stage 4a: the `sheet-tiles@1` wire constants (S4 §6, R12), shared by the stream server and client. React-free, and
 * it imports nothing from `src/xlsx/**`, so the stream runtime's module set and the eager gates do not move:
 * `TILE_ZOOM_STEPS` is a literal, and a test pins it to `SHEET_ZOOM_STEPS`.
 */

/** The HELLO capability that selects tiles mode for an XLSX session. */
export const SHEET_TILES_FEATURE = 'sheet-tiles@1'
/** A tile is 512 × 512 device px. */
export const TILE_PX = 512
/** At most 64 tiles per batch. */
export const TILE_BATCH_MAX = 64
export const TILE_DPRS = [1, 2] as const
export const TILE_PALETTES = ['light', 'dark'] as const
/** Equal to `SHEET_ZOOM_STEPS` (src/xlsx/view/sheetZoom.ts); a tile's zoom travels as an index into it. */
export const TILE_ZOOM_STEPS = [0.25, 0.33, 0.5, 0.67, 0.75, 1, 1.25, 1.5, 2, 3, 4] as const
export type TilePalette = typeof TILE_PALETTES[number]
/** HELLO `features`: at most 16 entries, each at most 32 characters (R12). */
export const HELLO_FEATURES_MAX = 16, HELLO_FEATURE_CHARS = 32
/** A tiles manifest lists at most 256 sheets, each name at most 31 characters (Excel's own bound). A workbook past
 * either bound is served in units mode. */
export const MAX_SHEETS = 256, SHEET_NAME_CHARS = 31
