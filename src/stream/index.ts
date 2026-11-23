/**
 * The lazy stream runtime (spec §5): reached from `officeview/core` only via
 * `loadStreamRuntime()`. Stage 1b/1d add the frame codec, replayer and the
 * remote document provider here.
 */
export const STREAM_RUNTIME_VERSION = 1
export * from './manifest'
export * from './raster'
export * from './asset-wire'
export { openRemoteDocument, type OpenRemoteOptions } from './remote/open'
export { createRemoteDoc, bucketFor, type RemoteOfficeDoc, type RemoteRange, type RemoteSearchHit, type RemoteSheets, type SessionLike } from './remote/remote-doc'
export { StreamSession, type StreamClient, type StreamClientFactory, type StreamNotice, type TextMessage, type TextRequest, type UnitPayload } from './remote/session'
/** Stage 4a (S4 §6.11): the tile protocol client. `streamSheetGeometry` is not here: it would pull `src/xlsx/view` engine modules into this runtime (F5). */
export {
  decodeSheetLayout, decodeTileBody, SHEET_TILES_FEATURE, TILE_PX, TILE_BATCH_MAX, TILE_DPRS, TILE_PALETTES, TILE_ZOOM_STEPS,
  HELLO_FEATURES_MAX, HELLO_FEATURE_CHARS, MAX_SHEETS, SHEET_NAME_CHARS, SHEET_MAX_ROWS, SHEET_MAX_COLS, LAYOUT_LIMITS, LAYOUT_PART_BYTES,
  LayoutInvalid, LayoutTooLarge, type StreamSheetLayout, type TileBody, type TilePalette,
} from './sheet-wire'
export { TILE_RETRY_FALLBACK_MS, TILE_RETRY_MAX_MS, type DecodedTile, type TilesRequest, type TilesResult } from './remote/sheet-channel'
export { StreamSessionError, type StreamErrorCode } from './remote/errors'
export { resolveFaces } from './remote/faces'
