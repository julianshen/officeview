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
export { createRemoteDoc, bucketFor, type RemoteOfficeDoc, type RemoteRange, type RemoteSearchHit, type SessionLike } from './remote/remote-doc'
export { StreamSession, type StreamClient, type StreamClientFactory, type StreamNotice, type TextMessage, type TextRequest, type UnitPayload } from './remote/session'
export { StreamSessionError, type StreamErrorCode } from './remote/errors'
export { resolveFaces } from './remote/faces'
