/**
 * Piece 5 (spec §4.4; plan D4, D12, D16, D20; Task 10): the remote `MediaPlayerDeps.open`. A streamed show's media
 * record (D16: `rId` is the session's mediaId, `partPath` is ignored, every record is audio) is fetched whole over the
 * session (`StreamSession.mediaRequest`), reassembled into a Blob and handed to the unchanged media player as a
 * `MediaLease` with `openMediaLease`'s span, retain, leave and release semantics.
 *
 * The workspace charge (ruling 4): the 32 MiB per-file cap is reserved before the request is sent (no room: it waits
 * unsent, after `evict`), shrunk to `start.bytes` at `media-start` (the chunks as they arrive), then, as
 * `openMediaLease` (slideshow/media/lease.ts), the peak `2 × bytes` while the chunks and the Blob coexist, then
 * `bytes` for the Blob alone. No room at that peak is the player's `resource-limit` (a defensive path).
 *
 * Lazy remote path only: everything from the local lease is imported as a type (its module reads package parts).
 * It rejects only on abort; every other outcome is a `{refused}`.
 */
import type { MediaLease, MediaLeaseResult } from '../../slideshow/media/lease'
import type { Lease, MediaRecord, SlideShowWorkspace } from '../../slideshow/types'
import type { MediaRoom, StreamSession } from './session'

/** The one type every streamed media record plays as (D16, D22: audio-only M4A, AAC-LC). */
export const REMOTE_MEDIA_TYPE = 'audio/mp4'
/** D20: what the browser must accept before any byte is requested (`''` refuses `media-codec`). */
export const REMOTE_MEDIA_CODEC = 'audio/mp4; codecs="mp4a.40.2"'

interface WindowUrl { URL?: { createObjectURL?(blob: Blob): string; revokeObjectURL?(url: string): void } }

const aborted = (signal: AbortSignal) => signal.reason ?? Object.assign(new Error('The media request was aborted.'), { name: 'AbortError' })

/** Options of `openRemoteMediaLease`. */
export interface RemoteMediaLeaseOptions {
  /** The slide the record is played on, sent with the request (look-ahead requests especially). */
  unit?: number
  /** Ruling 4: frees room by evicting delivered entries no one is using (the caller's); true when it freed some. Called
   * once per reservation attempt that finds no room, before the request waits. */
  evict?: () => boolean
}

export async function openRemoteMediaLease(
  session: Pick<StreamSession, 'mediaRequest'>,
  rec: MediaRecord,
  ws: SlideShowWorkspace,
  win: Window,
  signal: AbortSignal,
  o: RemoteMediaLeaseOptions = {},
): Promise<MediaLeaseResult> {
  if (signal.aborted) throw aborted(signal)
  // A realm without Blob-URL support (a torn-down popup) is refused before anything is requested.
  const url = (win as unknown as WindowUrl | undefined)?.URL
  if (!url || typeof url.createObjectURL !== 'function' || typeof url.revokeObjectURL !== 'function') return { refused: 'media-unavailable' }
  const create = url.createObjectURL.bind(url), revoke = url.revokeObjectURL.bind(url)
  // D20: a browser without AAC is told so before a request spends a pacing slot.
  let playable = ''
  try { playable = win.document.createElement('audio').canPlayType(REMOTE_MEDIA_CODEC) } catch { playable = '' }
  if (playable === '') return { refused: 'media-codec' }

  // Ruling 4: the session reserves the per-file cap here before it sends, shrinks it to the file at media-start, and
  // returns it on any settlement without a delivery; a delivery leaves it to this lease.
  let lease: Lease | undefined
  const tag = `media:${rec.rId}`
  // `evict` runs once, at the first attempt that finds no room, before the request waits (review round 1).
  let evicted = false
  const room: MediaRoom = {
    reserve(bytes) {
      if (lease) return true
      lease = ws.lease(bytes, tag)
      if (!lease && !evicted) { evicted = true; if (o.evict?.()) lease = ws.lease(bytes, tag) }
      return lease !== undefined
    },
    ...(ws.onFree ? { onFree: (wake: () => void) => ws.onFree!(wake) } : {}),
    fit(start) { if (lease && lease.bytes > start.bytes) lease.shrink(lease.bytes - start.bytes) },
    release() { lease?.release(); lease = undefined },
  }
  let got: Awaited<ReturnType<StreamSession['mediaRequest']>>
  try {
    got = await session.mediaRequest(rec.rId, signal, room, o.unit)
  } catch {
    room.release()
    if (signal.aborted) throw aborted(signal)
    // A session that ended (or closed) under the request: the media is unavailable, never a rejection.
    return { refused: 'media-unavailable' }
  }
  if ('refused' in got) { room.release(); return { refused: got.refused } }
  const held = lease!
  const { chunks, start } = got
  if (signal.aborted) { held.release(); chunks.length = 0; throw aborted(signal) }
  // Peak: the chunks plus the Blob they are about to become are both on the ledger while `new Blob` runs.
  if (!held.grow(start.bytes)) { held.release(); chunks.length = 0; return { refused: 'resource-limit' } }
  let blob: Blob
  try {
    blob = new Blob(chunks as unknown as BlobPart[], { type: REMOTE_MEDIA_TYPE })
  } catch {
    held.release(); chunks.length = 0
    return { refused: 'resource-limit' }
  }
  held.shrink(start.bytes)
  chunks.length = 0
  let objectUrl: string
  try {
    objectUrl = create(blob)
  } catch {
    held.release()
    return { refused: 'media-unavailable' }
  }
  let spans = 0
  let revoked = false
  const out: MediaLease = {
    get url() { return objectUrl },
    get bytes() { return blob.size },
    get spans() { return spans },
    retain(slides) {
      if (slides === Infinity) { spans = Infinity; return }
      if (!Number.isSafeInteger(slides) || slides < 0) throw new RangeError(`media lease: retain wants a whole number of slides or Infinity, got ${String(slides)}`)
      spans = Math.max(spans, slides)
    },
    leave() {
      if (revoked || spans === Infinity) return
      if (spans > 0) spans--
      if (spans === 0) out.release()
    },
    release() {
      if (revoked) return
      revoked = true
      // The charge is back before the revocation is attempted, so a realm that throws cannot strand the bytes.
      held.release()
      try { revoke(objectUrl) } catch { /* the realm is gone */ }
    },
  }
  return out
}
