/**
 * Piece 5 (spec §4.4; plan D4, D12, D16, D20; Task 10): the remote `MediaPlayerDeps.open`. A streamed show's media
 * record (D16: `rId` is the session's mediaId, `partPath` is ignored, every record is audio) is fetched whole over the
 * session (`StreamSession.mediaRequest`), reassembled into a Blob and handed to the unchanged media player as a
 * `MediaLease` with `openMediaLease`'s span, retain, leave and release semantics.
 *
 * The workspace charge follows `openMediaLease` (slideshow/media/lease.ts): `start.bytes` from `media-start` (the
 * chunks as they arrive), then the peak `2 × bytes` while the chunks and the Blob coexist, then `bytes` for the Blob
 * alone. No room is the player's `resource-limit`, its signal to drop carried media first and try again.
 *
 * Lazy remote path only: everything from the local lease is imported as a type (its module reads package parts).
 * It rejects only on abort; every other outcome is a `{refused}`.
 */
import type { MediaLease, MediaLeaseResult } from '../../slideshow/media/lease'
import type { Lease, MediaRecord, SlideShowWorkspace } from '../../slideshow/types'
import type { StreamSession } from './session'

/** The one type every streamed media record plays as (D16, D22: audio-only M4A, AAC-LC). */
export const REMOTE_MEDIA_TYPE = 'audio/mp4'
/** D20: what the browser must accept before any byte is requested (`''` refuses `media-codec`). */
export const REMOTE_MEDIA_CODEC = 'audio/mp4; codecs="mp4a.40.2"'

interface WindowUrl { URL?: { createObjectURL?(blob: Blob): string; revokeObjectURL?(url: string): void } }

const aborted = (signal: AbortSignal) => signal.reason ?? Object.assign(new Error('The media request was aborted.'), { name: 'AbortError' })

/** `unit`: the slide the record is played on, sent with the request (look-ahead requests especially). */
export async function openRemoteMediaLease(
  session: Pick<StreamSession, 'mediaRequest'>,
  rec: MediaRecord,
  ws: SlideShowWorkspace,
  win: Window,
  signal: AbortSignal,
  unit?: number,
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

  let lease: Lease | undefined
  let got: Awaited<ReturnType<StreamSession['mediaRequest']>>
  try {
    got = await session.mediaRequest(rec.rId, signal, (start) => {
      lease = ws.lease(start.bytes, `media:${rec.rId}`)
      return lease !== undefined
    }, unit)
  } catch (e) {
    lease?.release()
    if (signal.aborted) throw aborted(signal)
    // A session that ended (or closed) under the request: the media is unavailable, never a rejection.
    void e
    return { refused: 'media-unavailable' }
  }
  if ('refused' in got) { lease?.release(); return { refused: got.refused } }
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
