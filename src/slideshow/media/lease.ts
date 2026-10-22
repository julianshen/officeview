/**
 * Media leases (slideshow spec 6.1, 7): playing one media part needs its bytes
 * in a Blob URL. The part is read bounded (read.ts), the chunks and the Blob
 * share one workspace charge during construction, and the finished lease holds
 * only the Blob's inflated length on the ledger. The lease owns the Blob URL;
 * `leave()` (one slide exit) revokes it as soon as the retention span ends,
 * and `release()` (show exit, teardown) revokes it now regardless of the span.
 * Both refund the charge exactly once, whether or not the URL revocation
 * itself succeeds.
 */
import { isPackageError, throwIfAborted } from '../../core/package-error'
import type { MediaRecord, SlideShowWorkspace } from '../types'
import type { PptxPackageHandle } from '../../pptx/types'
import { readMediaPart, MEDIA_PART_LIMIT } from './read'

/** A media part's declared cap, refuted unread at playback time too (spec 6.1). */
export { MEDIA_PART_LIMIT }

/** The Blob-URL statics the lease uses, taken from the caller's realm window. */
interface WindowUrl {
  URL: { createObjectURL(blob: Blob): string; revokeObjectURL(url: string): void }
}

export interface MediaLease {
  /** The Blob URL, valid until the span ends (`leave`) or `release`. */
  url: string
  /** The Blob's inflated length; exactly what this leaves charged to the workspace. */
  bytes: number
  /** How many slides the media is retained across (cross-slide sound, spec 7); 0 stops at its slide, Infinity until stopped. */
  readonly spans: number
  /** Extends the media's life for `slides` more slide leaves (`Infinity` plays until it is stopped); a larger span wins. */
  retain(slides: number): void
  /** One slide exit. Retained media loses one slide of span and revokes at zero. */
  leave(): void
  /** Revokes the URL exactly once now, regardless of the span, and returns the charge. */
  release(): void
}

export type MediaLeaseResult = MediaLease | { refused: string }

export async function openMediaLease(
  handle: PptxPackageHandle,
  rec: MediaRecord,
  ws: SlideShowWorkspace,
  win: Window,
  signal?: AbortSignal,
): Promise<MediaLeaseResult> {
  throwIfAborted(signal)
  // A realm without Blob-URL support (a torn-down popup) is refused before a
  // single chunk is inflated, let alone charged.
  const url = (win as unknown as WindowUrl | undefined)?.URL
  if (!url || typeof url.createObjectURL !== 'function' || typeof url.revokeObjectURL !== 'function') return { refused: 'media-unavailable' }
  // The parser gate, repeated at playback time: a part that declares more than
  // the cap is refused without a single inflated chunk.
  if (rec.declaredBytes > MEDIA_PART_LIMIT) return { refused: 'media-limit' }
  const lease = ws.lease(0, `media:${rec.partPath}`)
  if (!lease) return { refused: 'resource-limit' }
  try {
    const read = await readMediaPart(handle, rec, n => lease.grow(n), signal)
    if (!read) { lease.release(); return { refused: 'media-missing' } }
    // A real empty part gets no live URL to play.
    if (read.inflated === 0) { lease.release(); return { refused: 'media-empty' } }
    // Peak: the chunk list plus the Blob it is about to become are both on the
    // ledger while `new Blob(chunks)` runs; the chunk charge leaves right after.
    if (!lease.grow(read.inflated)) { lease.release(); return { refused: 'resource-limit' } }
    // A hostile content type cannot make the Blob a renderable document type.
    const type = /^(?:audio|video)\//.test(rec.contentType) ? rec.contentType : ''
    let blob: Blob
    try {
      blob = new Blob(read.chunks as unknown as BlobPart[], { type })
    } catch {
      lease.release()
      return { refused: 'resource-limit' }
    }
    lease.shrink(read.inflated)
    read.chunks.length = 0
    let objectUrl: string
    try {
      objectUrl = url.createObjectURL(blob)
    } catch {
      lease.release()
      return { refused: 'media-unavailable' }
    }
    let spans = 0
    let revoked = false
    const out: MediaLease = {
      url: objectUrl,
      bytes: blob.size,
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
        lease.release()
        // The charge is back before the revocation is attempted, so a realm
        // that throws cannot strand the bytes on the ledger. A failed revoke
        // is swallowed: the life of the URL ended with the show either way.
        try { url.revokeObjectURL(objectUrl) } catch { /* the realm is gone */ }
      },
    }
    return out
  } catch (e) {
    lease.release()
    throwIfAborted(signal)
    if ((e as { mediaLimit?: unknown })?.mediaLimit === true) return { refused: 'media-limit' }
    if (isPackageError(e)) {
      if (e.code === 'resource-limit') return { refused: 'resource-limit' }
      return { refused: 'media-refused' }
    }
    throw e
  }
}
