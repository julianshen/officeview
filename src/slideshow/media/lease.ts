/**
 * Media leases (slideshow spec 6.1, 7): playing one media part needs its bytes
 * in a Blob URL. The part is read bounded (read.ts), the chunks and the Blob
 * share one workspace charge during construction, and the finished lease holds
 * only the Blob's inflated length on the ledger. The lease owns the Blob URL
 * and revokes it exactly once - when the slide is left and the caller releases,
 * unless the media was retained across slides through `retain`.
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
  /** The Blob URL, valid until `release`. */
  url: string
  /** The Blob's inflated length; exactly what this leaves charged to the workspace. */
  bytes: number
  /** How many slides the media is retained across (cross-slide sound, spec 7); 0 stops at its slide. */
  spans: number
  /** Records that the media should stay alive for `slides` more slides; a larger span wins. */
  retain(slides: number): void
  /** Revokes the URL exactly once and returns the charge, whatever the outcome. */
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
  // The parser gate, repeated at playback time: a part that declares more than
  // the cap is refused without a single inflated chunk.
  if (rec.declaredBytes > MEDIA_PART_LIMIT) return { refused: 'media-limit' }
  const url = (win as unknown as WindowUrl).URL
  const lease = ws.lease(0, `media:${rec.partPath}`)
  if (!lease) return { refused: 'resource-limit' }
  try {
    const read = await readMediaPart(handle, rec, n => lease.grow(n), signal)
    if (!read) { lease.release(); return { refused: 'media-missing' } }
    // Peak: the chunk list plus the Blob it is about to become are both on the
    // ledger while `new Blob(chunks)` runs; the chunk charge leaves right after.
    if (!lease.grow(read.inflated)) { lease.release(); return { refused: 'resource-limit' } }
    const blob = new Blob(read.chunks as unknown as BlobPart[], { type: rec.contentType })
    lease.shrink(read.inflated)
    const objectUrl = url.createObjectURL(blob)
    let revoked = false
    const out: MediaLease = {
      url: objectUrl,
      bytes: blob.size,
      spans: 0,
      retain(slides) { if (Number.isFinite(slides) && slides > this.spans) this.spans = slides },
      release() {
        if (revoked) return
        revoked = true
        url.revokeObjectURL(objectUrl)
        lease.release()
      },
    }
    return out
  } catch (e) {
    lease.release()
    throwIfAborted(signal)
    if (isPackageError(e) && e.code === 'resource-limit') return { refused: 'resource-limit' }
    throw e
  }
}
