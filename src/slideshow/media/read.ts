/**
 * Bounded media reads (slideshow spec 6.1, 7): a media part is read with the
 * uncached bounded reader - never the caching `OfficePackage.bytes()` - and
 * counted inflated chunk by chunk, so a forged declared size is caught at the
 * read cap and a tight aggregate budget refuses mid-inflate, at the crossing
 * chunk, before the part is buffered.
 */
import { readPartBounded } from '../../pptx/appearance/read'
import type { MediaRecord } from '../types'
import type { PptxPackageHandle } from '../../pptx/types'

const MiB = 1024 * 1024
/** One media part may inflate to at most 512 MiB (spec 6.1). */
export const MEDIA_PART_LIMIT = 512 * MiB

/** Marker for a real inflate crossing the cap, `media-limit` whatever the declaration claims. */
class MediaLimitOverflow extends Error {
  readonly mediaLimit = true
}

/** What one read produced: the inflated chunks and their total (the counter the budget is driven by). */
export interface ReadMediaResult {
  chunks: Uint8Array[]
  inflated: number
}

/**
 * Reads one media part under the cap. The inflated counter carries the
 * cap itself: a part whose real length crosses 512 MiB is refused with the
 * `media-limit` label whatever its declaration claims. Every chunk is first
 * admitted through `adopt` - false aborts the inflate before the chunk is
 * kept, so the refusal happens at the crossing chunk - and the abort signal
 * is honoured at every chunk boundary. `undefined` means the part is absent
 * from the package.
 */
export async function readMediaPart(
  handle: PptxPackageHandle,
  rec: MediaRecord,
  adopt: (chunkBytes: number) => boolean,
  signal?: AbortSignal,
): Promise<ReadMediaResult | undefined> {
  let inflated = 0
  const chunks = await readPartBounded(handle.pkg, rec.partPath, Infinity, {
    signal,
    output: 'chunks',
    admit: n => {
      // The cap owns this label: a forged low declaration cannot buy a
      // `resource-limit` renaming of the same condition.
      if (inflated + n > MEDIA_PART_LIMIT) throw new MediaLimitOverflow(`the media part inflates past ${MEDIA_PART_LIMIT} bytes: ${rec.partPath}`)
      if (!adopt(n)) return false
      inflated += n
      return true
    },
  })
  return chunks && { chunks, inflated }
}
