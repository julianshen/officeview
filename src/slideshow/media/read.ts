/**
 * Bounded media reads (slideshow spec 6.1, 7): a media part is read with the
 * uncached bounded reader - never the caching `OfficePackage.bytes()` - and
 * counted inflated chunk by chunk, so a forged declared size is caught at the
 * read cap and a tight aggregate budget refuses mid-inflate, before the part
 * is buffered.
 */
import { readPartBounded } from '../../pptx/appearance/read'
import type { MediaRecord } from '../types'
import type { PptxPackageHandle } from '../../pptx/types'

const MiB = 1024 * 1024
/** One media part may inflate to at most 512 MiB (spec 6.1). */
export const MEDIA_PART_LIMIT = 512 * MiB

/** What one read produced: the inflated chunks and their total (the counter the budget is driven by). */
export interface ReadMediaResult {
  chunks: Uint8Array[]
  inflated: number
}

/**
 * Reads one media part under the cap. Every inflated chunk is first admitted
 * through `adopt` - false aborts the inflate before the chunk is kept, so the
 * refusal happens at the crossing chunk - and the abort signal is honoured at
 * every chunk boundary. `undefined` means the part is absent from the package.
 */
export async function readMediaPart(
  handle: PptxPackageHandle,
  rec: MediaRecord,
  adopt: (chunkBytes: number) => boolean,
  signal?: AbortSignal,
): Promise<ReadMediaResult | undefined> {
  let inflated = 0
  const chunks = await readPartBounded(handle.pkg, rec.partPath, MEDIA_PART_LIMIT, {
    signal,
    output: 'chunks',
    admit: n => {
      if (!adopt(n)) return false
      inflated += n
      return true
    },
  })
  return chunks && { chunks, inflated }
}
