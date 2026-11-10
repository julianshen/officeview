/**
 * Stage-0 measurement-only corpus additions (owner ruling 2026-11-08, spec §2
 * item 14d): the DOCX top-up and the certified image-heavy decks. They live in
 * `corpus-s0/`, NOT `corpus/`, so the shared corpus manifest
 * (`scripts/fetch-corpus.ts`) and its golden suite are unchanged.
 *
 *   bun scripts/stream-s0/fetch-extra.ts          # download + verify
 *   bun scripts/stream-s0/fetch-extra.ts --check  # verify only
 *
 * Source: apache/poi test-data (Apache-2.0), pinned commit below.
 */
import { createHash } from 'node:crypto'
import { existsSync, mkdirSync, readFileSync, writeFileSync } from 'node:fs'
import { join } from 'node:path'

const POI = '93b6a820f4aaefd7b062615f843799947872cb24'
export const S0_EXTRA_DIR = join(import.meta.dir, '..', '..', 'corpus-s0')
export const S0_EXTRA = [
  { dir: 'document', name: 'poi-heading123.docx', upstream: 'heading123.docx', probe: 'heading styles 1-3 (text-heavy)' },
  { dir: 'document', name: 'poi-bib-chernigovka.netdo.ru_download_docs_17459.docx', upstream: 'bib-chernigovka.netdo.ru_download_docs_17459.docx', probe: 'real-world Cyrillic document (text-heavy)' },
  { dir: 'document', name: 'poi-drawing.docx', upstream: 'drawing.docx', probe: 'DrawingML shapes' },
  { dir: 'document', name: 'poi-VariousPictures.docx', upstream: 'VariousPictures.docx', probe: 'several picture formats' },
  { dir: 'document', name: 'poi-issue_51265_1.docx', upstream: 'issue_51265_1.docx', probe: 'embedded pictures' },
  { dir: 'document', name: 'poi-Bug51170.docx', upstream: 'Bug51170.docx', probe: 'picture-bearing Word bug repro' },
  { dir: 'slideshow', name: 'poi-keyframes.pptx', upstream: 'keyframes.pptx', probe: 'full-slide photo background behind text (image-heavy)' },
  { dir: 'slideshow', name: 'poi-aptia.pptx', upstream: 'aptia.pptx', probe: 'photographic title slide (image-heavy)' },
  { dir: 'slideshow', name: 'poi-2411-Performance_Up.pptx', upstream: '2411-Performance_Up.pptx', probe: 'chart screenshot raster (image-heavy)' },
  { dir: 'slideshow', name: 'poi-60810.pptx', upstream: '60810.pptx', probe: 'document screenshots (image-heavy)' },
  { dir: 'slideshow', name: 'poi-at.ecodesign.www_downloads_Vertiefungsvortrag_elektronik.pptx', upstream: 'at.ecodesign.www_downloads_Vertiefungsvortrag_elektronik.pptx', probe: 'product photographs (image-heavy)' },
] as const

if (import.meta.main) {
  const check = process.argv.includes('--check')
  const lockPath = join(S0_EXTRA_DIR, 'corpus.lock.json')
  const lock = JSON.parse(readFileSync(lockPath, 'utf8')) as Record<string, string>
  mkdirSync(S0_EXTRA_DIR, { recursive: true })
  let failed = 0
  for (const e of S0_EXTRA) {
    const at = join(S0_EXTRA_DIR, e.name)
    if (!existsSync(at)) {
      if (check) { console.error(`  MISSING ${e.name}`); failed++; continue }
      const res = await fetch(`https://raw.githubusercontent.com/apache/poi/${POI}/test-data/${e.dir}/${e.upstream}`)
      if (!res.ok) { console.error(`  FAILED ${e.name}: HTTP ${res.status}`); failed++; continue }
      writeFileSync(at, new Uint8Array(await res.arrayBuffer()))
    }
    const digest = createHash('sha256').update(readFileSync(at)).digest('hex')
    if (digest !== lock[e.name]) { console.error(`  HASH MISMATCH ${e.name}`); failed++ }
  }
  console.log(`corpus-s0: ${S0_EXTRA.length - failed} verified, ${failed} failed`)
  process.exit(failed ? 1 : 0)
}
