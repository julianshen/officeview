/**
 * Slide appearance goldens (slideshow S0-7, spec 11 D2): real LibreOffice decks whose backgrounds and
 * master/layout artwork go through the lazy enrichment in getPaintables. The general corpus suite excludes
 * the `lo-` decks; their goldens live here, in tests/goldens/appearance/<file>.<slide>.png.
 *
 * Mac-only, like tests/golden.test.ts (CI excludes tests/golden*.test.ts). A missing golden is recorded on
 * first run unless OFFICEVIEW_STRICT_GOLDEN=1. The decks are fetched, not vendored
 * (`bun scripts/fetch-corpus.ts`): the suite skips when all of them are absent and fails when only some are.
 */
import { describe, test, expect } from 'vitest'
import { existsSync, mkdirSync, readFileSync } from 'node:fs'
import { join, dirname } from 'node:path'
import { fileURLToPath } from 'node:url'
import { loadOfficeFile } from '../src/components/OfficeFile'
import { getPaintables } from '../src/render/paint'
import { diffBitmaps, loadPng, renderPaintables, savePng } from '../src/test/pixel-diff'
import type { PptxDocument } from '../src/pptx/types'

const HERE = dirname(fileURLToPath(import.meta.url))
const GOLDEN_DIR = join(HERE, 'goldens', 'appearance')
const CORPUS_DIR = join(HERE, '..', 'corpus')
const THRESHOLD = 8 // per-channel AA forgiveness, as tests/golden.test.ts
const MAX_RATIO = 0.005

/** Background and master-art decks (docs/validation/slideshow-corpus.md). */
const DECKS = [
  'lo-tdf105739.pptx', // slide gradient background
  'lo-tdf146223.pptx', // picture background, master picture artwork
  'lo-tdf123684.pptx', // bgRef 1003 theme picture background, six master artwork nodes
  'lo-ooxtheme.pptx', // master artwork group (its layout hides it with showMasterSp="0")
  'lo-tdf93868.pptx', // inherited master gradient background, master artwork
]
const present = DECKS.filter(file => existsSync(join(CORPUS_DIR, file)))
const strict = process.env.OFFICEVIEW_STRICT_GOLDEN === '1'

describe('slide appearance goldens', () => {
  if (present.length === 0) {
    test.skip('slideshow corpus not fetched — run: bun scripts/fetch-corpus.ts', () => {})
    return
  }
  test('the background and master-art decks are all present', () => {
    expect(present).toEqual(DECKS)
  })
  for (const file of present) {
    test(`${file} matches its appearance golden`, async () => {
      const doc = await loadOfficeFile(readFileSync(join(CORPUS_DIR, file))) as PptxDocument
      const paintables = await getPaintables(doc)
      try {
        // the enrichment ran: every slide carries an underlay
        expect(doc.slides.every(slide => slide.underlay !== undefined), `${file}: slides without an underlay`).toBe(true)
        const actuals = await renderPaintables(paintables)
        const results: string[] = []
        for (let i = 0; i < actuals.length; i++) {
          const goldenPath = join(GOLDEN_DIR, `${file}.${i}.png`)
          if (!existsSync(goldenPath)) {
            if (strict) throw new Error(`missing golden ${goldenPath} (strict mode) — record it first`)
            mkdirSync(GOLDEN_DIR, { recursive: true })
            await savePng(actuals[i], goldenPath)
            results.push(`${i}: recorded`)
            continue
          }
          const d = diffBitmaps(await loadPng(goldenPath), actuals[i], THRESHOLD)
          results.push(`${i}: ${(d.ratio * 100).toFixed(3)}%`)
          expect(d.ratio, `slide ${i}: ${d.changedPixels.toLocaleString()} px changed — inspect tests/goldens/appearance/${file}.${i}.png`).toBeLessThanOrEqual(MAX_RATIO)
        }
        expect(existsSync(join(GOLDEN_DIR, `${file}.${actuals.length}.png`)), `${file}: stale golden beyond the slide count`).toBe(false)
        if (results.some(r => r.endsWith('recorded'))) console.warn(`[golden] auto-recorded ${file}:`, results.join(' '))
      } finally { paintables.dispose() }
    }, 30_000)
  }
})
