/**
 * Corpus golden-image regression suite. Renders every real-world fixture in
 * corpus/ through the real pipeline and pixel-diffs against
 * tests/goldens/corpus/<file>.<unit>.png (one PNG per page/sheet/slide).
 *
 * Goldens are recorded with `bun scripts/golden-corpus.ts record` and
 * compared with `bun scripts/golden-corpus.ts compare` — this suite only
 * compares, never auto-records. Delete a golden and re-run record after an
 * intentional layout change. Missing goldens for files that DO render fail
 * the suite; files that legitimately render nothing (e.g.
 * pypptx-minimal.pptx has no slides) or are malformed on purpose
 * (poi-crash-*) have no goldens and are skipped. Files in
 * SKIP_RENDERER_BUGS (known src/ bugs that throw in the harness) stay as
 * visible skipped tests until the renderer is fixed.
 *
 * The corpus is fetched, not vendored: `bun scripts/fetch-corpus.ts`. When
 * the binaries are absent the suite skips cleanly instead of failing.
 */
import { describe, test, expect } from 'vitest'
import { existsSync, readFileSync, readdirSync } from 'node:fs'
import { join, dirname } from 'node:path'
import { fileURLToPath } from 'node:url'
import { loadOfficeFile } from '../src/components/OfficeFile'
import { getPaintables } from '../src/render/paint'
import { diffBitmaps, loadPng, renderPaintables } from '../src/test/pixel-diff'

const GOLDEN_DIR = join(dirname(fileURLToPath(import.meta.url)), 'goldens', 'corpus')
const CORPUS_DIR = join(dirname(fileURLToPath(import.meta.url)), '..', 'corpus')
// The directory always exists once corpus.lock.json is tracked — the binaries
// are the thing that may be absent, so probe for a real fixture, not the dir.
const PROBE_FIXTURE = 'poi-Bug66263-table.docx'
// Same threshold conventions as tests/golden.test.ts.
const THRESHOLD = 8 // per-channel AA forgiveness
const MAX_RATIO = 0.005 // 0.5% of pixels may differ
/** Fixtures that are deliberately malformed upstream — refusing them is correct. */
const EXPECTED_REJECTS = ['poi-crash-']

/**
 * Files that hit a known renderer bug and cannot go through the golden
 * harness until src/ is fixed — kept as visible skipped tests (do NOT
 * "fix" by clamping here; that would hide the bug from the goldens).
 * - poi-56295.xlsx: sheet "pets" declares <col min=0 max=1024>, so sheet
 *   metrics size to 66625px wide; node-canvas caps at 32767px and
 *   renderPaintables throws "Canvas width cannot exceed 32767".
 *   scripts/corpus-report.ts survives only via its 4000px clamp. The
 *   renderer should size to the used cell range, not the full <cols> span.
 */
const SKIP_RENDERER_BUGS = ['poi-56295.xlsx']

const corpusAvailable = existsSync(join(CORPUS_DIR, PROBE_FIXTURE))

function corpusFiles(): string[] {
  return readdirSync(CORPUS_DIR)
    .filter((f) => /\.(docx|xlsx|pptx)$/i.test(f))
    .filter((f) => !EXPECTED_REJECTS.some((n) => f.includes(n)))
    .sort()
}

function goldenUnits(file: string): number[] {
  const units: number[] = []
  for (let i = 0; ; i++) {
    if (!existsSync(join(GOLDEN_DIR, `${file}.${i}.png`))) break
    units.push(i)
  }
  return units
}

describe('corpus golden pixel regression', () => {
  // The corpus is fetched, not vendored: `bun scripts/fetch-corpus.ts`.
  if (!corpusAvailable) {
    test.skip('corpus fixtures not fetched — run: bun scripts/fetch-corpus.ts', () => {})
    return
  }

  for (const file of corpusFiles()) {
    if (SKIP_RENDERER_BUGS.includes(file)) {
      test.skip(`${file} matches corpus golden (known renderer bug — see SKIP_RENDERER_BUGS)`, () => {})
      continue
    }
    test(`${file} matches corpus golden`, async () => {
      const doc = await loadOfficeFile(readFileSync(join(CORPUS_DIR, file)))
      const paintables = await getPaintables(doc)
      const actuals = await renderPaintables(paintables)
      const expected = goldenUnits(file)

      if (actuals.length === 0) {
        // Legitimately empty fixture (e.g. a deck with no slides): no golden
        // should exist. A golden here would be a blank image hiding regressions.
        expect(expected, `${file} renders nothing — remove stale goldens via: bun scripts/golden-corpus.ts record`).toEqual([])
        return
      }

      expect(
        expected.length,
        `${file}: missing goldens — run: bun scripts/golden-corpus.ts record`,
      ).toBeGreaterThan(0)
      expect(
        actuals.length,
        `${file}: unit count changed (golden ${expected.length} vs actual ${actuals.length}) — re-record: bun scripts/golden-corpus.ts record --only ${file}`,
      ).toBe(expected.length)

      for (let i = 0; i < actuals.length; i++) {
        const golden = await loadPng(join(GOLDEN_DIR, `${file}.${i}.png`))
        const d = diffBitmaps(golden, actuals[i], THRESHOLD)
        expect(
          d.ratio,
          `unit ${i}: ${d.changedPixels.toLocaleString()} px changed — re-record only after an intentional layout change: bun scripts/golden-corpus.ts record --only ${file}`,
        ).toBeLessThanOrEqual(MAX_RATIO)
      }
    }, 30_000)
  }
})
