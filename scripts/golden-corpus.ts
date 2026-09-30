/**
 * Corpus golden-image CLI — golden coverage for the real-file corpus.
 *
 *   bun scripts/golden-corpus.ts record [--only SUBSTR] [--threshold N]
 *   bun scripts/golden-corpus.ts compare [--only SUBSTR] [--threshold N] [--max-ratio R]
 *
 * record:
 *   Renders every file in corpus/ through the real pipeline
 *   (loadOfficeFile → getPaintables → renderPaintables) and writes one PNG
 *   per page/sheet/slide to tests/goldens/corpus/<file>.<unit>.png, where
 *   <file> keeps its .docx/.xlsx/.pptx extension (e.g.
 *   poi-Bug66263-table.docx.0.png) so fixtures with the same stem cannot
 *   collide.
 *   - Files that render zero units (e.g. pypptx-minimal.pptx has no slides)
 *     are SKIPPED rather than recorded as blank goldens — a blank golden
 *     would hide regressions where content silently stops rendering.
 *   - Files that fail to parse (e.g. poi-crash-* minimized fuzzer cases,
 *     malformed on purpose) are SKIPPED with a warning.
 *   - Files listed in SKIP_RENDERER_BUGS (known src/ bugs that throw in the
 *     harness) are SKIPPED with a warning — see the constant. The list is
 *     currently empty: poi-56295.xlsx used to oversize its canvas and was
 *     covered here until main's XLSX grid-bounding fix (933d883).
 *   - Units whose natural size exceeds MAX_HARNESS_SIDE are rendered
 *     downscaled by the harness (see renderPaintables) — deterministic, but
 *     lower fidelity, so record/compare print a NOTE rather than passing
 *     silently. A NOTE here means a metrics blowup is being hidden by the
 *     clamp: investigate instead of re-recording blindly.
 *   - Idempotent: goldens whose pixels are exactly identical are left
 *     untouched (mtime/hash preserved); only new or pixel-changed goldens
 *     are written. Stale goldens (unit index >= current unit count, or
 *     orphans whose corpus file is gone) are deleted.
 *
 * compare:
 *   Re-renders every corpus file and diffs against the committed goldens
 *   using the same threshold conventions as tests/golden.test.ts
 *   (default threshold 8, max-ratio 0.005). Writes actual+diff PNGs to
 *   /tmp/officeview-diff/corpus on mismatch; exits 1 on any failure, 0
 *   when all pass. Files with no goldens (legitimately empty / correctly
 *   rejected) are skipped, not failed. Missing goldens for files that DO
 *   render are failures — run the record command first.
 *
 *   bun scripts/fetch-corpus.ts   # restores corpus/ binaries if absent
 *
 * Exit codes: 0 = compare passed / record complete; 1 = failures;
 * 2 = usage error (unknown command, missing corpus/).
 */
import { loadOfficeFile } from '../src/components/OfficeFile'
import { getPaintables } from '../src/render/paint'
import { diffBitmaps, savePng, loadPng, renderPaintables } from '../src/test/pixel-diff'
import { readFileSync, existsSync, mkdirSync, readdirSync, unlinkSync } from 'node:fs'
import { join } from 'node:path'

const CORPUS_DIR = join(import.meta.dir, '..', 'corpus')
const GOLDEN_CORPUS_DIR = join(import.meta.dir, '..', 'tests', 'goldens', 'corpus')
const DIFF_DIR = '/tmp/officeview-diff/corpus'

/** Fixtures that are deliberately malformed upstream — refusing them is correct. */
const EXPECTED_REJECTS = ['poi-crash-']

/**
 * Files that hit a known renderer bug and cannot go through the golden
 * harness until src/ is fixed — skipped in record/compare/test (do NOT
 * "fix" by clamping here; that would hide the bug from the goldens).
 * Currently empty: poi-56295.xlsx (sheet "pets" declared
 * <col min="1" max="1025"> while its data is A1:C10, so metrics sized a
 * 66625px canvas and renderPaintables threw) was covered here until it was
 * fixed by bounding the XLSX grid to the used range (main commit 933d883),
 * which now renders it at 195x200. Re-add entries here if a new
 * throw-in-harness renderer bug appears.
 */
const SKIP_RENDERER_BUGS: string[] = []

function parseArgs(argv: string[]): { cmd: string | undefined; only: string | undefined; threshold: number; maxRatio: number } {
  const [cmd, ...rest] = argv
  let only: string | undefined
  let threshold = 8
  let maxRatio = 0.005
  for (let i = 0; i < rest.length; i++) {
    if (rest[i] === '--only') only = rest[++i]
    else if (rest[i] === '--threshold') threshold = parseInt(rest[++i], 10)
    else if (rest[i] === '--max-ratio') maxRatio = parseFloat(rest[++i])
  }
  return { cmd, only, threshold, maxRatio }
}

function isExpectedReject(file: string): boolean {
  return EXPECTED_REJECTS.some((n) => file.includes(n))
}

function goldenPath(file: string, unit: number): string {
  return join(GOLDEN_CORPUS_DIR, `${file}.${unit}.png`)
}

/** Indices of committed goldens for a corpus file (contiguous 0..N-1). */
function existingGoldenUnits(file: string): number[] {
  const units: number[] = []
  for (let i = 0; ; i++) {
    if (!existsSync(goldenPath(file, i))) break
    units.push(i)
  }
  return units
}

/** Must match MAX_CANVAS_SIDE in src/test/pixel-diff.ts: larger units render downscaled. */
const MAX_HARNESS_SIDE = 8192

async function renderFile(absPath: string) {
  const doc = await loadOfficeFile(readFileSync(absPath))
  const paintables = await getPaintables(doc)
  const bitmaps = await renderPaintables(paintables)
  return { bitmaps, specs: paintables.map((p) => p.spec) }
}

/** Warn when a unit's golden is a downscaled render, not full fidelity. */
function noteScaledUnits(file: string, specs: Array<{ widthPx: number; heightPx: number }>): void {
  specs.forEach((s, i) => {
    if (s.widthPx > MAX_HARNESS_SIDE || s.heightPx > MAX_HARNESS_SIDE) {
      console.log(
        `NOTE ${file}.${i}: natural size ${Math.ceil(s.widthPx)}x${Math.ceil(s.heightPx)} exceeds harness clamp ${MAX_HARNESS_SIDE}px — golden is a downscaled render, investigate a metrics blowup before re-recording`,
      )
    }
  })
}

async function record(only: string | undefined, threshold: number): Promise<number> {
  if (!existsSync(CORPUS_DIR)) {
    console.error('corpus/ not found — run: bun scripts/fetch-corpus.ts')
    return 2
  }
  mkdirSync(GOLDEN_CORPUS_DIR, { recursive: true })
  const files = readdirSync(CORPUS_DIR)
    .filter((f) => /\.(docx|xlsx|pptx)$/i.test(f))
    .filter((f) => (only ? f.includes(only) : true))
    .sort()

  let recorded = 0
  let unchanged = 0
  let skipped = 0
  const errors: string[] = []
  const seen = new Set(files)
  void threshold // record uses an exact-pixel (threshold 0) identity check; flag kept for CLI symmetry

  for (const file of files) {
    if (SKIP_RENDERER_BUGS.includes(file)) {
      console.log(`SKIP ${file}: known renderer bug (see SKIP_RENDERER_BUGS) — no golden recorded`)
      skipped++
      continue
    }
    let bitmaps: Awaited<ReturnType<typeof renderFile>>['bitmaps']
    let specs: Awaited<ReturnType<typeof renderFile>>['specs'] = []
    try {
      ;({ bitmaps, specs } = await renderFile(join(CORPUS_DIR, file)))
    } catch (e) {
      const msg = e instanceof Error ? e.message : String(e)
      if (isExpectedReject(file)) {
        console.log(`SKIP ${file}: correctly rejected (${msg.slice(0, 80)})`)
      } else {
        console.error(`ERROR ${file}: ${msg}`)
        errors.push(file)
      }
      skipped++
      continue
    }
    if (bitmaps.length === 0) {
      // Legitimately empty (e.g. a deck with no slides): never record a blank
      // golden. Remove stale goldens if a previous run left any behind.
      const stale = existingGoldenUnits(file)
      for (const u of stale) unlinkSync(goldenPath(file, u))
      if (stale.length > 0) console.log(`SKIP ${file}: renders nothing — removed ${stale.length} stale golden(s)`)
      else console.log(`SKIP ${file}: renders nothing (no slides/pages — no golden recorded)`)
      skipped++
      continue
    }
    noteScaledUnits(file, specs)
    for (let i = 0; i < bitmaps.length; i++) {
      const path = goldenPath(file, i)
      const bm = bitmaps[i]
      if (existsSync(path)) {
        const golden = await loadPng(path)
        const sameSize = golden.width === bm.width && golden.height === bm.height
        // Exact-pixel check (threshold 0): only skip the write when nothing
        // changed, so intentional re-renders always update the golden.
        const identical = sameSize && diffBitmaps(golden, bm, 0).ratio === 0
        if (identical) {
          unchanged++
          continue
        }
        await savePng(bm, path)
        console.log(`updated ${path} (${bm.width}x${bm.height})`)
        recorded++
      } else {
        await savePng(bm, path)
        console.log(`recorded ${path} (${bm.width}x${bm.height})`)
        recorded++
      }
    }
    // Remove stale trailing units from a previous run with more units.
    for (let i = bitmaps.length; existsSync(goldenPath(file, i)); i++) {
      unlinkSync(goldenPath(file, i))
      console.log(`removed stale ${goldenPath(file, i)}`)
      recorded++
    }
  }

  // Remove orphans whose corpus file no longer exists (only when recording
  // the whole corpus — with --only the file set is intentionally partial).
  if (!only && existsSync(GOLDEN_CORPUS_DIR)) {
    for (const g of readdirSync(GOLDEN_CORPUS_DIR).filter((f) => f.endsWith('.png'))) {
      const m = g.match(/^(.*)\.(\d+)\.png$/)
      if (!m) continue
      if (!seen.has(m[1])) {
        unlinkSync(join(GOLDEN_CORPUS_DIR, g))
        console.log(`removed orphan ${g}`)
        recorded++
      }
    }
  }

  console.log(`\nrecord: ${recorded} written/removed, ${unchanged} unchanged, ${skipped} skipped (${files.length} files)`)
  return errors.length > 0 ? 1 : 0
}

async function compare(only: string | undefined, threshold: number, maxRatio: number): Promise<number> {
  if (!existsSync(CORPUS_DIR)) {
    console.error('corpus/ not found — run: bun scripts/fetch-corpus.ts')
    return 2
  }
  const files = readdirSync(CORPUS_DIR)
    .filter((f) => /\.(docx|xlsx|pptx)$/i.test(f))
    .filter((f) => (only ? f.includes(only) : true))
    .sort()

  let failures = 0
  let passed = 0
  let skipped = 0
  mkdirSync(DIFF_DIR, { recursive: true })

  for (const file of files) {
    if (SKIP_RENDERER_BUGS.includes(file)) {
      console.log(`SKIP ${file}: known renderer bug (see SKIP_RENDERER_BUGS) — excluded from compare`)
      skipped++
      continue
    }
    let bitmaps: Awaited<ReturnType<typeof renderFile>>['bitmaps']
    let specs: Awaited<ReturnType<typeof renderFile>>['specs'] = []
    try {
      ;({ bitmaps, specs } = await renderFile(join(CORPUS_DIR, file)))
    } catch (e) {
      const msg = e instanceof Error ? e.message : String(e)
      if (isExpectedReject(file)) {
        console.log(`SKIP ${file}: correctly rejected`)
        skipped++
        continue
      }
      console.error(`FAIL ${file}: threw ${msg}`)
      failures++
      continue
    }
    const goldenUnits = existingGoldenUnits(file)
    if (bitmaps.length === 0) {
      if (goldenUnits.length > 0) {
        console.error(`FAIL ${file}: renders nothing but ${goldenUnits.length} golden(s) exist (stale)`)
        failures++
      } else {
        console.log(`SKIP ${file}: renders nothing (no golden expected)`)
        skipped++
      }
      continue
    }
    if (goldenUnits.length === 0) {
      console.error(`MISSING goldens for ${file} — run: bun scripts/golden-corpus.ts record`)
      failures++
      continue
    }
    if (goldenUnits.length !== bitmaps.length) {
      console.error(`FAIL ${file}: unit count changed (golden ${goldenUnits.length} vs actual ${bitmaps.length}) — run record to update`)
      failures++
      continue
    }
    noteScaledUnits(file, specs)
    for (let i = 0; i < bitmaps.length; i++) {
      const golden = await loadPng(goldenPath(file, i))
      const d = diffBitmaps(golden, bitmaps[i], threshold)
      const pct = (d.ratio * 100).toFixed(3)
      const label = `${file}.${i}`
      if (d.ratio > maxRatio) {
        failures++
        await savePng(bitmaps[i], join(DIFF_DIR, `${label}.actual.png`))
        await savePng(
          { width: d.width, height: d.height, data: d.mask },
          join(DIFF_DIR, `${label}.diff.png`),
        )
        console.error(`FAIL ${label}: ${d.changedPixels.toLocaleString()}/${d.totalPixels.toLocaleString()} px changed (${pct}% > ${(maxRatio * 100).toFixed(2)}%) — actual+diff PNGs in ${DIFF_DIR}`)
      } else {
        passed++
        console.log(`PASS ${label}: ${pct}% changed (maxDelta ${d.maxDelta})`)
      }
    }
  }

  console.log(`\ncompare: ${passed} passed, ${failures} failed, ${skipped} skipped`)
  return failures > 0 ? 1 : 0
}

const { cmd, only, threshold, maxRatio } = parseArgs(process.argv.slice(2))
if (cmd === 'record') process.exit(await record(only, threshold))
else if (cmd === 'compare') process.exit(await compare(only, threshold, maxRatio))
else {
  console.error('usage: bun scripts/golden-corpus.ts record|compare [--only SUBSTR] [--threshold N] [--max-ratio R]')
  process.exit(2)
}
