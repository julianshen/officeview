/**
 * Run every corpus file through the real render pipeline and report how it
 * behaved. This is the bug-finder: synthetic fixtures pass by construction,
 * real-world files do not.
 *
 *   bun scripts/corpus-report.ts            # all files
 *   bun scripts/corpus-report.ts --only docx # filter by extension
 *   bun scripts/corpus-report.ts --render poi-table_test   # also dump PNGs
 *
 * Exit code 1 if any file fails to parse or renders nothing at all.
 */

import { existsSync, mkdirSync, readFileSync, readdirSync, writeFileSync } from 'node:fs'
import { join } from 'node:path'
import { fileURLToPath } from 'node:url'
import { createCanvas } from 'canvas'
import { loadOfficeFile } from '../src/components/OfficeFile'
import { getPaintables } from '../src/render/paint'
import type { XlsxDocument } from '../src/xlsx/types'
import { hasRecordingLoss, measureXlsxTextCoverage, type XlsxTextCoverage } from './corpus-report-text'
import { buildTextIndex } from '../src/core/search'
import { isOfficeCorpusFile } from '../src/test/corpus'

const CORPUS_DIR = fileURLToPath(new URL('../corpus', import.meta.url))

/**
 * Fixtures that are deliberately malformed upstream — a minimized fuzzer case
 * and friends. Rejecting these IS the correct outcome, so they are not
 * counted as failures.
 */
const EXPECTED_REJECTS = ['poi-crash-']

type Outcome = 'ok' | 'degraded' | 'empty' | 'failed'

interface Report {
  file: string
  ext: string
  outcome: Outcome
  units: number
  inkRatio: number
  /** recorded line characters (legacy DOCX/PPTX count) */
  chars: number
  /** characters the parsed model held (before rendering) */
  modelChars: number
  ms: number
  note: string
  xlsx?: XlsxTextCoverage
}

function inkRatioOf(data: Uint8ClampedArray): number {
  let ink = 0
  const total = data.length / 4
  for (let i = 0; i < data.length; i += 4) {
    // anything meaningfully darker than white counts as drawn content
    if (data[i] < 235 || data[i + 1] < 235 || data[i + 2] < 235) ink++
  }
  return total === 0 ? 0 : ink / total
}

/** Legacy parsed string count for the ordinary DOCX/PPTX diagnostic. */
function modelCharCount(doc: unknown): number {
  let n = 0
  if (!doc || typeof doc !== 'object') return 0
  const d = doc as Record<string, unknown>
  const countParagraphs = (paragraphs: unknown): void => {
    if (!Array.isArray(paragraphs)) return
    for (const p of paragraphs as Array<{ runs?: Array<{ text?: string }> }>) {
      for (const r of p?.runs ?? []) n += (r?.text ?? '').trim().length
    }
  }
  if ('sections' in d) {
    for (const section of d.sections as Array<Record<string, unknown>>) {
      countParagraphs(section.paragraphs)
      for (const block of (section.blocks ?? []) as Array<{ kind: string; table?: { rows: Array<{ cells: Array<{ paragraphs: unknown }> }> } }>) {
        if (block.kind !== 'table' || !block.table) continue
        for (const row of block.table.rows) for (const cell of row.cells) countParagraphs(cell.paragraphs)
      }
    }
  }
  if ('sheets' in d) {
    for (const sheet of d.sheets as Array<{ rows: Array<{ cells: Array<{ value: unknown }> }> }>) {
      for (const row of sheet.rows) {
        for (const cell of row.cells) n += typeof cell.value === 'string' ? cell.value.trim().length : 0
      }
    }
  }
  if ('slides' in d) {
    for (const slide of d.slides as Array<{ shapes: Array<{ textBody?: { paragraphs: unknown } }> }>) {
      for (const shape of slide.shapes) if (shape.textBody) countParagraphs(shape.textBody.paragraphs)
    }
  }
  return n
}

async function runOne(path: string, render: boolean): Promise<Report> {
  const file = path.split('/').pop()!
  const ext = file.split('.').pop() ?? ''
  const started = Date.now()
  try {
    const doc = await loadOfficeFile(readFileSync(path))
    const paintables = await getPaintables(doc as never)
    try {
      const modelChars = modelCharCount(doc)
      const xlsxDoc = 'sheets' in doc ? doc as XlsxDocument : undefined
      let chars = 0
      let xlsx: XlsxTextCoverage | undefined
      try {
        const index = await buildTextIndex(paintables as never)
        if (xlsxDoc) {
          const context = createCanvas(4, 4).getContext('2d') as unknown as CanvasRenderingContext2D
          xlsx = measureXlsxTextCoverage(xlsxDoc, paintables, index, context)
          chars = xlsx.recordedSpanChars
        } else {
          for (const page of index.pages) for (const line of page.lines) chars += line.text.trim().length
        }
      } catch (error) {
        if (xlsxDoc) throw new Error(`XLSX recording measurement unavailable: ${error instanceof Error ? error.message : String(error)}`)
        chars = -1
      }
      let ink = 0
      let pixels = 0
      for (const [i, p] of paintables.entries()) {
        const w = Math.min(4000, Math.max(1, Math.ceil(p.spec.widthPx)))
        const h = Math.min(4000, Math.max(1, Math.ceil(p.spec.heightPx)))
        const canvas = createCanvas(w, h)
        const ctx = canvas.getContext('2d')!
        ctx.fillStyle = '#ffffff'
        ctx.fillRect(0, 0, w, h)
        // Only geometry established as legitimately collapsed may skip replay.
        // Other zero/invalid units remain a measurement or raster failure.
        if (!xlsx?.sheets[i]?.collapsedGeometry) {
          ctx.save()
          ctx.scale(w / p.spec.widthPx, h / p.spec.heightPx)
          try { p.paint(ctx as unknown as CanvasRenderingContext2D) } finally { ctx.restore() }
        }
        const img = ctx.getImageData(0, 0, w, h)
        ink += inkRatioOf(img.data) * (w * h)
        pixels += w * h
        if (render) {
          mkdirSync(join('/tmp/corpus-render'), { recursive: true })
          writeFileSync(join('/tmp/corpus-render', `${file}.unit${i}.png`), canvas.toBuffer('image/png'))
        }
      }
      const ratio = pixels === 0 ? 0 : ink / pixels
      const dropped = xlsx ? hasRecordingLoss(xlsx) : modelChars > 0 && chars < modelChars * 0.5
      const outcome: Outcome = xlsx
        ? dropped ? 'degraded' : 'ok'
        : paintables.length === 0 && chars <= 0
          ? 'empty'
          : paintables.length === 0 || dropped ? 'degraded' : 'ok'
      const note = xlsx ? xlsxNote(xlsx) : dropped ? `dropped ${modelChars - chars}/${modelChars} model chars` : ''
      return { file, ext, outcome, units: paintables.length, inkRatio: ratio, chars, modelChars, ms: Date.now() - started, note, ...(xlsx ? { xlsx } : {}) }
    } finally { paintables.dispose() }
  } catch (e) {
    const message = e instanceof Error ? e.message : String(e)
    return {
      file,
      ext,
      outcome: 'failed',
      units: 0,
      inkRatio: 0,
      chars: 0,
      modelChars: 0,
      ms: Date.now() - started,
      note: message.length > 90 ? `${message.slice(0, 90)}…` : message,
    }
  }
}

function xlsxNote(result: XlsxTextCoverage): string {
  const notes: string[] = []
  for (const sheet of result.sheets) for (const stratum of ['string', 'other'] as const) {
    const c = sheet.coverage[stratum]
    if (c.expected > 0 && c.matched < c.expected * 0.5)
      notes.push(`recording coverage below 50%: sheet ${sheet.sheetIndex} (${sheet.sheetName}) ${stratum === 'string' ? 'strings' : 'other'} ${c.matched}/${c.expected}`)
  }
  if (result.status !== 'comparable') notes.push(...result.diagnostics)
  else if (result.expectedSpanChars === 0) notes.push('no eligible cell text')
  else if (result.matchedSpans < result.expectedSpans) notes.push(...result.diagnostics)
  return notes.join('; ')
}
function xlsxFields(r: XlsxTextCoverage): string {
  const excluded = r.excluded.map(e => `${e.reason}=${e.cells}cells/${e.sourceStringChars}source-strings`).join(' ')
  return `source-strings=${r.allSourceStringChars} eligible-strings=${r.eligibleSourceStringChars} eligible-cells=${r.eligibleCells} formatted=${r.formattedExpectedChars} recorded=${r.recordedSpanChars} matched/expected=${r.matchedSpanChars}/${r.expectedSpanChars} strings=${r.coverage.string.matched}/${r.coverage.string.expected} other=${r.coverage.other.matched}/${r.coverage.other.expected} unmatched-recorded=${r.unmatchedRecordedChars} spans=${r.matchedSpans}/${r.expectedSpans}/${r.recordedSpans} unmatched-recorded-spans=${r.unmatchedRecordedSpans}${excluded ? ` ${excluded}` : ''} clipping=${r.clipping}`
}

const onlyArg = process.argv.indexOf('--only')
const only = onlyArg > 0 ? process.argv[onlyArg + 1] : null
const renderArg = process.argv.indexOf('--render')
const renderOne = renderArg > 0 ? process.argv[renderArg + 1] : null

if (!existsSync(CORPUS_DIR)) {
  console.error('corpus/ not found — run: bun scripts/fetch-corpus.ts')
  process.exit(1)
}

const files = readdirSync(CORPUS_DIR)
  .filter(isOfficeCorpusFile)
  .filter((f) => (only ? f.toLowerCase().endsWith(only.toLowerCase()) : true))
  .sort()

console.log(`corpus report — ${files.length} files${only ? ` (only ${only})` : ''}\n`)

console.log('XLSX counts use UTF-16 code units with per-span trimming; whitespace-only spans count as spans. Shared format/layout/replay omissions, parsed-source preservation, drawing completeness and native appearance are not certified. Identical text+geometry replacements can impersonate cell spans; clipping is not-assessed.\n')

const reports: Report[] = []
for (const f of files) {
  const r = await runOne(join(CORPUS_DIR, f), renderOne !== null && f.includes(renderOne))
  reports.push(r)
  const mark = r.outcome === 'ok' ? '✓' : r.outcome === 'empty' ? '·' : r.outcome === 'degraded' ? '!' : '✗'
  console.log(
    `${mark} ${r.file.padEnd(42)} ${r.ext.padEnd(5)} units=${String(r.units).padStart(3)}  ${r.xlsx ? xlsxFields(r.xlsx) : `text=${String(r.chars).padStart(5)}/${String(r.modelChars).padEnd(5)}`}  ink=${(r.inkRatio * 100).toFixed(2).padStart(6)}%  ${String(r.ms).padStart(5)}ms${r.note ? `  ${r.note}` : ''}`,
  )
}

const failed = reports.filter((r) => r.outcome === 'failed')
const degraded = reports.filter((r) => r.outcome === 'degraded')
const empty = reports.filter((r) => r.outcome === 'empty')
// fixtures that are malformed on purpose: refusing them is the correct result
const expected = failed.filter((r) => EXPECTED_REJECTS.some((n) => r.file.includes(n)))
const unexpected = failed.filter((r) => !EXPECTED_REJECTS.some((n) => r.file.includes(n)))
console.log(
  `\n${reports.length - degraded.length - empty.length - expected.length} ok · ${degraded.length} degraded · ${empty.length} legitimately empty · ${expected.length} correctly rejected (malformed on purpose)`,
)
if (empty.length > 0) console.log(`empty (no text in the file itself): ${empty.map((r) => r.file).join(', ')}`)
if (degraded.length > 0) {
  console.log('\ndegraded (recording/content checks need investigation):')
  for (const d of degraded) console.log(`  ${d.file}  ${d.note}`)
}
if (unexpected.length > 0) {
  console.log('\nunexpected failures:')
  for (const f of unexpected) console.log(`  ${f.file}\n    ${f.note}`)
}
process.exit(unexpected.length > 0 || degraded.length > 0 ? 1 : 0)
