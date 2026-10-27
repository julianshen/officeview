// XLSX sparkline reference comparison (plan 2026-10-27-xlsx-sparklines Task 7, spec section 8 and section 11).
//
//   bun scripts/sparkline-compare.ts [--stage N] [--case id] [--out dir]
//
// For every case of stage ≤ N (default 1) with a reference under tests/fixtures/xlsx-sparklines/refs/, renders the
// case workbook through the viewer's static sheet renderer and compares the host cells with Excel's capture.
// Prints `<id>: pass|fail|unverified meanAbs rowMax (tol …)` per case and a total line. `--out` writes
// <id>.ref.png, <id>.ours.png and <id>.diff.png (the cropped host block) for inspection.
// Exit status: 0 when no compared case fails, 1 otherwise, 2 on a usage error.
//
// Rendering. The references are Excel for Mac `copy picture` bitmaps at a fixed 1 px per point (72 dpi); the viewer
// lays out in CSS px (96 dpi). So the sheet renders through `renderSheet` with a 0.75 canvas transform AND
// `sparklineScale = 0.75`: exactly what the interactive viewport does at 75 % zoom, so the geometry snaps its vertices
// on the same 1 px-per-point device grid Excel snapped on. (`getPaintables` renders at scale 1; scaling its bitmap
// down afterwards would resample snapped strokes, which is not how either side draws.)
//
// Cell size. Excel for Mac at 72 dpi measures a 40.71-character column as 204 px (pt), while the viewer's 96 dpi
// column width (285 CSS px) is 213.75 pt. Slot x-positions scale with the plot width, so comparing at the natural
// widths would move vertices by up to ~5 px for a reason that is not sparkline geometry. The host columns and rows
// are therefore forced to the reference's measured gridline-to-gridline size (found from the `-bordered` twin's
// black border pixels), and the natural-versus-reference size is reported separately on every case line.
//
// Crop. The host block is gridline-to-gridline: in the reference [left border, right border) × [top border, bottom
// border) — the left/top border pixels included, the right/bottom ones excluded — and in the viewer
// [cell.x, cell.x + w) × [cell.y, cell.y + h) scaled by 0.75, where the viewer also draws a cell's gridline on its
// first pixel column/row. The PLAIN capture is compared; its border pixels are white.
//
// Units: meanAbs and rowMax are fractions of 255, as in scripts/slideshow-compare.ts: the pixel difference is the
// largest absolute R, G, B or A difference; meanAbs is its mean over the block and rowMax the largest row mean.
// `strong` counts pixels (of either image) that have no pixel in the same column within ±1 row of the other image
// closer than STRONG_DELTA
// levels (pixel-diff.ts's changed-pixel convention, with the doc's ±1 px vertex tolerance built in: the vertical
// rounding is underdetermined at ±1 px, `verticalMapping`). A vertex 2 px off, a missing marker or a wrong colour
// leaves unmatched pixels that a 200-px-wide row mean would dilute.
import { existsSync, mkdirSync, readFileSync, writeFileSync } from 'node:fs'
import { dirname, join, resolve as resolvePath } from 'node:path'
import { fileURLToPath } from 'node:url'
import { createCanvas, loadImage } from 'canvas'
import { OfficePackage } from '../src/core/zip'
import { parseXlsx } from '../src/xlsx/parse'
import { computeMetrics, renderSheet, type GridMetrics } from '../src/xlsx/render'
import { preloadSparklines } from '../src/xlsx/sparklineLoader'
import type { XlsxSheet } from '../src/xlsx/types'
import { SPARKLINE_CASES, type SparklineCase } from './sparkline-refs/cases'

const REPO = resolvePath(dirname(fileURLToPath(import.meta.url)), '..')
export const FIXTURES = join(REPO, 'tests/fixtures/xlsx-sparklines')
export const REFS = join(FIXTURES, 'refs')
export const DOC = join(REPO, 'docs/validation/xlsx-sparklines.md')
/** Viewer CSS px → capture px (96 dpi → 72 dpi). */
export const COMPARE_SCALE = 0.75
/** A pixel counts as strongly different above this many levels (of 255) in any channel. */
export const STRONG_DELTA = 64

export interface Tolerance { meanAbs: number; rowMax: number; strong: number }
export interface Frame { width: number; height: number; data: Uint8ClampedArray }
export interface BlockDiff { meanAbs: number; rowMax: number; worstRow: number; strong: number }

const pixelDelta = (a: Uint8ClampedArray, i: number, b: Uint8ClampedArray, j: number): number =>
  Math.max(Math.abs(a[i] - b[j]), Math.abs(a[i + 1] - b[j + 1]), Math.abs(a[i + 2] - b[j + 2]), Math.abs(a[i + 3] - b[j + 3]))

/** Pixels of `a` with no pixel of `b` in the same column within ±1 row closer than STRONG_DELTA. */
function unmatched(a: Frame, b: Frame): number {
  let n = 0
  for (let y = 0; y < a.height; y++) for (let x = 0; x < a.width; x++) {
    const i = (y * a.width + x) * 4
    let best = 255
    for (let yy = Math.max(0, y - 1); yy <= Math.min(b.height - 1, y + 1); yy++) best = Math.min(best, pixelDelta(a.data, i, b.data, (yy * b.width + x) * 4))
    if (best > STRONG_DELTA) n++
  }
  return n
}

export function diffFrames(a: Frame, b: Frame): BlockDiff {
  if (a.width !== b.width || a.height !== b.height) throw new Error(`size mismatch ${a.width}x${a.height} vs ${b.width}x${b.height}`)
  let total = 0, rowMax = 0, worstRow = 0
  for (let y = 0; y < a.height; y++) {
    let row = 0
    for (let i = y * a.width * 4, end = i + a.width * 4; i < end; i += 4) row += pixelDelta(a.data, i, b.data, i)
    total += row
    const mean = row / a.width / 255
    if (mean > rowMax) { rowMax = mean; worstRow = y }
  }
  return { meanAbs: total / (a.width * a.height) / 255, rowMax, worstRow, strong: unmatched(a, b) + unmatched(b, a) }
}

export const withinTolerance = (d: BlockDiff, t: Tolerance): boolean => d.meanAbs <= t.meanAbs && d.rowMax <= t.rowMax && d.strong <= t.strong

// ---------------------------------------------------------------------------------------------------------------
// Doc table: the Cases table of docs/validation/xlsx-sparklines.md is the source of tolerances and statuses.
// ---------------------------------------------------------------------------------------------------------------

export interface CaseEntry { id: string; status: string; tolerance?: Tolerance; unverifiedReason?: string }

/** Parses `| id | status | meanAbs | rowMax | strong | tolerance | note |` rows; tolerance is `meanAbs/rowMax/strong`. */
export function readCaseTable(markdown = readFileSync(DOC, 'utf8')): Map<string, CaseEntry> {
  const out = new Map<string, CaseEntry>()
  const at = markdown.indexOf('## Cases')
  if (at < 0) return out
  for (const line of markdown.slice(at).split('\n')) {
    const cells = line.split('|').slice(1, -1).map(c => c.trim())
    if (cells.length < 2 || !/^[a-z0-9-_]+$/.test(cells[0]) || cells[0] === 'Case') continue
    const [id, status, , , , tol, note] = cells
    const entry: CaseEntry = { id, status }
    const m = /^([\d.]+)\s*\/\s*([\d.]+)\s*\/\s*(\d+)$/.exec(tol ?? '')
    if (m) entry.tolerance = { meanAbs: Number(m[1]), rowMax: Number(m[2]), strong: Number(m[3]) }
    if (/^\*?unverified/.test(status)) entry.unverifiedReason = (note ?? '').trim()
    out.set(id, entry)
  }
  return out
}

// ---------------------------------------------------------------------------------------------------------------
// Reference geometry
// ---------------------------------------------------------------------------------------------------------------

export async function loadFrame(path: string): Promise<Frame> {
  const img = await loadImage(path)
  const canvas = createCanvas(img.width, img.height), ctx = canvas.getContext('2d')
  // Excel's copy-picture bitmap is transparent where nothing paints (coverage lives in alpha); the sheet behind it
  // is white, so the reference is composited over white like the viewer's own white sheet background.
  ctx.fillStyle = '#ffffff'
  ctx.fillRect(0, 0, img.width, img.height)
  ctx.drawImage(img, 0, 0)
  return { width: img.width, height: img.height, data: ctx.getImageData(0, 0, img.width, img.height).data as unknown as Uint8ClampedArray }
}

export function crop(f: Frame, x: number, y: number, w: number, h: number): Frame {
  const data = new Uint8ClampedArray(w * h * 4)
  for (let row = 0; row < h; row++) {
    const sy = y + row
    for (let col = 0; col < w; col++) {
      const sx = x + col, o = (row * w + col) * 4
      if (sx < 0 || sy < 0 || sx >= f.width || sy >= f.height) { data[o] = data[o + 1] = data[o + 2] = data[o + 3] = 255; continue }
      const i = (sy * f.width + sx) * 4
      data[o] = f.data[i]; data[o + 1] = f.data[i + 1]; data[o + 2] = f.data[i + 2]; data[o + 3] = f.data[i + 3]
    }
  }
  return { width: w, height: h, data }
}

/**
 * Drops the gridline pixels of a gridline-to-gridline host block: its first column and the given rows (each host
 * row's top gridline). The static renderer always draws gridlines (it does not honour showGridLines = 0, a
 * non-sparkline difference), and the line never reaches a gridline (`layering`: ≥ 2 pt in from every inner edge), so
 * the comparison covers each host row's inner box exactly as the doc's measurements define it.
 */
export function innerBox(f: Frame, gridRows: readonly number[]): Frame {
  const keep: number[] = []
  for (let y = 0; y < f.height; y++) if (!gridRows.includes(y)) keep.push(y)
  const w = f.width - 1, data = new Uint8ClampedArray(w * keep.length * 4)
  keep.forEach((sy, row) => data.set(f.data.subarray((sy * f.width + 1) * 4, (sy * f.width + f.width) * 4), row * w * 4))
  return { width: w, height: keep.length, data }
}

export interface HostBorders { xs: number[]; ys: number[] }
/** Columns/rows of the `-bordered` capture that are at least half black: the thin host borders. */
export function findBorders(f: Frame): HostBorders {
  const black = (i: number) => f.data[i] < 64 && f.data[i + 1] < 64 && f.data[i + 2] < 64 && f.data[i + 3] > 192
  const xs: number[] = [], ys: number[] = []
  for (let x = 0; x < f.width; x++) {
    let n = 0
    for (let y = 0; y < f.height; y++) if (black((y * f.width + x) * 4)) n++
    if (n * 2 >= f.height) xs.push(x)
  }
  for (let y = 0; y < f.height; y++) {
    let n = 0
    const x0 = xs.length >= 2 ? xs[xs.length - 2] : 0
    for (let x = x0; x < f.width; x++) if (black((y * f.width + x) * 4)) n++
    if (n * 2 >= f.width - x0) ys.push(y)
  }
  return { xs, ys }
}

export function refBase(c: SparklineCase, sheet: string, bordered: boolean): string {
  return `${c.id}${bordered ? '-bordered' : ''}${c.sheets.length > 1 ? `-${sheet}` : ''}`
}

// ---------------------------------------------------------------------------------------------------------------
// Viewer render
// ---------------------------------------------------------------------------------------------------------------

export interface HostBlock { row0: number; row1: number; col0: number; col1: number }
/** The host rows and columns of every authored sparkline (merged hosts expand to their range). */
export function hostBlock(sheet: XlsxSheet): HostBlock | undefined {
  let row0 = Infinity, row1 = -1, col0 = Infinity, col1 = -1
  for (const group of sheet.sparklineGroups ?? []) for (const { host } of group.sparklines) {
    if (!host) continue
    const merge = sheet.mergeRanges?.find(r => r.minRow === host.row && r.minCol === host.col)
    row0 = Math.min(row0, host.row); row1 = Math.max(row1, merge?.maxRow ?? host.row)
    col0 = Math.min(col0, host.col); col1 = Math.max(col1, merge?.maxCol ?? host.col)
  }
  return row1 < 0 ? undefined : { row0, row1, col0, col1 }
}

export interface Render { frame: Frame; block: { x: number; y: number; w: number; h: number }; natural: { w: number; h: number } }
/**
 * Renders the sheet at COMPARE_SCALE with the host block forced to `size` (capture px), the reference's
 * gridline-to-gridline host size; `natural` is the viewer's own host size at the comparison scale.
 */
export function renderHosts(sheet: XlsxSheet, size: { w: number; h: number }): Render {
  const host = hostBlock(sheet)
  if (!host) throw new Error('no sparkline host')
  const natural = computeMetrics(sheet)
  const sum = (a: number[], from: number, to: number) => a.slice(from, to + 1).reduce((s, v) => s + v, 0)
  const m: GridMetrics = { ...natural, colWidthsPx: [...natural.colWidthsPx], rowHeightsPx: [...natural.rowHeightsPx] }
  // Spread the forced size over the visible host columns/rows in proportion to their natural sizes.
  const force = (sizes: number[], from: number, to: number, target: number) => {
    const total = sum(sizes, from, to)
    for (let i = from; i <= to; i++) if (sizes[i] > 0) sizes[i] = sizes[i] / total * target
  }
  force(m.colWidthsPx, host.col0, host.col1, size.w / COMPARE_SCALE)
  force(m.rowHeightsPx, host.row0, host.row1, size.h / COMPARE_SCALE)
  m.widthPx = m.colWidthsPx.reduce((a, b) => a + b, 0)
  m.heightPx = m.rowHeightsPx.reduce((a, b) => a + b, 0)
  const canvas = createCanvas(Math.ceil(m.widthPx * COMPARE_SCALE), Math.ceil(m.heightPx * COMPARE_SCALE))
  const ctx = canvas.getContext('2d')
  ctx.scale(COMPARE_SCALE, COMPARE_SCALE)
  renderSheet(sheet, ctx as unknown as CanvasRenderingContext2D, m, undefined, undefined, undefined, COMPARE_SCALE)
  const frame: Frame = { width: canvas.width, height: canvas.height, data: ctx.getImageData(0, 0, canvas.width, canvas.height).data as unknown as Uint8ClampedArray }
  const x = Math.round(sum(m.colWidthsPx, 0, host.col0 - 1) * COMPARE_SCALE)
  const y = Math.round(sum(m.rowHeightsPx, 0, host.row0 - 1) * COMPARE_SCALE)
  return {
    frame,
    block: { x, y, w: Math.round(sum(m.colWidthsPx, host.col0, host.col1) * COMPARE_SCALE), h: Math.round(sum(m.rowHeightsPx, host.row0, host.row1) * COMPARE_SCALE) },
    natural: { w: sum(natural.colWidthsPx, host.col0, host.col1) * COMPARE_SCALE, h: sum(natural.rowHeightsPx, host.row0, host.row1) * COMPARE_SCALE },
  }
}

export async function loadCaseSheet(c: SparklineCase, sheetName: string): Promise<XlsxSheet> {
  await preloadSparklines()
  const doc = await parseXlsx(await OfficePackage.load(readFileSync(join(FIXTURES, `${c.id}.xlsx`))))
  const sheet = doc.sheets.find(s => s.name === sheetName)
  if (!sheet) throw new Error(`${c.id}: no sheet ${sheetName}`)
  return sheet
}

// ---------------------------------------------------------------------------------------------------------------
// One case
// ---------------------------------------------------------------------------------------------------------------

export interface CaseResult {
  id: string
  sheet: string
  diff: BlockDiff
  /** Reference host block size (capture px) and the viewer's natural size at the comparison scale. */
  refSize: { w: number; h: number }
  naturalSize: { w: number; h: number }
  ref: Frame
  ours: Frame
}

export interface CompareOptions {
  /** Perturb the parsed sheet (its resolved sparklines) before rendering: negative controls. */
  mutate?: (sheet: XlsxSheet) => void
}

/** Compares every host sheet of a case that has a reference. */
export async function compareCase(c: SparklineCase, opts: CompareOptions = {}): Promise<CaseResult[]> {
  const results: CaseResult[] = []
  for (const s of c.sheets) {
    if (!s.groups.length) continue
    const plainPath = join(REFS, `${refBase(c, s.name, false)}.png`)
    const borderedPath = join(REFS, `${refBase(c, s.name, true)}.png`)
    if (!existsSync(plainPath) || !existsSync(borderedPath)) continue
    const plain = await loadFrame(plainPath), bordered = await loadFrame(borderedPath)
    const { xs, ys } = findBorders(bordered)
    if (xs.length < 2 || ys.length < 2) throw new Error(`${c.id}: host borders not found (x ${xs}, y ${ys})`)
    const rx = xs[xs.length - 2], rw = xs[xs.length - 1] - rx
    const ry = ys[0], rh = ys[ys.length - 1] - ry
    const sheet = await loadCaseSheet(c, s.name)
    opts.mutate?.(sheet)
    const render = renderHosts(sheet, { w: rw, h: rh })
    const gridRows = ys.slice(0, -1).map(y => y - ry)
    const ref = innerBox(crop(plain, rx, ry, rw, rh), gridRows)
    const ours = innerBox(crop(render.frame, render.block.x, render.block.y, rw, rh), gridRows)
    results.push({ id: c.id, sheet: s.name, diff: diffFrames(ref, ours), refSize: { w: rw, h: rh }, naturalSize: render.natural, ref, ours })
  }
  return results
}

/** Worst metrics over a case's sheets. */
export function combine(results: CaseResult[]): BlockDiff {
  return results.reduce<BlockDiff>((a, r) => ({
    meanAbs: Math.max(a.meanAbs, r.diff.meanAbs), rowMax: Math.max(a.rowMax, r.diff.rowMax),
    worstRow: r.diff.rowMax > a.rowMax ? r.diff.worstRow : a.worstRow, strong: a.strong + r.diff.strong,
  }), { meanAbs: 0, rowMax: 0, worstRow: -1, strong: 0 })
}

export async function savePng(f: Frame, path: string): Promise<void> {
  const canvas = createCanvas(f.width, f.height), ctx = canvas.getContext('2d')
  const img = ctx.createImageData(f.width, f.height)
  img.data.set(f.data)
  ctx.putImageData(img, 0, 0)
  writeFileSync(path, canvas.toBuffer('image/png'))
}

function diffImage(a: Frame, b: Frame): Frame {
  const data = new Uint8ClampedArray(a.data.length)
  for (let i = 0; i < data.length; i += 4) {
    const d = Math.max(Math.abs(a.data[i] - b.data[i]), Math.abs(a.data[i + 1] - b.data[i + 1]), Math.abs(a.data[i + 2] - b.data[i + 2]))
    data[i] = 255; data[i + 1] = data[i + 2] = 255 - d; data[i + 3] = 255
  }
  return { width: a.width, height: a.height, data }
}

export const stageCases = (stage: number, only?: string): SparklineCase[] =>
  SPARKLINE_CASES.filter(c => c.stage <= stage && (!only || c.id === only))

const fmt = (n: number) => n.toFixed(4)

export async function main(argv: string[]): Promise<number> {
  const arg = (name: string) => { const i = argv.indexOf(name); return i >= 0 ? argv[i + 1] : undefined }
  const stage = Number(arg('--stage') ?? 1)
  const only = arg('--case'), out = arg('--out')
  if (!Number.isInteger(stage) || stage < 1) { console.error('usage: bun scripts/sparkline-compare.ts [--stage N] [--case id] [--out dir]'); return 2 }
  if (out) mkdirSync(out, { recursive: true })
  const table = readCaseTable()
  let pass = 0, fail = 0, unverified = 0, missing = 0
  for (const c of stageCases(stage, only)) {
    const entry = table.get(c.id)
    const results = await compareCase(c)
    if (!results.length) { missing++; console.log(`${c.id}: no reference`); continue }
    const d = combine(results)
    const tol = entry?.tolerance
    const size = results.map(r => {
      const nat = `${r.naturalSize.w.toFixed(2)}x${r.naturalSize.h.toFixed(2)}`
      return r.naturalSize.w === r.refSize.w && r.naturalSize.h === r.refSize.h ? '' : ` [cell ${r.sheet}: ref ${r.refSize.w}x${r.refSize.h}, viewer ${nat}]`
    }).join('')
    let verdict: string
    if (entry?.unverifiedReason !== undefined) { verdict = 'unverified'; unverified++ }
    else if (tol && withinTolerance(d, tol)) { verdict = 'pass'; pass++ }
    else { verdict = 'fail'; fail++ }
    const tolText = tol ? `tol ${fmt(tol.meanAbs)}/${fmt(tol.rowMax)}/${tol.strong}` : 'tol none'
    console.log(`${c.id}: ${verdict} ${fmt(d.meanAbs)} ${fmt(d.rowMax)} strong ${d.strong} row ${d.worstRow} (${tolText})${size}`)
    if (out) for (const r of results) {
      const base = join(out, results.length > 1 ? `${c.id}-${r.sheet}` : c.id)
      await savePng(r.ref, `${base}.ref.png`); await savePng(r.ours, `${base}.ours.png`); await savePng(diffImage(r.ref, r.ours), `${base}.diff.png`)
    }
  }
  console.log(`total: ${pass} pass, ${fail} fail, ${unverified} unverified${missing ? `, ${missing} without reference` : ''}`)
  return fail ? 1 : 0
}

if ((import.meta as { main?: boolean }).main) process.exit(await main(process.argv.slice(2)))
