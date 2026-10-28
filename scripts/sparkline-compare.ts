// XLSX sparkline reference comparison (plan 2026-10-27-xlsx-sparklines Task 7, spec section 8 and section 11).
//
//   bun scripts/sparkline-compare.ts [--stage N] [--case id] [--out dir] [--refs dir]
//
// For every case of stage ≤ N (default 1), renders the case workbook through the viewer's static sheet renderer and
// compares the host cells with Excel's capture under tests/fixtures/xlsx-sparklines/refs/ (or --refs). Every captured
// sheet needs its plain and -bordered PNG and both .json sidecars; the sidecar's route, scale/pxPerPt, sheet and
// range are validated and the render scale is derived from pxPerPt. A missing or unsupported reference FAILS the
// case unless the doc marks it *unverified*. Prints `<id>: pass|fail|unverified meanAbs rowMax (tol …)` per case and
// a total line. `--out` writes <id>.ref.png, <id>.ours.png and <id>.diff.png (the cropped host block).
// With --stage 3 (or --case xlsxwriter-written[-row9]) it also compares the XlsxWriter-written workbook, an
// independent-writer check with a forced cell size and no -bordered twin (see `compareXlsxWriter`).
// Exit status: 0 when no case fails, 1 otherwise, 2 on a usage error.
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
import { SPARKLINE_CASES, captureRange, type SparklineCase } from './sparkline-refs/cases'

const REPO = resolvePath(dirname(fileURLToPath(import.meta.url)), '..')
export const FIXTURES = join(REPO, 'tests/fixtures/xlsx-sparklines')
export const REFS = join(FIXTURES, 'refs')
export const DOC = join(REPO, 'docs/validation/xlsx-sparklines.md')
/** Viewer CSS px per point (96 dpi): a capture at `pxPerPt` renders at pxPerPt / CSS_PX_PER_PT. */
export const CSS_PX_PER_PT = 96 / 72
/** The comparison scale of the Stage 1 references (1 px per point, 96 dpi → 72 dpi). */
export const COMPARE_SCALE = 1 / CSS_PX_PER_PT
/**
 * Capture densities the comparison supports. The geometry constants and the Cases tolerances were measured from
 * 1 px/pt copy-picture captures; any other density is refused rather than compared against those tolerances.
 */
export const SUPPORTED_PX_PER_PT: readonly number[] = [1]
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
 * Renders the sheet at `scale` (capture px per CSS px, from the sidecar) with the host block forced to `size` (capture px), the reference's
 * gridline-to-gridline host size; `natural` is the viewer's own host size at the comparison scale.
 */
export function renderHosts(sheet: XlsxSheet, size: { w: number; h: number }, scale = COMPARE_SCALE): Render {
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
  force(m.colWidthsPx, host.col0, host.col1, size.w / scale)
  force(m.rowHeightsPx, host.row0, host.row1, size.h / scale)
  m.widthPx = m.colWidthsPx.reduce((a, b) => a + b, 0)
  m.heightPx = m.rowHeightsPx.reduce((a, b) => a + b, 0)
  const canvas = createCanvas(Math.ceil(m.widthPx * scale), Math.ceil(m.heightPx * scale))
  const ctx = canvas.getContext('2d')
  ctx.scale(scale, scale)
  renderSheet(sheet, ctx as unknown as CanvasRenderingContext2D, m)
  const frame: Frame = { width: canvas.width, height: canvas.height, data: ctx.getImageData(0, 0, canvas.width, canvas.height).data as unknown as Uint8ClampedArray }
  const x = Math.round(sum(m.colWidthsPx, 0, host.col0 - 1) * scale)
  const y = Math.round(sum(m.rowHeightsPx, 0, host.row0 - 1) * scale)
  return {
    frame,
    block: { x, y, w: Math.round(sum(m.colWidthsPx, host.col0, host.col1) * scale), h: Math.round(sum(m.rowHeightsPx, host.row0, host.row1) * scale) },
    natural: { w: sum(natural.colWidthsPx, host.col0, host.col1) * scale, h: sum(natural.rowHeightsPx, host.row0, host.row1) * scale },
  }
}

/** Parses `<FIXTURES>/<file>` and returns its sheet `sheetName`, with the lazy sparkline painter loaded. */
export async function loadSheetFile(file: string, sheetName: string): Promise<XlsxSheet> {
  await preloadSparklines()
  const doc = await parseXlsx(await OfficePackage.load(readFileSync(join(FIXTURES, file))))
  const sheet = doc.sheets.find(s => s.name === sheetName)
  if (!sheet) throw new Error(`${file}: no sheet ${sheetName}`)
  return sheet
}

export const loadCaseSheet = (c: SparklineCase, sheetName: string): Promise<XlsxSheet> => loadSheetFile(`${c.id}.xlsx`, sheetName)

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
  /** Reference directory (default REFS); tests point it at a temporary copy. */
  refsDir?: string
}

/** A reference that cannot be compared: a missing image or sidecar, or unsupported capture metadata. */
export class UnusableReferenceError extends Error {}

export interface Sidecar { route: string; scale: number; pxPerPt: number; sheet: string; range: string }
/**
 * Reads and validates a capture sidecar (scripts/sparkline-refs/capture.ts writes it) against the case sheet it
 * belongs to, and returns the capture density. Unsupported metadata throws UnusableReferenceError.
 */
export function readSidecar(path: string, expected: { sheet: string; range: string }): Sidecar {
  if (!existsSync(path)) throw new UnusableReferenceError(`missing sidecar ${path}`)
  let v: Partial<Sidecar>
  try { v = JSON.parse(readFileSync(path, 'utf8')) } catch { throw new UnusableReferenceError(`unreadable sidecar ${path}`) }
  const bad = (why: string) => new UnusableReferenceError(`unsupported sidecar ${path}: ${why}`)
  if (v.route !== 'copy-picture') throw bad(`route ${JSON.stringify(v.route)} (expected "copy-picture")`)
  if (typeof v.pxPerPt !== 'number' || !Number.isFinite(v.pxPerPt) || v.pxPerPt <= 0) throw bad(`pxPerPt ${JSON.stringify(v.pxPerPt)}`)
  if (v.scale !== v.pxPerPt) throw bad(`scale ${JSON.stringify(v.scale)} differs from pxPerPt ${v.pxPerPt}`)
  if (!SUPPORTED_PX_PER_PT.includes(v.pxPerPt)) throw bad(`pxPerPt ${v.pxPerPt} (supported: ${SUPPORTED_PX_PER_PT.join(', ')}; constants and tolerances were measured at 1 px/pt)`)
  if (v.sheet !== expected.sheet) throw bad(`sheet ${JSON.stringify(v.sheet)} (expected ${JSON.stringify(expected.sheet)})`)
  if (v.range !== expected.range) throw bad(`range ${JSON.stringify(v.range)} (expected ${JSON.stringify(expected.range)})`)
  return v as Sidecar
}

/** Compares every host sheet of a case that has a reference. */
export async function compareCase(c: SparklineCase, opts: CompareOptions = {}): Promise<CaseResult[]> {
  const results: CaseResult[] = []
  const dir = opts.refsDir ?? REFS
  for (const s of c.sheets) {
    // Every sheet capture.ts captures is required: both images and both sidecars, with matching metadata.
    const range = captureRange(s)
    if (!range) continue
    const base = (bordered: boolean) => join(dir, refBase(c, s.name, bordered))
    const plainPath = `${base(false)}.png`, borderedPath = `${base(true)}.png`
    const missing = [plainPath, borderedPath].filter(f => !existsSync(f))
    if (missing.length) throw new UnusableReferenceError(`${c.id}: missing reference ${missing.join(', ')}`)
    const meta = readSidecar(`${base(false)}.json`, { sheet: s.name, range })
    const metaBordered = readSidecar(`${base(true)}.json`, { sheet: s.name, range })
    if (meta.pxPerPt !== metaBordered.pxPerPt) throw new UnusableReferenceError(`${c.id}: plain and bordered captures differ in pxPerPt`)
    const scale = meta.pxPerPt / CSS_PX_PER_PT
    const plain = await loadFrame(plainPath), bordered = await loadFrame(borderedPath)
    const { xs, ys } = findBorders(bordered)
    if (xs.length < 2 || ys.length < 2) throw new Error(`${c.id}: host borders not found (x ${xs}, y ${ys})`)
    const rx = xs[xs.length - 2], rw = xs[xs.length - 1] - rx
    const ry = ys[0], rh = ys[ys.length - 1] - ry
    const sheet = await loadCaseSheet(c, s.name)
    opts.mutate?.(sheet)
    const render = renderHosts(sheet, { w: rw, h: rh }, scale)
    const gridRows = ys.slice(0, -1).map(y => y - ry)
    const ref = innerBox(crop(plain, rx, ry, rw, rh), gridRows)
    const ours = innerBox(crop(render.frame, render.block.x, render.block.y, rw, rh), gridRows)
    results.push({ id: c.id, sheet: s.name, diff: diffFrames(ref, ours), refSize: { w: rw, h: rh }, naturalSize: render.natural, ref, ours })
  }
  if (!results.length) throw new UnusableReferenceError(`${c.id}: no sparkline host sheet to compare`)
  return results
}

// ---------------------------------------------------------------------------------------------------------------
// The XlsxWriter-written workbook (plan Task 13): an independent-writer check against an Excel capture
// ---------------------------------------------------------------------------------------------------------------

/**
 * `xlsxwriter-written.xlsx` (scripts/sparkline-refs/xlsxwriter-case.py) was captured by Excel at 1 px/pt as
 * `refs/xlsxwriter-written.png` + `.json` (range I1:J11 of Sheet1), but has no `-bordered` twin, so the host cell size
 * cannot be read from border pixels as `compareCase` does. It is forced instead to what the other 1 px/pt captures
 * measured for the same layout (column I is 45 px; host column J of 40.7109375 characters is 204 px gridline to
 * gridline; every host row is 60 pt), exactly as `renderHosts` forces `compareCase`'s sizes.
 */
export const XLSXWRITER = {
  id: 'xlsxwriter-written', file: 'xlsxwriter-written.xlsx', sheet: 'Sheet1', range: 'I1:J11',
  /** Left gridline of the host column, its width and one row's height, in capture px. */
  hostX: 45, hostW: 204, rowH: 60, rows: 11,
} as const

/**
 * The Cases rows that gate the XlsxWriter comparison, each over a set of host rows (1-based). Row 9 is the near-flat
 * group-scaled line (values 1..8 on the shared 1..80 scale), where the documented ±1 px vertex ambiguity moves a
 * whole row of ink, so it gets its own measured tolerance instead of loosening the rest.
 */
export const XLSXWRITER_PARTS: ReadonlyArray<{ id: string; rows: readonly number[] }> = [
  { id: 'xlsxwriter-written', rows: [1, 2, 3, 4, 5, 6, 7, 8, 10, 11] },
  { id: 'xlsxwriter-written-row9', rows: [9] },
]

/**
 * Row 11's host also holds the text "Growth": Excel draws it bottom-left in its capture font (a serif face), the
 * static renderer in node-canvas's fallback font, which is the same cell-text difference as `layering`. Ours is
 * therefore rendered WITHOUT that text (its cell value is cleared before rendering), and Excel's own glyphs are
 * covered by one rectangle painted white in both images. The rectangle (x, y, w, h in the row's inner box) is the
 * bounding box of Excel's neutral-grey glyph pixels, x 2–31 and y 49–57, padded by 1 px. It also covers the sparkline's
 * first vertex (x 15–17, y 54–55) and about 20 px of the first segment: the one patch of row 11 that is not compared.
 */
export const XLSXWRITER_TEXT_MASK: { row: number; cell: { row: number; col: number }; rect: readonly [number, number, number, number] } = {
  row: 11, cell: { row: 10, col: 9 }, rect: [1, 48, 33, 11],
}

export interface XlsxWriterResult {
  /** Metrics over the stacked requested rows. */
  diff: BlockDiff
  /** The same metrics per host row (inner box), row 1 first. */
  rows: BlockDiff[]
  refSize: { w: number; h: number }
  naturalSize: { w: number; h: number }
  ref: Frame
  ours: Frame
}

export interface XlsxWriterOptions {
  mutate?: (sheet: XlsxSheet) => void
  refsDir?: string
  /** Host rows (1-based) the aggregate covers; every row's own metrics are always reported. Default: all. */
  rows?: readonly number[]
  /** Skip the text handling (keeps ours' text and the unmasked row 11: the raw text residual). */
  noMask?: boolean
}

export async function compareXlsxWriter(opts: XlsxWriterOptions = {}): Promise<XlsxWriterResult> {
  const x = XLSXWRITER, dir = opts.refsDir ?? REFS
  const png = join(dir, `${x.id}.png`)
  if (!existsSync(png)) throw new UnusableReferenceError(`${x.id}: missing reference ${png}`)
  const meta = readSidecar(join(dir, `${x.id}.json`), { sheet: x.sheet, range: x.range })
  const scale = meta.pxPerPt / CSS_PX_PER_PT
  const w = x.hostW, h = x.rowH * x.rows
  const sheet = await loadSheetFile(x.file, x.sheet)
  if (!opts.noMask) for (const row of sheet.rows) for (const cell of row.cells) if (cell.row === XLSXWRITER_TEXT_MASK.cell.row && cell.col === XLSXWRITER_TEXT_MASK.cell.col) cell.value = null
  opts.mutate?.(sheet)
  const render = renderHosts(sheet, { w, h }, scale)
  const refBlock = crop(await loadFrame(png), x.hostX, 0, w, h), oursBlock = crop(render.frame, render.block.x, render.block.y, w, h)
  // Each host row's inner box (its top gridline row and the left gridline column dropped), text mask applied to both.
  const strip = (f: Frame, r: number): Frame => {
    const box = innerBox(crop(f, 0, r * x.rowH, w, x.rowH), [0])
    if (!opts.noMask && r + 1 === XLSXWRITER_TEXT_MASK.row) {
      const [rx, ry, rw, rh] = XLSXWRITER_TEXT_MASK.rect
      for (let yy = ry; yy < Math.min(ry + rh, box.height); yy++) for (let xx = rx; xx < Math.min(rx + rw, box.width); xx++) box.data.fill(255, (yy * box.width + xx) * 4, (yy * box.width + xx) * 4 + 4)
    }
    return box
  }
  const refRows = Array.from({ length: x.rows }, (_, r) => strip(refBlock, r)), oursRows = Array.from({ length: x.rows }, (_, r) => strip(oursBlock, r))
  const rows = refRows.map((f, r) => diffFrames(f, oursRows[r]))
  const keep = (opts.rows ?? Array.from({ length: x.rows }, (_, i) => i + 1)).map(r => r - 1)
  const stack = (parts: Frame[]): Frame => {
    const data = new Uint8ClampedArray(parts.reduce((n, p) => n + p.data.length, 0))
    let at = 0
    for (const p of parts) { data.set(p.data, at); at += p.data.length }
    return { width: parts[0].width, height: parts.length * parts[0].height, data }
  }
  const ref = stack(keep.map(r => refRows[r])), ours = stack(keep.map(r => oursRows[r]))
  return { diff: diffFrames(ref, ours), rows, refSize: { w, h }, naturalSize: render.natural, ref, ours }
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
  const only = arg('--case'), out = arg('--out'), refsDir = arg('--refs')
  if (!Number.isInteger(stage) || stage < 1) { console.error('usage: bun scripts/sparkline-compare.ts [--stage N] [--case id] [--out dir] [--refs dir]'); return 2 }
  if (out) mkdirSync(out, { recursive: true })
  const table = readCaseTable()
  let pass = 0, fail = 0, unverified = 0
  const selected = stageCases(stage, only)
  const withWriter = !only ? stage >= 3 : XLSXWRITER_PARTS.some(p => p.id === only)
  if (!selected.length && !withWriter) { console.error(`no stage ≤ ${stage} case${only ? ` named ${only}` : ''}`); return 2 }
  for (const c of selected) {
    const entry = table.get(c.id)
    let results: CaseResult[]
    try { results = await compareCase(c, { refsDir }) } catch (error) {
      if (!(error instanceof UnusableReferenceError)) throw error
      // Only a case the doc marks unverified may lack a usable reference.
      if (entry?.unverifiedReason !== undefined) { unverified++; console.log(`${c.id}: unverified (${error.message})`) }
      else { fail++; console.log(`${c.id}: fail (${error.message})`) }
      continue
    }
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
  if (withWriter) for (const part of XLSXWRITER_PARTS.filter(p => !only || p.id === only)) {
    const entry = table.get(part.id)
    try {
      const r = await compareXlsxWriter({ refsDir, rows: part.rows })
      const tol = entry?.tolerance
      const verdict = entry?.unverifiedReason !== undefined ? 'unverified' : tol && withinTolerance(r.diff, tol) ? 'pass' : 'fail'
      if (verdict === 'pass') pass++; else if (verdict === 'fail') fail++; else unverified++
      const tolText = tol ? `tol ${fmt(tol.meanAbs)}/${fmt(tol.rowMax)}/${tol.strong}` : 'tol none'
      console.log(`${part.id}: ${verdict} ${fmt(r.diff.meanAbs)} ${fmt(r.diff.rowMax)} strong ${r.diff.strong} row ${r.diff.worstRow} (${tolText}, host rows ${part.rows.join(',')}) [cell ${XLSXWRITER.sheet}: forced ${r.refSize.w}x${r.refSize.h}, viewer ${r.naturalSize.w.toFixed(2)}x${r.naturalSize.h.toFixed(2)}]`)
      if (part === XLSXWRITER_PARTS[0]) console.log(`${XLSXWRITER.id} per host row (meanAbs rowMax strong): ${r.rows.map((d, i) => `${i + 1}: ${fmt(d.meanAbs)} ${fmt(d.rowMax)} ${d.strong}`).join(' | ')}`)
      if (out) { await savePng(r.ref, join(out, `${part.id}.ref.png`)); await savePng(r.ours, join(out, `${part.id}.ours.png`)); await savePng(diffImage(r.ref, r.ours), join(out, `${part.id}.diff.png`)) }
    } catch (error) {
      if (!(error instanceof UnusableReferenceError)) throw error
      if (entry?.unverifiedReason !== undefined) { unverified++; console.log(`${part.id}: unverified (${error.message})`) }
      else { fail++; console.log(`${part.id}: fail (${error.message})`) }
    }
  }
  console.log(`total: ${pass} pass, ${fail} fail, ${unverified} unverified`)
  return fail ? 1 : 0
}

if ((import.meta as { main?: boolean }).main) process.exit(await main(process.argv.slice(2)))
