// XLSX pivot, table, slicer and timeline reference comparison (plan 2026-10-29-xlsx-pivots Task 14, spec §9).
//
//   bun scripts/pivot-compare.ts [--stage N] [--case id] [--out dir] [--refs dir] [--fit] [--no-mask]
//
// For every case of stage ≤ N (default 1), renders the case workbook through the interactive viewport painter and
// compares it with Excel's capture under tests/fixtures/xlsx-pivots/refs/ (or --refs), with the tolerance and status
// of its row in the Cases table of docs/validation/xlsx-pivots.md (read by sparkline-compare's `readCaseTable`, whose
// id pattern also admits the table's `---` rule row; `readPivotCaseTable` drops it). A missing or unusable reference FAILS the case
// unless the doc marks it *unverified*. Prints `<id>: pass|fail|unverified meanAbs rowMax strong row (tol …)` per
// case and a total line; `--out` writes <id>.ref.png, <id>.ours.png and <id>.diff.png (the compared block, masks
// applied); `--fit` also prints each compared case's Cases row with its tolerance fitted as measured + 0.002 / 0.010 /
// 3, rounded up (the sparkline convention); `--no-mask` compares without the masks below (diagnosis only). Exit status: 0 when no case fails, 1 otherwise, 2 on a usage error.
//
// Cases. Every generated reference case (scripts/xlsx-refs/cases.ts: tables, pivots, slicers, timelines), the
// measurement workbooks that have a capture (`measure-*`), and the two window screenshots: `measure-field-buttons`
// (pivot-page-field.xlsx) and `pivot-layout-outline-window` (pivot-layout-outline.xlsx, cropped from Task 10's
// screenshot). `pivot-chart-buttons-all-window` is listed and unverified (its Cases note says why).
//
// Mode. A case renders in the mode of its capture route (spec §7.10):
// - copy picture (`route: copy-picture`, appearance screen) draws what the screen shows minus the field dropdowns:
//   the doc's capture notes and refs/pivot-collapsed-item.png keep the ± buttons and the compact gutter (labels at
//   15 + 10 × level pt), which page mode ('print') drops unless printDrill is set. So these render in SCREEN mode with
//   every pivot's `showDataDropDown` turned off on the freshly parsed definition: ± and the indent stay, the
//   page-field and field-header dropdowns go. Not print mode.
// - window screenshots (`route: window-screenshot`) render in screen mode as parsed, dropdowns included.
//
// Rendering. References are 1 px per point; the viewer lays out at 96/72 CSS px per point. As in
// scripts/sparkline-compare.ts the page renders through the real painter under a 0.75 canvas transform
// (`renderSheetViewport` with zoom 0.75 — exactly the interactive viewer at 75 %), never by resampling a 1× bitmap.
//
// Cell size. Excel for Mac at 72 dpi sizes columns from its own digit width, not the viewer's 96 dpi one (a default
// column is 53 px in the captures, 48 at the viewer's 75 %). There are no bordered twins for these captures (Ruling
// J): the column and row boundaries are read from the capture's own gridlines (#D4D4D4 runs at least 12 px long,
// across a column, or 30 px, along a row) where they show, each snapped to its natural position; boundaries hidden
// under fills (inside a table, below a pivot's top rows) are interpolated in proportion to the natural sizes. The
// columns and rows of the captured range are forced to those sizes through session sizes (`buildSheetGeometry`
// `sizes`), so drawings anchored to cells move with them as in Excel. A copy picture starts on the range's top-left
// gridline; a window crop one pixel inside it (measure-field-buttons: A is 77 px in the copy pictures, its right
// gridline at x 76 in the crop).
//
// Block. Each case compares one block: a table its table range, a pivot the whole capture (A1 to the location's
// bottom-right, page fields included), a slicer or timeline its drawing box plus its 1 px frame, other cases the
// whole capture. Gridline and border pixels on the block's edges are included.
//
// Masks (both images, the same rectangles, in the block). The comparison covers fills, borders, gridlines, buttons and
// control chrome; text is pinned by the display tests and by these masks' geometry only:
// - cell text: each laid-out cell with text has its interior (1 px inside its box; from 2 px left of the text when
//   the text starts at least 12 px in, so a ± button stays compared; widened to a spill; its gridline/border
//   pixels on the box's top and left edges stay compared) replaced by that
//   rectangle's most frequent colour in each image, so the fill under the text is still compared. Text is not:
//   Excel draws Calibri and the viewer whatever its font stack finds, and two strings are localized by Excel
//   (Ruling I) or drawn from the UI language (the timeline labels, doc "Timeline geometry").
//   A dropdown box in the cell (window cases) is left out of the rectangle.
// - slicers: the caption (left of the header icons) and each item's text box, flattened the same way.
// - timelines: the caption, the label rows between the header and the bar (selection label, level picker, year,
//   period labels: Excel's UI language), flattened; and the bar of each selected period, blanked: the viewer
//   fills it flat at the gradient's mid tone, an [approx] (Ruling R). Unselected periods, dividers and the
//   scrollbar are compared.
//
// Units: meanAbs and rowMax are fractions of 255 and `strong` counts unmatched pixels, exactly diffFrames of
// scripts/sparkline-compare.ts (largest channel difference; ±1 row matching; STRONG_DELTA 64).
import { existsSync, mkdirSync, readFileSync } from 'node:fs'
import { dirname, join, resolve as resolvePath } from 'node:path'
import { fileURLToPath } from 'node:url'
import { createCanvas } from 'canvas'
import { OfficePackage } from '../src/core/zip'
import { parseXlsx } from '../src/xlsx/parse'
import { ensureTableStyleEngine } from '../src/xlsx/tableStyleLoader'
import { buildSheetGeometry } from '../src/xlsx/view/viewport'
import { renderSheetViewport } from '../src/xlsx/view/renderViewport'
import { LayoutCache, type CellLayout } from '../src/xlsx/view/textLayout'
import { rangeRect } from '../src/xlsx/regionGrid'
import { dropdownBox, pivotButtons } from '../src/xlsx/pivots/buttons'
import { slicerLayout, timelineLayout } from '../src/xlsx/pivots/slicers/paint'
import type { SheetRect } from '../src/xlsx/tableStyles/model'
import type { XlsxDrawing, XlsxSheet } from '../src/xlsx/types'
import { REF_CASES, type RefKind } from './xlsx-refs/cases'
import { crop, diffFrames, loadFrame, readCaseTable, savePng, withinTolerance, type BlockDiff, type CaseEntry, type Frame } from './sparkline-compare'

const REPO = resolvePath(dirname(fileURLToPath(import.meta.url)), '..')
export const FIXTURES = join(REPO, 'tests/fixtures/xlsx-pivots')
export const REFS = join(FIXTURES, 'refs')
export const DOC = join(REPO, 'docs/validation/xlsx-pivots.md')
const EMU_PER_CSS = 9525
/**
 * Excel's timeline frame in every generated timeline capture, in capture px (docs/validation/xlsx-pivots.md
 * "Timeline geometry": x 450–649, y 105–225, both borders included).
 */
export const TIMELINE_FRAME = { x: 450, y: 105, w: 200, h: 121 } as const
/** Capture px per CSS px: the references are 1 px per pt, the viewer 96/72 CSS px per pt. */
export const SCALE = 0.75
/** Excel's gridline colour in the screen captures, and the channel slack of a gridline pixel. */
const GRIDLINE = 0xd4, GRIDLINE_SLACK = 3

export type CaseSubject = 'table' | 'pivot' | 'slicer' | 'timeline' | 'capture'
export type CaptureRoute = 'copy-picture' | 'window-screenshot'
export interface PivotRefCase {
  id: string
  stage: number
  /** The workbook under FIXTURES; the sidecar's `source` when it names one. */
  file: string
  sheet: string
  route: CaptureRoute
  subject: CaseSubject
}

const SUBJECT: Record<RefKind, CaseSubject> = { table: 'table', pivot: 'pivot', slicer: 'slicer', timeline: 'timeline' }
/** All Stage 1 cases: the generated reference cases, then the measurement captures and window screenshots. */
export const PIVOT_CASES: readonly PivotRefCase[] = [
  ...REF_CASES.map(c => ({ id: c.id, stage: 1, file: `${c.id}.xlsx`, sheet: c.sheet, route: 'copy-picture' as const, subject: SUBJECT[c.kind] })),
  { id: 'measure-precedence', stage: 1, file: 'measure-precedence.xlsx', sheet: 'Pivot', route: 'copy-picture', subject: 'pivot' },
  { id: 'measure-cell-override', stage: 1, file: 'measure-cell-override.xlsx', sheet: 'Pivot', route: 'copy-picture', subject: 'pivot' },
  { id: 'measure-apply-flags', stage: 1, file: 'measure-apply-flags.xlsx', sheet: 'Data', route: 'copy-picture', subject: 'table' },
  { id: 'measure-indent', stage: 1, file: 'measure-indent.xlsx', sheet: 'Data', route: 'copy-picture', subject: 'capture' },
  { id: 'measure-timeline-selection', stage: 1, file: 'measure-timeline-selection.xlsx', sheet: 'Report', route: 'copy-picture', subject: 'timeline' },
  { id: 'measure-field-buttons', stage: 1, file: 'pivot-page-field.xlsx', sheet: 'Pivot', route: 'window-screenshot', subject: 'pivot' },
  { id: 'pivot-layout-outline-window', stage: 1, file: 'pivot-layout-outline.xlsx', sheet: 'Pivot', route: 'window-screenshot', subject: 'pivot' },
  { id: 'pivot-chart-buttons-all-window', stage: 1, file: 'pivot-chart-buttons-all.xlsx', sheet: 'Pivot', route: 'window-screenshot', subject: 'capture' },
]

export const stageCases = (stage: number, only?: string): PivotRefCase[] => PIVOT_CASES.filter(c => c.stage <= stage && (!only || c.id === only))
export function readPivotCaseTable(markdown = readFileSync(DOC, 'utf8')): Map<string, CaseEntry> {
  const table = readCaseTable(markdown)
  for (const id of table.keys()) if (/^-+$/.test(id)) table.delete(id) // the header rule row
  return table
}

/** A reference that cannot be compared: a missing image or sidecar, or unsupported capture metadata. */
export class UnusableReferenceError extends Error {}

export interface Sidecar { route: CaptureRoute; appearance: string; pxPerPt: number; sheet: string; range: string; source?: string }
/** Reads and validates a capture sidecar against its case. */
export function readSidecar(path: string, c: PivotRefCase): Sidecar {
  if (!existsSync(path)) throw new UnusableReferenceError(`missing sidecar ${path}`)
  let v: Partial<Sidecar>
  try { v = JSON.parse(readFileSync(path, 'utf8')) } catch { throw new UnusableReferenceError(`unreadable sidecar ${path}`) }
  const bad = (why: string) => new UnusableReferenceError(`unsupported sidecar ${path}: ${why}`)
  if (v.route !== c.route) throw bad(`route ${JSON.stringify(v.route)} (expected ${JSON.stringify(c.route)})`)
  if (v.appearance !== 'screen') throw bad(`appearance ${JSON.stringify(v.appearance)} (only screen captures exist; print is unverified)`)
  if (v.pxPerPt !== 1) throw bad(`pxPerPt ${JSON.stringify(v.pxPerPt)} (the comparison and its tolerances are at 1 px per pt)`)
  if (v.sheet !== c.sheet) throw bad(`sheet ${JSON.stringify(v.sheet)} (expected ${JSON.stringify(c.sheet)})`)
  if (typeof v.range !== 'string' || !rangeRect(v.range)) throw bad(`range ${JSON.stringify(v.range)}`)
  if (v.source !== undefined && v.source !== c.file) throw bad(`source ${JSON.stringify(v.source)} (expected ${JSON.stringify(c.file)})`)
  return v as Sidecar
}

// ---------------------------------------------------------------------------------------------------------------
// Geometry: the capture's gridlines
// ---------------------------------------------------------------------------------------------------------------

const isGridline = (f: Frame, x: number, y: number): boolean => {
  const i = (y * f.width + x) * 4
  return Math.abs(f.data[i] - GRIDLINE) <= GRIDLINE_SLACK && Math.abs(f.data[i + 1] - GRIDLINE) <= GRIDLINE_SLACK && Math.abs(f.data[i + 2] - GRIDLINE) <= GRIDLINE_SLACK
}
/** Positions along `axis` holding a gridline run of at least `run` px across it. */
export function gridlineCandidates(f: Frame, axis: 'x' | 'y', run: number): number[] {
  const out: number[] = []
  const [n, m] = axis === 'x' ? [f.width, f.height] : [f.height, f.width]
  for (let p = 0; p < n; p++) {
    let length = 0
    for (let q = 0; q < m; q++) {
      length = (axis === 'x' ? isGridline(f, p, q) : isGridline(f, q, p)) ? length + 1 : 0
      if (length >= run) { out.push(p); break }
    }
  }
  return out
}
/**
 * The capture's cell boundaries along one axis: `first` and `last` are the outer gridlines, `natural` the cells'
 * natural sizes in capture px. Each inner boundary snaps to a gridline candidate near its natural position (scaled to
 * the capture's extent); the rest are interpolated between their snapped neighbours by natural size.
 */
export function fitBoundaries(candidates: readonly number[], first: number, last: number, natural: readonly number[]): { at: number[]; snapped: number } {
  const total = natural.reduce((a, b) => a + b, 0)
  const ratio = total > 0 ? (last - first) / total : 1
  const at: Array<number | undefined> = [first]
  let acc = 0, snapped = 0
  for (let k = 1; k < natural.length; k++) {
    acc += natural[k - 1]!
    const expected = first + acc * ratio, slack = Math.max(3, Math.min(natural[k - 1]!, natural[k]!) * ratio * 0.3)
    let best: number | undefined
    for (const c of candidates) if (c > first && c < last && Math.abs(c - expected) <= slack && (best === undefined || Math.abs(c - expected) < Math.abs(best - expected))) best = c
    at.push(best)
    if (best !== undefined) snapped++
  }
  at.push(last)
  // Interpolate the unsnapped boundaries between their snapped neighbours.
  for (let k = 1; k < at.length - 1; k++) {
    if (at[k] !== undefined) continue
    let j = k
    while (at[j] === undefined) j++
    const a = at[k - 1]!, b = at[j]!, span = natural.slice(k - 1, j).reduce((s, v) => s + v, 0)
    let run = 0
    for (let i = k; i < j; i++) { run += natural[i - 1]!; at[i] = Math.round(a + (b - a) * (span > 0 ? run / span : (i - k + 1) / (j - k + 1))) }
  }
  return { at: at as number[], snapped }
}

// ---------------------------------------------------------------------------------------------------------------
// One case
// ---------------------------------------------------------------------------------------------------------------

export interface Rect { x: number; y: number; w: number; h: number }
export interface Mask { rect: Rect; kind: 'flatten' | 'blank' }
export interface CaseResult {
  id: string
  diff: BlockDiff
  ref: Frame
  ours: Frame
  /** Boundaries snapped to visible gridlines / all inner boundaries, per axis. */
  snapped: { cols: string; rows: string }
}
export interface CompareOptions {
  /** Perturb the parsed sheet before rendering (negative controls). */
  mutate?: (sheet: XlsxSheet) => void
  refsDir?: string
  /** Compare without masks (diagnosis only). */
  noMask?: boolean
}

async function loadSheet(c: PivotRefCase): Promise<XlsxSheet> {
  const path = join(FIXTURES, c.file)
  if (!existsSync(path)) throw new UnusableReferenceError(`${c.id}: missing workbook ${path}`)
  const doc = await parseXlsx(await OfficePackage.load(readFileSync(path)))
  await ensureTableStyleEngine(doc.sheets)
  const sheet = doc.sheets.find(s => s.name === c.sheet)
  if (!sheet) throw new Error(`${c.file}: no sheet ${c.sheet}`)
  return sheet
}

const controlDrawing = (drawings: readonly XlsxDrawing[], kind: 'slicer' | 'timeline'): XlsxDrawing | undefined => drawings.find(d => d.content?.kind === kind)

/** Most frequent colour of a rectangle (RGBA packed). */
function modeColour(f: Frame, r: Rect): number | undefined {
  const counts = new Map<number, number>()
  let best: number | undefined, bestCount = 0
  for (let y = Math.max(0, r.y); y < Math.min(f.height, r.y + r.h); y++) for (let x = Math.max(0, r.x); x < Math.min(f.width, r.x + r.w); x++) {
    const i = (y * f.width + x) * 4, key = ((f.data[i] << 24) | (f.data[i + 1] << 16) | (f.data[i + 2] << 8) | f.data[i + 3]) >>> 0
    const n = (counts.get(key) ?? 0) + 1
    counts.set(key, n)
    if (n > bestCount) { bestCount = n; best = key }
  }
  return best
}
export function applyMask(f: Frame, m: Mask): void {
  const colour = m.kind === 'blank' ? 0xffffffff : modeColour(f, m.rect)
  if (colour === undefined) return
  for (let y = Math.max(0, m.rect.y); y < Math.min(f.height, m.rect.y + m.rect.h); y++) for (let x = Math.max(0, m.rect.x); x < Math.min(f.width, m.rect.x + m.rect.w); x++) {
    const i = (y * f.width + x) * 4
    f.data[i] = colour >>> 24; f.data[i + 1] = (colour >>> 16) & 255; f.data[i + 2] = (colour >>> 8) & 255; f.data[i + 3] = colour & 255
  }
}

/** Device-pixel rectangle (capture px) of a logical CSS box. */
const device = (x0: number, y0: number, x1: number, y1: number): Rect => {
  const x = Math.round(x0 * SCALE), y = Math.round(y0 * SCALE)
  return { x, y, w: Math.round(x1 * SCALE) - x, h: Math.round(y1 * SCALE) - y }
}

/**
 * Text masks of the laid-out cells (capture px, in the rendered frame). `buttons` are dropdown boxes to keep;
 * `drill` holds the cells ("row:col") where Excel draws a ± button, whose box (x 6–12) a mask never covers, whatever
 * the viewer laid out there.
 */
export function textMasks(records: readonly CellLayout[], buttons: readonly Rect[], drill: ReadonlySet<string> = new Set()): Mask[] {
  const out: Mask[] = []
  // A medium, thick or double border is wider than the gridline pixel it sits on: Excel draws its second line
  // inside the cells beside it (pivot-medium-2: the grand total's double top border at y 314 and 316 around the
  // row line at 315), so a mask keeps one more pixel on that edge.
  const at = new Map(records.map(r => [`${r.source.row}:${r.source.col}`, r]))
  const heavy = (style: string | undefined) => !!style && /medium|thick|double|slantDashDot/.test(style)
  const edge = (r: CellLayout, side: 'top' | 'bottom' | 'left' | 'right') => heavy(r.body.borders?.[side])
  const neighbour = (r: CellLayout, dr: number, dc: number) => at.get(`${r.source.row + dr}:${r.source.col + dc}`)
  for (const record of records) {
    const text = record.body.text
    if (!text || !record.body.formattedText) continue
    const lb = record.logicalBounds
    const box = device(lb.x, lb.y, lb.x + lb.width, lb.y + lb.height)
    // Text extent in logical px: the body is laid out in a box at the origin.
    let [t0, t1] = text.kind === 'plain' ? [text.tx, text.tx + text.width] : [text.clip.x, text.clip.x + text.clip.width]
    if (text.kind === 'layout' && !text.wrap && !text.rotation && !text.stacked) {
      // One unrotated line: the glyph run's own extent (an indented label starts past its ± button).
      const runs = text.layout.lines.flatMap(line => line.segments.map(seg => { const x = seg.transform ? seg.transform.e : line.x + seg.x; return [x, x + seg.width] }))
      if (runs.length) [t0, t1] = [Math.min(...runs.map(r => r[0]!)), Math.max(...runs.map(r => r[1]!))]
    }
    t0 += lb.x - record.body.x; t1 += lb.x - record.body.x
    const ts = t0 * SCALE, te = t1 * SCALE
    const up = neighbour(record, -1, 0), down = neighbour(record, 1, 0), left = neighbour(record, 0, -1), right = neighbour(record, 0, 1)
    const insetTop = edge(record, 'top') || (up && edge(up, 'bottom')) ? 2 : 1
    const insetBottom = edge(record, 'bottom') || (down && edge(down, 'top')) ? 1 : 0
    const insetLeft = edge(record, 'left') || (left && edge(left, 'right')) ? 2 : 1
    const insetRight = edge(record, 'right') || (right && edge(right, 'left')) ? 1 : 0
    let x0 = ts - box.x >= 12 ? Math.floor(ts) - 2 : box.x + insetLeft
    let x1 = Math.max(box.x + box.w - insetRight, Math.ceil(te) + 3)
    if (ts < box.x) x0 = Math.floor(ts) - 3
    if (drill.has(`${record.source.row}:${record.source.col}`)) x0 = Math.max(x0, box.x + 13)
    for (const b of buttons) if (b.y < box.y + box.h && b.y + b.h > box.y && b.x < x1 && b.x + b.w > x0) x1 = Math.min(x1, b.x - 1)
    const y0 = box.y + insetTop, y1 = box.y + box.h - insetBottom
    if (x1 > x0 && y1 > y0) out.push({ rect: { x: x0, y: y0, w: x1 - x0, h: y1 - y0 }, kind: 'flatten' })
  }
  return out
}

/** Slicer and timeline masks (capture px, in the rendered frame) for a control drawn in `box`. */
export function controlMasks(drawing: XlsxDrawing, box: Rect): Mask[] {
  const out: Mask[] = []
  const b = { x: box.x, y: box.y, width: box.w, height: box.h }
  const content = drawing.content
  if (content?.kind === 'slicer') {
    const layout = slicerLayout(content.slicer, b, 1)
    // The caption, and the header icons (multi-select, clear filter) in the header's right 50 px, which the
    // viewer does not draw (Cases: unverified sub-element).
    if (layout.separatorY !== undefined) {
      out.push({ rect: { x: box.x + 2, y: box.y + 2, w: box.w - 52, h: layout.separatorY - box.y - 3 }, kind: 'flatten' })
      out.push({ rect: { x: box.x + box.w - 50, y: box.y + 2, w: 48, h: layout.separatorY - box.y - 3 }, kind: 'flatten' })
    }
    for (const item of layout.items) {
      const r = item.rect
      out.push({ rect: { x: Math.floor(item.text.x) - 2, y: Math.round(r.y) + 2, w: Math.round(r.x + r.width) - 2 - (Math.floor(item.text.x) - 2), h: Math.round(r.height) - 4 }, kind: 'flatten' })
    }
  } else if (content?.kind === 'timeline') {
    const layout = timelineLayout(content.timeline, b, 1)
    const bar = layout.bar
    // The caption, and the clear-filter icon in the header's right 30 px, which the viewer does not draw.
    if (layout.separatorY !== undefined) {
      out.push({ rect: { x: box.x + 2, y: box.y + 2, w: box.w - 32, h: layout.separatorY - box.y - 3 }, kind: 'flatten' })
      out.push({ rect: { x: box.x + box.w - 30, y: box.y + 2, w: 28, h: layout.separatorY - box.y - 3 }, kind: 'flatten' })
      out.push({ rect: { x: box.x + 2, y: Math.round(layout.separatorY) + 2, w: box.w - 4, h: Math.round(bar.y) - 2 - (Math.round(layout.separatorY) + 2) }, kind: 'flatten' })
    }
    // The bar: with no selection every period is selected and the whole bar is left out (its gradient and
    // Excel's calendar-length period spacing are [approx]); with a selection, the selected periods are.
    if (!content.timeline.selection) out.push({ rect: { x: Math.round(bar.x), y: Math.round(bar.y), w: Math.round(bar.width), h: Math.round(bar.height) }, kind: 'blank' })
    else for (const p of layout.periods) {
      if (!p.selected) continue
      const x0 = Math.max(Math.round(bar.x), Math.round(p.x)), x1 = Math.min(Math.round(bar.x + bar.width), Math.round(p.x + p.width) - 1)
      if (x1 > x0) out.push({ rect: { x: x0, y: Math.round(bar.y), w: x1 - x0, h: Math.round(bar.height) }, kind: 'blank' })
    }
    // The frame: Excel's 1 px timeline frame is a partial-coverage accent ([approx]: Light1 #5887BF left,
    // #7095C2 bottom, #7BA0CD right over #FFFFFF; the viewer draws the accent).
    out.push({ rect: { x: box.x, y: box.y, w: box.w, h: 1 }, kind: 'blank' }, { rect: { x: box.x, y: box.y + box.h - 1, w: box.w, h: 1 }, kind: 'blank' })
    out.push({ rect: { x: box.x, y: box.y, w: 1, h: box.h }, kind: 'blank' }, { rect: { x: box.x + box.w - 1, y: box.y, w: 1, h: box.h }, kind: 'blank' })
  }
  return out
}

/** Renders and compares one case. */
export async function compareCase(c: PivotRefCase, opts: CompareOptions = {}): Promise<CaseResult> {
  const dir = opts.refsDir ?? REFS
  const png = join(dir, `${c.id}.png`)
  if (!existsSync(png)) throw new UnusableReferenceError(`${c.id}: missing reference ${png}`)
  const meta = readSidecar(join(dir, `${c.id}.json`), c)
  const range = rangeRect(meta.range)!
  const ref = await loadFrame(png)
  const sheet = await loadSheet(c)
  // Where Excel draws ± buttons, from the saved definition before any perturbation (a mask never covers them).
  const drill = new Set((sheet.pivotTables ?? []).flatMap(def => pivotButtons({ ...def, showDrill: true }, 'screen')
    .filter(b => b.kind === 'expand' || b.kind === 'collapse').map(b => `${b.row}:${b.col}`)))
  if (c.route === 'copy-picture') for (const def of sheet.pivotTables ?? []) def.showDataDropDown = false
  opts.mutate?.(sheet)

  // Natural sizes (capture px) of the captured columns and rows, then the capture's own boundaries.
  const natural = buildSheetGeometry(sheet)
  if (natural.kind !== 'ready') throw new Error(`${c.id}: geometry ${natural.kind}`)
  const cols: number[] = [], rows: number[] = []
  for (let col = range.minCol; col <= range.maxCol; col++) cols.push(natural.metrics.cols.size(col) * SCALE)
  for (let row = range.minRow; row <= range.maxRow; row++) rows.push(natural.metrics.rows.size(natural.metrics.order.sourceToDisplay(row)) * SCALE)
  // A copy picture starts on the top-left gridline; a window crop one pixel inside it.
  const origin = c.route === 'copy-picture' ? 0 : -1
  const xs = fitBoundaries(gridlineCandidates(ref, 'x', 12), origin, ref.width - 1, cols)
  const ys = fitBoundaries(gridlineCandidates(ref, 'y', 30), origin, ref.height - 1, rows)
  const sizes = { cols: new Map<number, number>(), rows: new Map<number, number>() }
  for (let k = 0; k < cols.length; k++) sizes.cols.set(range.minCol + k, (xs.at[k + 1]! - xs.at[k]!) / SCALE)
  for (let k = 0; k < rows.length; k++) sizes.rows.set(range.minRow + k, (ys.at[k + 1]! - ys.at[k]!) / SCALE)
  const geometry = buildSheetGeometry(sheet, { sizes })
  if (geometry.kind !== 'ready') throw new Error(`${c.id}: forced geometry ${geometry.kind}`)

  // Render the captured range at 75 %: capture pixel (0, 0) is device pixel (-origin, -origin) from the range's corner.
  const vx = geometry.metrics.cols.start(range.minCol) - origin / SCALE, vy = geometry.metrics.rows.start(range.minRow) - origin / SCALE
  if (c.subject === 'timeline') {
    // The generated timelines are anchored at H8 + 62 pt in a 53 pt column; Excel draws the frame at
    // TIMELINE_FRAME regardless, by a rule not established. The timeline is placed there (its placement is
    // unverified in Cases) so that its chrome is compared in Excel's box.
    const drawing = controlDrawing(geometry.sizedDrawings ?? sheet.drawings ?? [], 'timeline')
    if (!drawing) throw new Error(`${c.id}: no timeline drawing`)
    drawing.xEmu = (vx + TIMELINE_FRAME.x / SCALE) * EMU_PER_CSS; drawing.yEmu = (vy + TIMELINE_FRAME.y / SCALE) * EMU_PER_CSS
    drawing.widthEmu = TIMELINE_FRAME.w / SCALE * EMU_PER_CSS; drawing.heightEmu = TIMELINE_FRAME.h / SCALE * EMU_PER_CSS
  }
  const canvas = createCanvas(ref.width, ref.height), ctx = canvas.getContext('2d')
  ctx.scale(SCALE, SCALE)
  const result = renderSheetViewport(sheet, ctx as unknown as CanvasRenderingContext2D, geometry, { box: { x: vx, y: vy, width: ref.width / SCALE, height: ref.height / SCALE }, zoom: SCALE, deviceScale: 1 }, undefined, { cache: new LayoutCache() })
  if (result.kind !== 'detail' || result.pending) throw new Error(`${c.id}: viewport painted ${result.kind}${result.pending ? ' (pending)' : ''}`)
  const ours: Frame = { width: ref.width, height: ref.height, data: ctx.getImageData(0, 0, ref.width, ref.height).data as unknown as Uint8ClampedArray }
  const shift = (r: Rect): Rect => ({ x: r.x - Math.round(vx * SCALE), y: r.y - Math.round(vy * SCALE), w: r.w, h: r.h })

  // The compared block, in capture px.
  let block: Rect = { x: 0, y: 0, w: ref.width, h: ref.height }
  const masks: Mask[] = []
  const cellBox = (rect: SheetRect): Rect => {
    const x0 = geometry.metrics.cols.start(rect.minCol), y0 = geometry.metrics.rows.start(geometry.metrics.order.sourceToDisplay(rect.minRow))
    const x1 = geometry.metrics.cols.start(rect.maxCol + 1), y1 = geometry.metrics.rows.start(geometry.metrics.order.sourceToDisplay(rect.maxRow) + 1)
    const r = shift(device(x0, y0, x1, y1))
    return { x: r.x, y: r.y, w: r.w + 1, h: r.h + 1 }
  }
  if (c.subject === 'table') {
    const extent = sheet.tables?.[0]?.extent
    if (!extent) throw new Error(`${c.id}: no table`)
    block = cellBox({ minRow: extent.firstRow, minCol: extent.firstCol, maxRow: extent.firstRow + extent.rows - 1, maxCol: extent.firstCol + extent.cols - 1 })
  } else if (c.subject === 'slicer' || c.subject === 'timeline') {
    const drawing = controlDrawing(geometry.sizedDrawings ?? sheet.drawings ?? [], c.subject)
    if (!drawing) throw new Error(`${c.id}: no ${c.subject} drawing`)
    const box = shift(device(drawing.xEmu / 9525, drawing.yEmu / 9525, (drawing.xEmu + drawing.widthEmu) / 9525, (drawing.yEmu + drawing.heightEmu) / 9525))
    masks.push(...controlMasks(drawing, box))
    block = { x: box.x - 1, y: box.y - 1, w: box.w + 2, h: box.h + 2 }
  }
  // Dropdown boxes stay compared in the window cases.
  const buttons: Rect[] = []
  if (c.route === 'window-screenshot') for (const def of sheet.pivotTables ?? []) for (const b of pivotButtons(def, 'screen')) {
    if (b.kind !== 'filter' && b.kind !== 'page') continue
    const cell = geometry.metrics.sourceCellToLogical({ row: b.row, col: b.col })
    if (!cell) continue
    const box = dropdownBox({ x: cell.x, y: cell.y, width: cell.width, height: cell.height })
    buttons.push(device(box.x, box.y, box.x + box.width, box.y + box.height))
  }
  masks.push(...textMasks(result.records, buttons, drill).map(m => ({ ...m, rect: shift(m.rect) })))

  const refBlock = crop(ref, block.x, block.y, block.w, block.h), oursBlock = crop(ours, block.x, block.y, block.w, block.h)
  for (const m of opts.noMask ? [] : masks) {
    const local = { kind: m.kind, rect: { x: m.rect.x - block.x, y: m.rect.y - block.y, w: m.rect.w, h: m.rect.h } }
    applyMask(refBlock, local); applyMask(oursBlock, local)
  }
  return {
    id: c.id, diff: diffFrames(refBlock, oursBlock), ref: refBlock, ours: oursBlock,
    snapped: { cols: `${xs.snapped}/${cols.length - 1}`, rows: `${ys.snapped}/${rows.length - 1}` },
  }
}

// ---------------------------------------------------------------------------------------------------------------
// CLI
// ---------------------------------------------------------------------------------------------------------------

function diffImage(a: Frame, b: Frame): Frame {
  const data = new Uint8ClampedArray(a.data.length)
  for (let i = 0; i < data.length; i += 4) {
    const d = Math.max(Math.abs(a.data[i] - b.data[i]), Math.abs(a.data[i + 1] - b.data[i + 1]), Math.abs(a.data[i + 2] - b.data[i + 2]))
    data[i] = 255; data[i + 1] = data[i + 2] = 255 - d; data[i + 3] = 255
  }
  return { width: a.width, height: a.height, data }
}

const fmt = (n: number) => n.toFixed(4)
const up = (v: number, places: number) => Math.ceil(v * 10 ** places - 1e-9) / 10 ** places
/** The fitted tolerance: measured + 0.002 / 0.010 / 3, rounded up (meanAbs and rowMax to 0.001). */
export const fitTolerance = (d: BlockDiff): string => `${up(d.meanAbs + 0.002, 3).toFixed(3)} / ${up(d.rowMax + 0.010, 3).toFixed(3)} / ${d.strong + 3}`

export async function main(argv: string[]): Promise<number> {
  const arg = (name: string) => { const i = argv.indexOf(name); return i >= 0 ? argv[i + 1] : undefined }
  const stage = Number(arg('--stage') ?? 1)
  const only = arg('--case'), out = arg('--out'), refsDir = arg('--refs'), fit = argv.includes('--fit'), noMask = argv.includes('--no-mask')
  if (!Number.isInteger(stage) || stage < 1) { console.error('usage: bun scripts/pivot-compare.ts [--stage N] [--case id] [--out dir] [--refs dir] [--fit]'); return 2 }
  const selected = stageCases(stage, only)
  if (!selected.length) { console.error(`no stage ≤ ${stage} case${only ? ` named ${only}` : ''}`); return 2 }
  if (out) mkdirSync(out, { recursive: true })
  const table = readPivotCaseTable()
  const rows: string[] = []
  let pass = 0, fail = 0, unverified = 0
  for (const c of selected) {
    const entry = table.get(c.id)
    let r: CaseResult
    try { r = await compareCase(c, { refsDir, noMask }) } catch (error) {
      if (!(error instanceof UnusableReferenceError)) throw error
      if (entry?.unverifiedReason !== undefined) { unverified++; console.log(`${c.id}: unverified (${error.message})`) }
      else { fail++; console.log(`${c.id}: fail (${error.message})`) }
      continue
    }
    const tol = entry?.tolerance
    let verdict: string
    if (entry?.unverifiedReason !== undefined) { verdict = 'unverified'; unverified++ }
    else if (tol && withinTolerance(r.diff, tol)) { verdict = 'pass'; pass++ }
    else { verdict = 'fail'; fail++ }
    const tolText = tol ? `tol ${fmt(tol.meanAbs)}/${fmt(tol.rowMax)}/${tol.strong}` : 'tol none'
    console.log(`${c.id}: ${verdict} ${fmt(r.diff.meanAbs)} ${fmt(r.diff.rowMax)} strong ${r.diff.strong} row ${r.diff.worstRow} (${tolText}) [${r.ref.width}x${r.ref.height}, gridlines cols ${r.snapped.cols} rows ${r.snapped.rows}]`)
    if (fit) rows.push(`| ${c.id} | ${entry?.unverifiedReason !== undefined ? entry.status : 'pass'} | ${fmt(r.diff.meanAbs)} | ${fmt(r.diff.rowMax)} | ${r.diff.strong} | ${fitTolerance(r.diff)} | ${entry?.unverifiedReason ?? ''} |`)
    if (out) { await savePng(r.ref, join(out, `${c.id}.ref.png`)); await savePng(r.ours, join(out, `${c.id}.ours.png`)); await savePng(diffImage(r.ref, r.ours), join(out, `${c.id}.diff.png`)) }
  }
  if (fit) console.log(['', '| Case id | Status | meanAbs | rowMax | strong | Tolerance | Note |', '|---|---|---|---|---|---|---|', ...rows].join('\n'))
  console.log(`total: ${pass} pass, ${fail} fail, ${unverified} unverified`)
  return fail ? 1 : 0
}

if ((import.meta as { main?: boolean }).main) process.exit(await main(process.argv.slice(2)))
