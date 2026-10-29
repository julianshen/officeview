/**
 * Slicer and timeline painters (spec §6.6, Task 12). Lazy chunk `pivots`.
 *
 * Geometry is Excel 16.114's, measured at 1 px per pt (docs/validation/xlsx-pivots.md
 * "Slicer geometry", "Timeline geometry") and scaled here by `scale`, the canvas
 * units (worksheet CSS px) per pt — 96/72 in the renderer, as for the pivot
 * buttons. Lines are one CSS px, like cell borders.
 *
 * Rasterisation is a pure function of global device coordinates: every edge is
 * rounded to the device pixel grid of the current transform, and drawing happens
 * in device space, so a control split across tiles paints exactly as on one
 * surface (Review Focus 4). A non-axis-aligned transform draws unsnapped.
 *
 * View-only: no hover state, no header buttons (multi-select, clear filter), and
 * the timeline's labels are English, where Excel draws its UI language.
 */
import { parseThemeContext } from '../../../drawing/style'
import type { Box, SlicerAppearance, SlicerItemState, SlicerModel, TimelineAppearance, TimelineLevel, TimelineModel, XlsxControlContent } from './model'
import { slicerAppearance, timelineAppearance } from './style'

/** Slicer geometry in pt (refs/slicer-light1.png, slicer-two-column, slicer-scrolled, slicer-no-data). */
export const SLICER_PT = {
  /** The header's separator lies 19 below the top border; the first item starts 4 below it (23). */
  header: 19, itemTop: 4,
  /** [estimate] With `showCaption="0"` the first item starts this far below the top border. */
  noHeaderTop: 4,
  captionX: 7, captionBaseline: 13,
  /** Items: 6 in from the frame, 2 apart (rows and columns), text 8 into the item, rounded corners. */
  inset: 6, gap: 2, textX: 8, textRight: 4, radius: 3,
  /** The last row ends at least 6 above the bottom border. */
  bottom: 6,
  /** Scrollbar: items end 20 inside the frame; the 8 wide thumb ends 8 inside it. */
  scrollItemsRight: 20, thumbRight: 8, thumbWidth: 8, thumbRadius: 4,
  fontPt: 11,
} as const
const SLICER_THUMB = '#BCBCBC'

/** Timeline geometry in pt (refs/timeline-light1-months.png, timeline-dark2-days, timeline-light1-years). */
export const TIMELINE_PT = {
  separator: 22, captionX: 6, captionBaseline: 15,
  labelX: 6, selectionBaseline: 38, levelRight: 10, yearBaseline: 56, periodBaseline: 74,
  barTop: 78, barHeight: 13, barLeft: 6, barRight: 7,
  trackGap: 6, trackHeight: 5, thumbHeight: 7,
  /** [estimate] Caption and labels: the captions draw Excel's UI font, here the theme's body font. */
  captionPt: 12, labelPt: 11,
} as const
/** Period width per level, pt [measure]: dividers every 39 (months, 495/534/573/612) and 21 (days, 476/497/518);
 * a quarter 65–67, a year block 38 wide (refs/timeline-light1-*.png). */
export const TIMELINE_PERIOD_PT: Record<TimelineLevel, number> = { years: 38, quarters: 66, months: 39, days: 21 }

export function slicerItemState(item: { selected: boolean; noData: boolean }): SlicerItemState {
  return item.noData ? item.selected ? 'selectedItemWithNoData' : 'unselectedItemWithNoData' : item.selected ? 'selectedItemWithData' : 'unselectedItemWithData'
}

export interface SlicerLayout {
  /** Top of the 1 px header separator; undefined without a caption. */
  separatorY?: number
  caption?: { x: number; baseline: number }
  items: Array<{ index: number; rect: Box; state: SlicerItemState; text: { x: number; baseline: number } }>
  rows: number
  firstRow: number
  visibleRows: number
  thumb?: Box
}

export function slicerLayout(m: SlicerModel, box: Box, scale: number): SlicerLayout {
  const s = scale, P = SLICER_PT
  const top = box.y + (m.showCaption ? P.header + P.itemTop : P.noHeaderTop) * s
  // Items are whole points tall at 1 px per pt (16.5 pt rowHeight → 17), then 2 apart: the measured 19 pitch.
  const itemH = Math.max(1, Math.ceil(m.rowHeightEmu / 12_700 - 1e-9)) * s, pitch = itemH + P.gap * s
  const cols = Math.max(1, Math.min(64, Math.floor(m.columnCount) || 1))
  const rows = Math.ceil(m.items.length / cols)
  const bottom = box.y + box.height - P.bottom * s
  const fit = Math.max(1, Math.floor((bottom - top - itemH) / pitch + 1e-9) + 1)
  const visibleRows = Math.min(rows, fit)
  const scroll = rows > visibleRows
  // Excel clamps startItem so the last row fills the view (slicer-scrolled: startItem 2 of 4 shows items 1–3).
  const firstRow = Math.max(0, Math.min(Math.floor(Math.max(0, m.startItem) / cols), rows - visibleRows))
  const left = box.x + P.inset * s, right = box.x + box.width - (scroll ? P.scrollItemsRight : P.inset) * s
  const colW = Math.max(0, (right - left - (cols - 1) * P.gap * s) / cols)
  const items: SlicerLayout['items'] = []
  for (let r = 0; r < visibleRows; r++) for (let c = 0; c < cols; c++) {
    const index = (firstRow + r) * cols + c
    const item = m.items[index]
    if (!item) continue
    const rect = { x: left + c * (colW + P.gap * s), y: top + r * pitch, width: colW, height: itemH }
    items.push({ index, rect, state: slicerItemState(item), text: { x: rect.x + P.textX * s, baseline: rect.y + itemH / 2 + 2.5 * s } })
  }
  const track = { top, bottom }
  return {
    ...(m.showCaption ? { separatorY: box.y + P.header * s, caption: { x: box.x + P.captionX * s, baseline: box.y + P.captionBaseline * s } } : {}),
    items, rows, firstRow, visibleRows,
    ...(scroll ? { thumb: { x: box.x + box.width - (P.thumbRight + P.thumbWidth) * s, y: track.top + (track.bottom - track.top) * firstRow / rows, width: P.thumbWidth * s, height: (track.bottom - track.top) * visibleRows / rows } } : {}),
  }
}

// ---------------------------------------------------------------- timeline layout

const DAY = 86_400_000
const MONTHS = ['JAN', 'FEB', 'MAR', 'APR', 'MAY', 'JUN', 'JUL', 'AUG', 'SEP', 'OCT', 'NOV', 'DEC']
const LEVEL_LABEL: Record<TimelineLevel, string> = { years: 'YEARS', quarters: 'QUARTERS', months: 'MONTHS', days: 'DAYS' }
const ymd = (t: number) => { const d = new Date(t); return { y: d.getUTCFullYear(), m: d.getUTCMonth(), d: d.getUTCDate() } }
/** The start of the period containing `t`. */
function periodStart(level: TimelineLevel, t: number): number {
  const { y, m, d } = ymd(t)
  return level === 'years' ? Date.UTC(y, 0, 1) : level === 'quarters' ? Date.UTC(y, m - m % 3, 1) : level === 'months' ? Date.UTC(y, m, 1) : Date.UTC(y, m, d)
}
/** Whole periods from `origin` (a period start) to the period containing `t`. */
function periodIndex(level: TimelineLevel, origin: number, t: number): number {
  if (level === 'days') return Math.floor((periodStart('days', t) - origin) / DAY)
  const a = ymd(origin), b = ymd(t), months = (b.y - a.y) * 12 + b.m - a.m
  return level === 'years' ? b.y - a.y : level === 'quarters' ? Math.floor(months / 3) : months
}
function periodAt(level: TimelineLevel, origin: number, index: number): number {
  if (level === 'days') return origin + index * DAY
  const { y, m } = ymd(origin)
  return level === 'years' ? Date.UTC(y + index, 0, 1) : Date.UTC(y, m + index * (level === 'quarters' ? 3 : 1), 1)
}
function periodLabel(level: TimelineLevel, t: number): string {
  const { y, m, d } = ymd(t)
  return level === 'years' ? String(y) : level === 'quarters' ? `Q${Math.floor(m / 3) + 1}` : level === 'months' ? MONTHS[m]! : String(d)
}
/** The label of the row above the periods where it changes: the year, or the month and year for days. */
function groupLabel(level: TimelineLevel, t: number): string | undefined {
  const { y, m } = ymd(t)
  return level === 'years' ? undefined : level === 'days' ? `${MONTHS[m]} ${y}` : String(y)
}
/** A date in the selection label, without its year. */
function formatPeriod(level: TimelineLevel, t: number): string {
  const { y, m, d } = ymd(t)
  return level === 'years' ? String(y) : level === 'quarters' ? `Q${Math.floor(m / 3) + 1}` : level === 'months' ? MONTHS[m]! : `${MONTHS[m]} ${d}`
}
/** Excel's selection label, year first ("2024 3月 - 5月" in the zh-TW capture → "2024 MAR - MAY"). */
function selectionText(level: TimelineLevel, start: number, end: number): string {
  const a = ymd(start).y, b = ymd(end).y, single = periodStart(level, start) === periodStart(level, end)
  if (level === 'years') return single ? String(a) : `${a} - ${b}`
  if (single) return `${a} ${formatPeriod(level, start)}`
  return a === b ? `${a} ${formatPeriod(level, start)} - ${formatPeriod(level, end)}` : `${a} ${formatPeriod(level, start)} - ${b} ${formatPeriod(level, end)}`
}

export interface TimelinePeriod { index: number; start: number; x: number; width: number; selected: boolean; label: string; group?: string }
export interface TimelineLayout {
  separatorY?: number
  caption?: { x: number; baseline: number }
  selectionLabel?: { x: number; baseline: number; text: string }
  level?: { right: number; baseline: number; text: string }
  yearBaseline: number
  periodBaseline: number
  bar: Box
  /** Period width in canvas units. */
  pitch: number
  /** All periods of the range, and the index of the first one in view. */
  total: number
  first: number
  /** The periods in view, left to right. */
  periods: TimelinePeriod[]
  track?: Box
  thumb?: Box
}

export function timelineLayout(m: TimelineModel, box: Box, scale: number): TimelineLayout {
  const s = scale, P = TIMELINE_PT, show = m.show ?? { header: true, selectionLabel: true, timeLevel: true, scrollbar: true }
  const bar = { x: box.x + P.barLeft * s, y: box.y + P.barTop * s, width: Math.max(0, box.width - (P.barLeft + P.barRight) * s), height: P.barHeight * s }
  const pitch = TIMELINE_PERIOD_PT[m.level] * s
  const origin = periodStart(m.level, m.range.start)
  const total = Math.max(1, periodIndex(m.level, origin, m.range.end - 1) + 1)
  const inView = pitch > 0 ? bar.width / pitch : total
  const scrolled = m.scrollPosition === undefined ? 0 : periodIndex(m.level, origin, m.scrollPosition)
  const first = Math.max(0, Math.min(scrolled, total - Math.floor(inView)))
  const count = Math.min(total - first, Math.ceil(inView) + 1)
  const sel = m.selection
  const periods: TimelinePeriod[] = []
  for (let k = 0; k < count; k++) {
    const index = first + k, start = periodAt(m.level, origin, index), end = periodAt(m.level, origin, index + 1)
    const selected = !sel || (start <= sel.end && end > sel.start)
    const previous = k === 0 ? undefined : periodAt(m.level, origin, index - 1)
    const group = groupLabel(m.level, start)
    periods.push({ index, start, x: bar.x + k * pitch, width: pitch, selected, label: periodLabel(m.level, start), ...(group !== undefined && (previous === undefined || groupLabel(m.level, previous) !== group) ? { group } : {}) })
  }
  // No selection handles (Ruling AK): Excel draws none in a static capture, with or without a selection
  // (refs/timeline-*.png, refs/measure-timeline-selection.png).
  const track = show.scrollbar ? { x: bar.x, y: bar.y + bar.height + P.trackGap * s, width: bar.width, height: P.trackHeight * s } : undefined
  const fraction = inView / total
  const thumb = track && fraction < 1 ? { x: bar.x + bar.width * first / total, y: track.y - s, width: Math.max(2 * s, bar.width * fraction), height: P.thumbHeight * s } : undefined
  const label = sel ? selectionText(m.level, sel.start, sel.end) : 'All Periods'
  return {
    ...(show.header ? { separatorY: box.y + P.separator * s, caption: { x: box.x + P.captionX * s, baseline: box.y + P.captionBaseline * s } } : {}),
    ...(show.selectionLabel ? { selectionLabel: { x: box.x + P.labelX * s, baseline: box.y + P.selectionBaseline * s, text: label } } : {}),
    ...(show.timeLevel ? { level: { right: box.x + box.width - P.levelRight * s, baseline: box.y + P.selectionBaseline * s, text: LEVEL_LABEL[m.level] } } : {}),
    yearBaseline: box.y + P.yearBaseline * s,
    periodBaseline: box.y + P.periodBaseline * s,
    bar, pitch, total, first, periods,
    ...(track ? { track } : {}),
    ...(thumb ? { thumb } : {}),
  }
}

// ---------------------------------------------------------------- device-space drawing

interface Device { x(v: number): number; y(v: number): number; k: number; snapped: boolean }
interface Rect { x0: number; y0: number; x1: number; y1: number }
function deviceOf(ctx: CanvasRenderingContext2D): Device {
  const t = typeof ctx.getTransform === 'function' ? ctx.getTransform() : undefined
  if (t && t.b === 0 && t.c === 0 && t.a > 0 && t.d > 0 && Math.abs(t.a - t.d) < 1e-9) return { x: v => Math.round(t.a * v + t.e), y: v => Math.round(t.d * v + t.f), k: t.a, snapped: true }
  return { x: v => v, y: v => v, k: 1, snapped: false }
}
const rectOf = (d: Device, b: Box): Rect => ({ x0: d.x(b.x), y0: d.y(b.y), x1: d.x(b.x + b.width), y1: d.y(b.y + b.height) })
function fill(ctx: CanvasRenderingContext2D, r: Rect, colour: string): void {
  if (r.x1 <= r.x0 || r.y1 <= r.y0) return
  ctx.fillStyle = colour
  ctx.fillRect(r.x0, r.y0, r.x1 - r.x0, r.y1 - r.y0)
}
function roundRect(ctx: CanvasRenderingContext2D, r: Rect, radius: number, colour: string): void {
  const w = r.x1 - r.x0, h = r.y1 - r.y0
  if (w <= 0 || h <= 0) return
  const k = Math.max(0, Math.min(radius, w / 2, h / 2))
  ctx.beginPath()
  ctx.moveTo(r.x0 + k, r.y0)
  ctx.arcTo(r.x1, r.y0, r.x1, r.y1, k)
  ctx.arcTo(r.x1, r.y1, r.x0, r.y1, k)
  ctx.arcTo(r.x0, r.y1, r.x0, r.y0, k)
  ctx.arcTo(r.x0, r.y0, r.x1, r.y0, k)
  ctx.closePath()
  ctx.fillStyle = colour
  ctx.fill()
}
function clipTo(ctx: CanvasRenderingContext2D, r: Rect): void {
  ctx.beginPath()
  ctx.rect(r.x0, r.y0, Math.max(0, r.x1 - r.x0), Math.max(0, r.y1 - r.y0))
  ctx.clip()
}
function frameBorder(ctx: CanvasRenderingContext2D, f: Rect, lw: number, colour: string): void {
  fill(ctx, { ...f, y1: f.y0 + lw }, colour)
  fill(ctx, { ...f, y0: f.y1 - lw }, colour)
  fill(ctx, { ...f, x1: f.x0 + lw }, colour)
  fill(ctx, { ...f, x0: f.x1 - lw }, colour)
}
interface TextStyle { colour: string; bold: boolean; pt: number; family: string; align?: CanvasTextAlign }
function text(ctx: CanvasRenderingContext2D, d: Device, value: string, x: number, baseline: number, style: TextStyle, scale: number, familyCss: (family: string) => string, clip?: Rect): void {
  if (!value) return
  ctx.save()
  try {
    if (clip) clipTo(ctx, clip)
    ctx.font = `${style.bold ? 'bold ' : ''}${style.pt * scale * d.k}px ${familyCss(style.family)}`
    ctx.fillStyle = style.colour
    ctx.textAlign = style.align ?? 'left'
    ctx.textBaseline = 'alphabetic'
    ctx.fillText(value, d.x(x), d.y(baseline))
  } finally { ctx.restore() }
}

const fallbackSlicer = (m: SlicerModel): SlicerAppearance => slicerAppearance(m.style, { dxfs: [], tableStyles: new Map(), slicerStyles: new Map() }, parseThemeContext()).appearance
const fallbackTimeline = (m: TimelineModel): TimelineAppearance => timelineAppearance(m.style, { dxfs: [], tableStyles: new Map(), slicerStyles: new Map() }, parseThemeContext()).appearance

export function paintSlicer(ctx: CanvasRenderingContext2D, m: SlicerModel, box: Box, scale: number, familyCss: (family: string) => string = JSON.stringify): void {
  if (![box.x, box.y, box.width, box.height, scale].every(Number.isFinite) || box.width <= 0 || box.height <= 0 || scale <= 0) return
  const a = m.appearance ?? fallbackSlicer(m), layout = slicerLayout(m, box, scale), P = SLICER_PT
  const d = deviceOf(ctx)
  ctx.save()
  try {
    if (d.snapped) ctx.setTransform(1, 0, 0, 1, 0, 0)
    const f = rectOf(d, box), lw = Math.max(1, Math.round(d.k))
    if (f.x1 <= f.x0 || f.y1 <= f.y0) return
    clipTo(ctx, f)
    fill(ctx, f, a.background)
    if (layout.separatorY !== undefined && layout.caption) {
      const sep = d.y(layout.separatorY)
      fill(ctx, { ...f, y1: sep }, a.header.fill)
      fill(ctx, { ...f, y0: sep, y1: sep + lw }, a.header.separator)
      text(ctx, d, m.caption ?? m.name, layout.caption.x, layout.caption.baseline, { colour: a.header.text, bold: a.header.bold, pt: P.fontPt, family: a.fontFamily }, scale, familyCss, { ...f, y1: sep })
    }
    const radius = P.radius * scale * d.k
    for (const item of layout.items) {
      const paint = a.items[item.state], r = rectOf(d, item.rect)
      roundRect(ctx, r, radius, paint.border)
      roundRect(ctx, { x0: r.x0 + lw, y0: r.y0 + lw, x1: r.x1 - lw, y1: r.y1 - lw }, Math.max(0, radius - lw), paint.fill)
      const clip = { x0: r.x0 + lw, y0: r.y0 + lw, x1: d.x(item.rect.x + item.rect.width - P.textRight * scale), y1: r.y1 - lw }
      text(ctx, d, m.items[item.index]!.label, item.text.x, item.text.baseline, { colour: paint.text, bold: paint.bold, pt: P.fontPt, family: a.fontFamily }, scale, familyCss, clip)
    }
    if (layout.thumb) roundRect(ctx, rectOf(d, layout.thumb), P.thumbRadius * scale * d.k, SLICER_THUMB)
    frameBorder(ctx, f, lw, a.border)
  } finally { ctx.restore() }
}

export function paintTimeline(ctx: CanvasRenderingContext2D, m: TimelineModel, box: Box, scale: number, familyCss: (family: string) => string = JSON.stringify): void {
  if (![box.x, box.y, box.width, box.height, scale, m.range.start, m.range.end].every(Number.isFinite) || box.width <= 0 || box.height <= 0 || scale <= 0) return
  const a = m.appearance ?? fallbackTimeline(m), layout = timelineLayout(m, box, scale), P = TIMELINE_PT
  const d = deviceOf(ctx)
  ctx.save()
  try {
    if (d.snapped) ctx.setTransform(1, 0, 0, 1, 0, 0)
    const f = rectOf(d, box), lw = Math.max(1, Math.round(d.k))
    if (f.x1 <= f.x0 || f.y1 <= f.y0) return
    clipTo(ctx, f)
    fill(ctx, f, a.background)
    const label = (value: string, x: number, baseline: number, colour: string, pt: number = P.labelPt, bold = false, align?: CanvasTextAlign) =>
      text(ctx, d, value, x, baseline, { colour, bold, pt, family: a.fontFamily, ...(align ? { align } : {}) }, scale, familyCss, f)
    if (layout.separatorY !== undefined && layout.caption) {
      const sep = d.y(layout.separatorY)
      fill(ctx, { ...f, y0: sep, y1: sep + lw }, a.header.separator)
      label(m.caption ?? m.name, layout.caption.x, layout.caption.baseline, a.header.text, P.captionPt, a.header.bold)
    }
    if (layout.selectionLabel) label(layout.selectionLabel.text, layout.selectionLabel.x, layout.selectionLabel.baseline, a.selectionLabel)
    if (layout.level) {
      // The level picker: its name and a small down arrow, right-aligned 10 pt inside the frame.
      const arrow = 7 * scale, gap = 3 * scale
      label(layout.level.text, layout.level.right - arrow - gap, layout.level.baseline, a.timeLevel, P.labelPt, false, 'right')
      const ax = d.x(layout.level.right - arrow), ay = d.y(layout.level.baseline - 5 * scale), aw = d.x(layout.level.right) - ax, ah = Math.max(1, Math.round(4 * scale * d.k))
      ctx.beginPath(); ctx.moveTo(ax, ay); ctx.lineTo(ax + aw, ay); ctx.lineTo(ax + aw / 2, ay + ah); ctx.closePath()
      ctx.fillStyle = a.timeLevel; ctx.fill()
    }
    // Year text is #989898, month/quarter/day text #909090 (refs/timeline-light1-years vs -months).
    const periodColour = m.level === 'years' ? a.yearLabel : a.periodLabel
    for (const period of layout.periods) {
      if (period.group !== undefined) label(period.group, period.x + scale, layout.yearBaseline, a.yearLabel)
      label(period.label, period.x + 2 * scale, layout.periodBaseline, periodColour)
    }
    // The bar: selected periods flat [approx] (Excel: a vertical gradient), unselected #DADADA. The 1 px divider
    // closing a period is accent tint 0.4 after a selected one and background after an unselected one
    // (refs/measure-timeline-selection.png: white at 495/533, #95B3D7 at 573/612).
    const bar = rectOf(d, layout.bar)
    ctx.save()
    try {
      clipTo(ctx, bar)
      for (const period of layout.periods) {
        const r = rectOf(d, { x: period.x, y: layout.bar.y, width: period.width, height: layout.bar.height })
        fill(ctx, r, period.selected ? a.selected : a.unselected.fill)
      }
      for (const period of layout.periods.slice(0, -1)) {
        const x = d.x(period.x + period.width)
        fill(ctx, { x0: x, y0: bar.y0, x1: x + lw, y1: bar.y1 }, period.selected ? a.divider : a.background)
      }
    } finally { ctx.restore() }
    if (layout.track) fill(ctx, rectOf(d, layout.track), a.track)
    if (layout.thumb) {
      const t = rectOf(d, layout.thumb)
      fill(ctx, t, a.thumbEdge)
      fill(ctx, { x0: t.x0, y0: t.y0 + lw, x1: t.x1, y1: t.y1 - lw }, a.thumb)
    }
    frameBorder(ctx, f, lw, a.border)
  } finally { ctx.restore() }
}

/** The drawing-content entry: a control painted into its frame's local box (worksheet CSS px). */
export function paintControl(ctx: CanvasRenderingContext2D, content: XlsxControlContent, width: number, height: number, familyCss?: (family: string) => string): void {
  const box = { x: 0, y: 0, width, height }, scale = 96 / 72
  if (content.kind === 'slicer') paintSlicer(ctx, content.slicer, box, scale, familyCss)
  else paintTimeline(ctx, content.timeline, box, scale, familyCss)
}
