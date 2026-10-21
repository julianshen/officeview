/**
 * Pie-of-pie and bar-of-pie (spec 4.1): the split decision, the deterministic
 * two-region geometry, the slice angles and fractions, the serLines
 * connectors, the label placement and the painter. The painter and the label
 * placement share these exports, so shapes and labels can never disagree.
 */
import { CHART_POINT_LIMIT } from '../limits'
import type { ChartGroup, ChartSeries } from '../model'
import { applyFillStyle, type ChartFillStyle } from '../fills'
import { applyLineStyle, lineHasColor, type ChartLineStyle } from '../style'

export interface Box { x: number; y: number; w: number; h: number }

/** A drawing value: only finite positives draw (Infinity would break every sweep). */
export const ofPieDisplayValue = (raw: number | undefined): number => raw !== undefined && Number.isFinite(raw) && raw > 0 ? raw : 0

export interface OfPieSplit { secondary: boolean[]; primaryTotal: number; secondaryTotal: number }

/**
 * The split decision over `count` slot-ordered points (`valueOf(i)` returns
 * the display value). `splitType` defaults to the Office default `auto`: the
 * last `ceil(n/3)` points move to the secondary plot. `pos` moves the last
 * `splitPos` points (clamped to the count), `val` every point below
 * `splitPos`, `percent` every point below `splitPos` % of the total, and
 * `cust` the authored `secondPiePt` indices. The two regional totals conserve
 * the drawn values: the synthetic Other slice always equals `secondaryTotal`.
 *
 * Blank, error, unavailable and zero points are never moved by `val` or
 * `percent` (the display value of each is 0): they keep their zero sweep in
 * the primary region, as the plain pie draws them. `pos` and `cust` are
 * position-based and move points regardless. Negative values read 0, so they
 * follow the pie family's convention too.
 *
 * The indices are in the caller's display space; `sourceOf(i)` (identity
 * unless given) translates a slot to its source index for the `cust` set,
 * whose indices are source indices as Office writes them.
 */
export function ofPieSplit(group: { splitType?: string; splitPos?: number; secondPiePt?: number[] }, count: number, valueOf: (i: number) => number, sourceOf: (i: number) => number = i => i): OfPieSplit {
  const secondary = new Array<boolean>(count).fill(false)
  let primaryTotal = 0, secondaryTotal = 0
  switch (group.splitType) {
    case 'cust': {
      const wanted = new Set<number>()
      for (const i of group.secondPiePt ?? []) if (Number.isInteger(i) && i >= 0) wanted.add(i)
      for (let i = 0; i < count; i++) if (wanted.has(sourceOf(i))) secondary[i] = true
      break
    }
    case 'pos': {
      const splitPos = Math.min(Math.max(Math.floor(group.splitPos ?? 0), 0), count)
      for (let i = Math.max(0, count - splitPos); i < count; i++) secondary[i] = true
      break
    }
    case 'val': {
      const splitPos = group.splitPos ?? 0
      for (let i = 0; i < count; i++) { const v = valueOf(i); if (v > 0 && v < splitPos) secondary[i] = true }
      break
    }
    case 'percent': {
      const splitPos = group.splitPos ?? 0
      let total = 0
      for (let i = 0; i < count; i++) total += valueOf(i)
      for (let i = 0; i < count; i++) { const v = valueOf(i); if (v > 0 && v < splitPos / 100 * total) secondary[i] = true }
      break
    }
    default: { // auto: the last ceil(n/3) points
      const moved = Math.min(count, Math.ceil(count / 3))
      for (let i = Math.max(0, count - moved); i < count; i++) secondary[i] = true
    }
  }
  for (let i = 0; i < count; i++) { if (secondary[i]) secondaryTotal += valueOf(i); else primaryTotal += valueOf(i) }
  return { secondary, primaryTotal, secondaryTotal }
}

export interface OfPieRegion { cx: number; cy: number; r: number }
export interface OfPieLayout {
  primary: OfPieRegion
  secondary: OfPieRegion
  gap: number
  /** `ofPieType='bar'`: the secondary plot is a stacked bar in this box (segments stacking down from the top). */
  bar?: { x: number; y: number; w: number; h: number }
}

/**
 * The two-region layout on `area`: horizontal margins of 5 on each side, a
 * radius cap of (h - 8) / 2, and one span 2r1 + gap + 2r2 = w - 10 (both
 * regions fit the area at once), so r1 = min((h - 8) / 2, (w - 10) / (2 + 2s + k)).
 * The gap is `gapWidth` % of r2 for both variants (the authored 0-500, whose
 * Office default is 50 for a pie of pie and 150 for a bar of pie).
 * s is `secondPieSize` (percent of r1), clamped 5..200; r2 = s r1.
 */
export function ofPieLayout(area: Box, type: 'pie' | 'bar', secondPieSize: number, gapWidth: number): OfPieLayout {
  const s = Math.min(Math.max(secondPieSize, 5), 200) / 100
  const k = s * Math.min(Math.max(gapWidth, 0), 500) / 100
  const r1 = Math.min((area.h - 8) / 2, (area.w - 10) / (2 + 2 * s + k))
  const gap = k * r1, r2 = s * r1, cy = area.y + area.h / 2
  const primary = { cx: area.x + 5 + r1, cy, r: r1 }
  const cx = primary.cx + r1 + gap + r2
  return type === 'bar'
    ? { primary, secondary: { cx, cy, r: r2 }, gap, bar: { x: cx - r2 / 2, y: cy - r2, w: r2, h: 2 * r2 } }
    : { primary, secondary: { cx, cy, r: r2 }, gap }
}

export interface OfPieSlice { slot: number; source: number; start: number; sweep: number }
export interface OfPieSlices {
  /** Primary slices in order (angles in radians, from the base angle). */
  primary: OfPieSlice[]
  /** The synthetic Other slice, drawn last; `slot` and `source` are -1. Its value is `secondaryTotal`. */
  other: OfPieSlice
  /**
   * Secondary slices in order. For pie-of-pie the angles are radians; for
   * bar-of-pie `start`/`sweep` are height fractions over the bar, stacking
   * down from its top.
   */
  secondary: OfPieSlices['primary']
}

export const ofPieType = (group: OfPieGroup): 'pie' | 'bar' => group.ofPieType === 'bar' ? 'bar' : 'pie'
export const ofPieSecondPieSize = (group: OfPieGroup): number => group.secondPieSize ?? 75

/**
 * The synthetic Other slice is Office's point `n` (index `n` in any point
 * indexing): at most the series' point count capped at the point limit, and
 * at least 1 (an empty series keeps index 1 for it).
 */
export const ofPieOtherIndex = (pointCount: number): number => Math.max(Math.min(pointCount, CHART_POINT_LIMIT), 1)

/**
 * Slice angles for both regions, from one split and layout. The primary pie
 * draws the primary points in order and then the Other slice, over the
 * conserved total `primaryTotal + secondaryTotal`; the secondary region draws
 * its own points over `secondaryTotal` (a full small pie, or the bar's
 * fractions stacking down from the top).
 *
 * Each slice carries the display slot it belongs to and the source index
 * `sourceOf(slot)` maps it to (for dPt, pointColors and cust matching); the
 * Other slice carries -1 for both.
 */
export function ofPieSlices(group: { firstSliceAngle?: number }, values: readonly number[], split: OfPieSplit, layout: OfPieLayout, sourceOf: (i: number) => number = i => i): OfPieSlices {
  const base = ((group.firstSliceAngle ?? 0) - 90) * Math.PI / 180
  const primary: OfPieSlice[] = []
  const secondary: OfPieSlice[] = []
  const total = split.primaryTotal + split.secondaryTotal
  let start = base, secondaryStart = base
  if (layout.bar) secondaryStart = 0 // the bar's fractions have no base angle: the first secondary point sits at the top
  for (let slot = 0; slot < values.length; slot++) {
    const source = sourceOf(slot)
    if (split.secondary[slot]) {
      // A share of the secondary total: radians on the small pie, a height fraction on the bar.
      const share = split.secondaryTotal > 0 ? values[slot] / split.secondaryTotal : 0
      const sweep = layout.bar ? share : share * Math.PI * 2
      secondary.push({ slot, source, start: secondaryStart, sweep })
      secondaryStart += sweep
    } else {
      const sweep = total > 0 ? values[slot] / total * Math.PI * 2 : 0
      primary.push({ slot, source, start, sweep })
      start += sweep
    }
  }
  const other = { slot: -1, source: -1, start, sweep: total > 0 ? split.secondaryTotal / total * Math.PI * 2 : 0 }
  return { primary, other, secondary }
}

/** One serLines connector: from the Other slice's rim to the secondary region's near edge. */
export interface OfPieConnector { x1: number; y1: number; x2: number; y2: number }

/**
 * The two serLines connectors start at the Other slice's first and last rim
 * points and meet the secondary region's left rim (pie) or the bar's left
 * edge (bar) at a quarter and three quarter of the region's height.
 */
export function ofPieConnectors(layout: OfPieLayout, other: OfPieSlice): OfPieConnector[] {
  const r = layout.primary.r
  const from = (angle: number) => ({ x1: layout.primary.cx + r * Math.cos(angle), y1: layout.primary.cy + r * Math.sin(angle) })
  const to = (fraction: number) => layout.bar
    ? { x2: layout.bar.x, y2: layout.bar.y + fraction * layout.bar.h }
    : { x2: layout.secondary.cx - layout.secondary.r, y2: layout.secondary.cy + (fraction - 0.5) * 2 * layout.secondary.r }
  return [{ ...from(other.start), ...to(0.25) }, { ...from(other.start + other.sweep), ...to(0.75) }]
}

export interface OfPiePlan { split: OfPieSplit; layout: OfPieLayout; slices: OfPieSlices; type: 'pie' | 'bar' }

/** The ofPie fields any group must carry; painters see a DrawGroup, labels a plain ChartGroup. */
export type OfPieGroup = Pick<ChartGroup, 'family' | 'gapWidth' | 'firstSliceAngle' | 'ofPieType' | 'splitType' | 'splitPos' | 'secondPieSize' | 'secondPiePt' | 'serLines'>

/** One step from a display-value array (order = the display slots) to everything about the two regions. */
export function ofPiePlan(group: OfPieGroup, values: readonly number[], area: Box, sourceOf: (i: number) => number = i => i): OfPiePlan {
  const type = ofPieType(group)
  const split = ofPieSplit(group, values.length, i => values[i], sourceOf)
  const layout = ofPieLayout(area, type, ofPieSecondPieSize(group), group.gapWidth)
  return { split, layout, slices: ofPieSlices(group, values, split, layout, sourceOf), type }
}

export interface OfPieLabelPoint { cx: number; cy: number; r: number; angle: number; fraction: number; bar?: boolean }

/**
 * Where a point's slice sits (for data-label placement): the centre, radius
 * and mid angle of its region — `fraction` the point's share of its
 * denominator (the conserved primary total including Other on the primary
 * region, the secondary total in the secondary region) for the best-fit rule.
 * On the bar the point is a segment; `fraction` its share of the bar.
 *
 * The point is a display slot; the pseudo point -1 is the synthetic Other
 * slice, placed on the primary region like any primary slice. Nothing to
 * place for unknown slots or a too-small Other slice.
 */
export function ofPiePointPlacement(plan: OfPiePlan, point: number): OfPieLabelPoint | undefined {
  const { split, layout, slices } = plan
  if (point < 0) {
    if (!(slices.other.sweep > 0)) return undefined
    return { cx: layout.primary.cx, cy: layout.primary.cy, r: layout.primary.r, angle: slices.other.start + slices.other.sweep / 2, fraction: slices.other.sweep / (Math.PI * 2) }
  }
  const slice = slices.primary.find(s => s.slot === point) ?? slices.secondary.find(s => s.slot === point)
  if (!slice) return undefined
  const primary = split.secondary[point] !== true
  const fraction = layout.bar && !primary ? slice.sweep : (slice.sweep / (Math.PI * 2))
  const mid = slice.start + slice.sweep / 2
  if (!primary && layout.bar) {
    return { cx: layout.bar.x + layout.bar.w / 2, cy: layout.bar.y + mid * layout.bar.h, r: layout.bar.w / 2, angle: mid, fraction, bar: true }
  }
  const region = primary ? layout.primary : layout.secondary
  return { cx: region.cx, cy: region.cy, r: region.r, angle: mid, fraction }
}

/**
 * Display access for one series, the same precedence the plain pie painter
 * uses: the display slot values (`value(slot)`), the underlying source point
 * (`source(slot)`) and the stylistic fields off the DrawSeries.
 */
export interface OfPieSeries {
  source: ChartSeries
  color: string
  pointColors?: readonly string[]
  value(slot: number): number | undefined
}

/** Fills the current shape with the slice's scheme: an authored fill, else the resolved colour. */
function fillWith(ctx: CanvasRenderingContext2D, fill: ChartFillStyle | undefined, color: string, box: Box): boolean {
  if (fill) return applyFillStyle(ctx, fill, box)
  ctx.fillStyle = `#${color}`
  return true
}

/** Strokes the current shape with an authored border (only one that carries a colour or width), like the plain pie. */
function strokeBorder(ctx: CanvasRenderingContext2D, line: ChartLineStyle | undefined, box: Box, rect?: Box): void {
  if (!line || line.noFill || !(lineHasColor(line) || line.width !== undefined)) return
  ctx.save()
  try {
    ctx.strokeStyle = '#FFFFFF'
    ctx.lineWidth = 0.75 * 96 / 72
    if (applyLineStyle(ctx, line, box)) { if (rect) ctx.strokeRect(rect.x, rect.y, rect.w, rect.h); else ctx.stroke() }
  } finally { ctx.restore() }
}

/** The fill and colour of one slice (a plain pie uses the same precedence): the series' own fill for a colourless series. */
function sliceFill(series: OfPieSeries, sourceIndex: number): { fill: ChartFillStyle | undefined; color: string } {
  const point = series.source.dPt?.get(sourceIndex)
  return {
    fill: point?.fill ?? (series.pointColors?.[sourceIndex] ? undefined : series.source.fill),
    color: series.pointColors?.[sourceIndex] ?? series.color,
  }
}

/** The border of one slice: the dPt override, else the series' own. */
const borderOf = (series: OfPieSeries, sourceIndex: number): ChartLineStyle | undefined => series.source.dPt?.get(sourceIndex)?.line ?? series.source.lineStyle

/**
 * The ofPie painter: the primary pie (primary slices in order, then the Other
 * slice), the secondary region (a small pie, or the stacked bar), and the
 * serLines connectors. Explosion shifts each primary slice along its mid
 * angle like the plain pie; the radius shrinks so the farthest stays inside.
 *
 * Style lookups (dPt, pointColors, line overrides) run at the source index
 * the slot maps to; the Other slice reads Office's point `n` (padding the
 * series' point count) for its fill, border and explosion. The connectors
 * draw only when `<c:serLines>` is present, in the authored style or grey.
 */
export function paintOfPie(ctx: CanvasRenderingContext2D, area: Box, group: OfPieGroup, series: OfPieSeries, sourceOf: (slot: number) => number, slots: number): void {
  const values: number[] = []
  for (let slot = 0; slot < slots; slot++) values.push(ofPieDisplayValue(series.value(slot)))
  const { split, layout, slices } = ofPiePlan(group, values, area, sourceOf)
  if (!((split.primaryTotal + split.secondaryTotal) > 0)) return
  // The synthetic Other point is the series' point `n`; its dPt overrides fill, line and explosion (spec 4.1).
  const otherIndex = ofPieOtherIndex(series.source.values.length)
  const explosionOf = (srcIndex: number) => series.source.dPt?.get(srcIndex)?.explosion ?? series.source.explosion ?? 0
  let farthest = explosionOf(otherIndex) // the Other slice is drawn on the primary region too
  for (const slice of slices.primary) farthest = Math.max(farthest, explosionOf(slice.source))
  const r1 = layout.primary.r / (1 + farthest / 100)
  for (const slice of slices.primary) {
    if (slice.sweep <= 0) continue
    const fill = sliceFill(series, slice.source), border = borderOf(series, slice.source)
    const mid = slice.start + slice.sweep / 2
    const offset = (explosionOf(slice.source) / 100) * r1
    const sx = layout.primary.cx + offset * Math.cos(mid), sy = layout.primary.cy + offset * Math.sin(mid)
    const box = { x: sx - r1, y: sy - r1, w: r1 * 2, h: r1 * 2 }
    ctx.beginPath()
    ctx.moveTo(sx, sy)
    ctx.arc(sx, sy, r1, slice.start, slice.start + slice.sweep)
    ctx.closePath()
    if (fillWith(ctx, fill.fill, fill.color, box)) ctx.fill()
    if (!border) { ctx.strokeStyle = "#FFFFFF"; ctx.lineWidth = 1; ctx.stroke() } else strokeBorder(ctx, border, box)
  }
  if (slices.other.sweep > 0) {
    const offset = (explosionOf(otherIndex) / 100) * r1
    const mid = slices.other.start + slices.other.sweep / 2
    const sx = layout.primary.cx + offset * Math.cos(mid), sy = layout.primary.cy + offset * Math.sin(mid)
    const box = { x: sx - r1, y: sy - r1, w: r1 * 2, h: r1 * 2 }
    ctx.beginPath()
    ctx.moveTo(sx, sy)
    ctx.arc(sx, sy, r1, slices.other.start, slices.other.start + slices.other.sweep)
    ctx.closePath()
    const fill = sliceFill(series, otherIndex), border = borderOf(series, otherIndex)
    if (fillWith(ctx, fill.fill, fill.color, box)) ctx.fill()
    if (!border) { ctx.strokeStyle = "#FFFFFF"; ctx.lineWidth = 1; ctx.stroke() } else strokeBorder(ctx, border, box)
  }
  if (layout.bar) {
    for (const slice of slices.secondary) {
      if (slice.sweep <= 0) continue
      const fill = sliceFill(series, slice.source), border = borderOf(series, slice.source)
      const box = { x: layout.bar.x, y: layout.bar.y + slice.start * layout.bar.h, w: layout.bar.w, h: slice.sweep * layout.bar.h }
      ctx.beginPath()
      if (fillWith(ctx, fill.fill, fill.color, box)) ctx.fillRect(box.x, box.y, box.w, box.h)
      if (!border) { ctx.strokeStyle = "#FFFFFF"; ctx.lineWidth = 1; ctx.strokeRect(box.x, box.y, box.w, box.h) } else strokeBorder(ctx, border, box, box)
    }
  } else {
    let far = 0
    for (const slice of slices.secondary) far = Math.max(far, explosionOf(slice.source))
    const r2 = layout.secondary.r / (1 + far / 100)
    for (const slice of slices.secondary) {
      if (slice.sweep <= 0) continue
      const fill = sliceFill(series, slice.source), border = borderOf(series, slice.source)
      const mid = slice.start + slice.sweep / 2
      const offset = (explosionOf(slice.source) / 100) * r2
      const sx = layout.secondary.cx + offset * Math.cos(mid), sy = layout.secondary.cy + offset * Math.sin(mid)
      const box = { x: sx - r2, y: sy - r2, w: r2 * 2, h: r2 * 2 }
      ctx.beginPath()
      ctx.moveTo(sx, sy)
      ctx.arc(sx, sy, r2, slice.start, slice.start + slice.sweep)
      ctx.closePath()
      if (fillWith(ctx, fill.fill, fill.color, box)) ctx.fill()
      if (!border) { ctx.strokeStyle = "#FFFFFF"; ctx.lineWidth = 1; ctx.stroke() } else strokeBorder(ctx, border, box)
    }
  }
  drawConnectors(ctx, group.serLines, layout, slices.other)
}

/** The serLines connectors: an authored line style, else a 1px grey line each. */
function drawConnectors(ctx: CanvasRenderingContext2D, serLines: ChartLineStyle | undefined, layout: OfPieLayout, other: OfPieSlice): void {
  // Office draws none when the element is absent; noFill suppresses them too.
  if (!serLines || serLines.noFill) return
  const authored = lineHasColor(serLines) || serLines.width !== undefined
  ctx.strokeStyle = '#D9D9D9'
  ctx.lineWidth = 1
  for (const line of ofPieConnectors(layout, other)) {
    ctx.beginPath()
    ctx.moveTo(line.x1, line.y1)
    ctx.lineTo(line.x2, line.y2)
    if (authored) {
      ctx.save()
      try { ctx.strokeStyle = '#595959'; applyLineStyle(ctx, serLines, layout.bar ?? { x: Math.min(line.x1, line.x2), y: Math.min(line.y1, line.y2), w: Math.abs(line.x2 - line.x1) || 1, h: Math.abs(line.y2 - line.y1) || 1 }); ctx.stroke() } finally { ctx.restore() }
    } else ctx.stroke()
  }
}
