import { barGeometry } from './barGeometry'
import { applyFillStyle, type ChartFillStyle } from './fills'
import { MARKER_AUTO_CYCLE, paintMarker } from './markers'
import { applyLineStyle, lineHasColor, type ChartLineStyle } from './style'
/**
 * Chart painter (lazy chunk): the common chart types (bar/column
 * clustered/stacked/100%, line, area, pie, doughnut, scatter). Anything else
 * paints the eager placeholder card. Moved from `src/xlsx/chart.ts` (A0a);
 * since A1a the cartesian families draw on the axis graph (`charts/axes.ts`).
 */
import { paintChartPlaceholder } from '../drawing/content-paint'
import { displayView, type DisplayView } from './display'
import type { ChartGroup, ChartModel, ChartSeries } from './model'
import { axisFraction, categoryFraction, type AxisGraph, type GroupAxes, type ResolvedAxis } from './axes'
import { paintDataLabels } from './labelPaint'
import { plotPoint } from './plotPoint'
import { ChartRefusal } from './own'
import { createChartWorkspace, type ChartParentLedger, type ChartProfile, type ChartWorkspace } from './workspace'
import { documentChartContext, type ChartDocumentContext } from './context'
import { dateUnitIndex } from './dates'
import { layoutChart, legendEntries, labelPlace, axisSide, ownedChartWorkspace, reserveChartLayoutWorkspace, invalidateChartFonts, holdChartLayout, paintWorkspaceBytes, releaseWorkspaceLayouts, type ChartLayout, textAngle } from './layout'

type FamilyCss = (family: string) => string
/**
 * What the painter draws: the display view (plotVisOnly per role, points
 * capped), read slot by slot from the compact storage; no per-point array is
 * copied. Every non-value state (blank, error, unavailable) is a gap here, as
 * before D0; the per-family point matrix (spec 5.8) is A2's.
 */
interface DrawSeries extends Pick<ChartSeries, 'name' | 'color' | 'pointColors' | 'line' | 'markers'> {
  /** The model series drawn (its formatting: fill, line, marker, dPt overrides) and its position in its group. */
  source: ChartSeries
  order: number
  value(slot: number): number | undefined
  x(slot: number): number | undefined
  hasX: boolean
  plot(slot:number, x?:boolean): number|'gap'|'span'
}
interface DrawGroup extends Omit<ChartGroup, 'series'> { series: DrawSeries[] }
interface DrawModel extends Omit<ChartModel, 'groups'> { groups: DrawGroup[]; view: DisplayView; source: ChartModel }
function drawModel(model: ChartModel, view = displayView(model)): DrawModel {
  return {
    ...model,
    view,
    source: model,
    groups: model.groups.map((group, g) => ({ ...group, series: view.groupSeries(g).map(({ series: s, value, x, hasX, point }): DrawSeries => ({ source: s, order: Math.max(0, group.series.indexOf(s)), name: s.name, color: s.color, pointColors: s.pointColors, line: s.line, markers: s.markers, value, x, hasX, plot:(slot,x)=>{ const p=point(slot,x);return p ? plotPoint(p,group.family,group.grouping,model.dispBlanksAs,model.dispNaAsBlank) : 'gap' } })) })),
  }
}
const font = (sizePx: number, family: string, familyCss: FamilyCss, bold = false) => `${bold ? 'bold ' : ''}${sizePx}px ${familyCss(family)}`
const TEXT = '#595959', GRID = '#D9D9D9'

interface Box { x: number; y: number; w: number; h: number }
/** Points to CSS px (marker sizes are points). */
const PT = 96 / 72

/** Fills `box`'s path or rect with an authored fill; false for noFill (nothing painted). */
function fillWith(ctx: CanvasRenderingContext2D, fill: ChartFillStyle | undefined, color: string, box: Box): boolean {
  // An authored fill keeps its alpha (applyFillStyle); an unauthored colour is opaque.
  if (fill) return applyFillStyle(ctx, fill, box)
  ctx.fillStyle = `#${color}`
  return true
}
/** Strokes the current path with an authored border (only one that carries a colour or width). */
function strokeBorder(ctx: CanvasRenderingContext2D, line: ChartLineStyle | undefined, color: string, box: Box, rect?: Box): void {
  if (!line || line.noFill || !(lineHasColor(line) || line.width !== undefined)) return
  ctx.save()
  try {
    ctx.strokeStyle = `#${color}`; ctx.lineWidth = PT * 0.75
    if (applyLineStyle(ctx, line, box)) { if (rect) ctx.strokeRect(rect.x, rect.y, rect.w, rect.h); else ctx.stroke() }
  } finally { ctx.restore() }
}
/**
 * One marker (spec 5.4). A series without an authored marker keeps the engine's default (a filled circle of
 * `legacyRadius` px, as before A4a); an authored one takes its symbol (`auto`: the series cycle), its size in
 * points and its spPr fill and border, a dPt marker overriding the series' field by field.
 */
function drawMarker(ctx: CanvasRenderingContext2D, s: DrawSeries, sourceIndex: number, x: number, y: number, legacyRadius: number): void {
  const override = s.source.dPt?.get(sourceIndex)?.marker
  const m = override ? { ...s.source.markerStyle, ...override } : s.source.markerStyle
  if (!m) { ctx.fillStyle = `#${s.color}`; ctx.beginPath(); ctx.arc(x, y, legacyRadius, 0, Math.PI * 2); ctx.fill(); return }
  const symbol = m.symbol === 'auto' ? MARKER_AUTO_CYCLE[s.order % MARKER_AUTO_CYCLE.length] : m.symbol ?? 'circle'
  if (symbol === 'none') return
  const size = m.size !== undefined ? m.size * PT : legacyRadius * 2
  const box = { x: x - size / 2, y: y - size / 2, w: size, h: size }
  ctx.save()
  try {
    ctx.strokeStyle = `#${s.color}`; ctx.lineWidth = PT * 0.75
    const filled = m.fill ? fillWith(ctx, m.fill, s.color, box) : (ctx.fillStyle = `#${s.color}`, true)
    const stroked = Boolean(m.line && !m.line.noFill && (lineHasColor(m.line) || m.line.width !== undefined)) && applyLineStyle(ctx, m.line, box)
    paintMarker(ctx, x, y, symbol, size, filled, stroked)
  } finally { ctx.restore() }
}

function paintLegend(ctx: CanvasRenderingContext2D, model: DrawModel, area: Box, fontPx: number, family: string, familyCss: FamilyCss, layout: ChartLayout): Box {
  const entries = legendEntries(model.source, model.view)
  if (!model.legend || !entries.length) return area
  if (model.legendFill && fillWith(ctx, model.legendFill, 'FFFFFF', area)) ctx.fillRect(area.x, area.y, area.w, area.h)
  strokeBorder(ctx, model.legendLine, '000000', area, area)
  ctx.font = font(fontPx, family, familyCss)
  const sw = fontPx * 0.7, gap = 6
  const widths = entries.map(e => layout.width(e.name, font(fontPx, family, familyCss)) + sw + gap * 2)
  const line = fontPx + 6
  const vertical = model.legend === 'r' || model.legend === 'l' || model.legend === 'tr'
  let x = area.x, y = area.y
  ctx.textBaseline = 'middle'; ctx.textAlign = 'left'
  entries.forEach((entry, i) => {
    if (vertical ? y + line > area.y + area.h + .5 : x + widths[i] > area.x + area.w + .5) return
    ctx.fillStyle = `#${entry.color}`; ctx.fillRect(x + gap, y + (line - sw) / 2, sw, sw)
    ctx.fillStyle = TEXT
    if (vertical) ctx.fillText(entry.name, x + gap * 2 + sw, y + line / 2, Math.max(1, area.w - sw - gap * 3))
    else ctx.fillText(entry.name, x + gap * 2 + sw, y + line / 2)
    if (vertical) y += line; else x += widths[i]
  })
  return area
}

function paintPie(ctx: CanvasRenderingContext2D, model: DrawModel, group: DrawGroup, box: Box): void {
  const series = group.series[0]
  if (!series) return
  const src = series.source
  const values: number[] = []
  for (let slot = 0; slot < model.view.slots; slot++) { const v = series.value(slot); values.push(v !== undefined && v > 0 ? v : 0) }
  const total = values.reduce((a, b) => a + b, 0)
  if (!(total > 0)) return
  // Explosion (percent of the radius, series then dPt): the radius shrinks so the farthest slice stays in the box.
  const explosionOf = (i: number) => src.dPt?.get(model.view.source(i))?.explosion ?? src.explosion ?? 0
  let farthest = 0
  values.forEach((value, i) => { if (value) farthest = Math.max(farthest, explosionOf(i)) })
  const full = Math.max(1, Math.min(box.w, box.h) / 2 - 4), r = full / (1 + farthest / 100), cx = box.x + box.w / 2, cy = box.y + box.h / 2
  const hole = group.family === 'doughnut' ? r * (group.holeSize ?? 50) / 100 : 0
  let angle = ((group.firstSliceAngle ?? 0) - 90) * Math.PI / 180
  values.forEach((value, i) => {
    if (!value) return
    const sweep = value / total * Math.PI * 2, srcIdx = model.view.source(i), point = src.dPt?.get(srcIdx)
    const offset = explosionOf(i) / 100 * r, mid = angle + sweep / 2
    const sx = cx + offset * Math.cos(mid), sy = cy + offset * Math.sin(mid)
    ctx.beginPath()
    if (farthest && hole) { ctx.arc(sx, sy, r, angle, angle + sweep); ctx.arc(sx, sy, hole, angle + sweep, angle, true) }
    else { ctx.moveTo(sx, sy); ctx.arc(sx, sy, r, angle, angle + sweep) }
    ctx.closePath()
    const sliceBox = { x: sx - r, y: sy - r, w: r * 2, h: r * 2 }
    // dPt fill, else the varyColors (or dPt) colour, else the series fill and colour.
    const fill = point?.fill ?? (series.pointColors?.[srcIdx] ? undefined : src.fill)
    if (fillWith(ctx, fill, series.pointColors?.[srcIdx] ?? series.color, sliceBox)) ctx.fill()
    const border = point?.line ?? src.lineStyle
    if (border) strokeBorder(ctx, border, 'FFFFFF', sliceBox)
    else { ctx.strokeStyle = '#FFFFFF'; ctx.lineWidth = 1; ctx.stroke() }
    angle += sweep
  })
  // Unexploded doughnuts keep the engine's white hole; exploded rings are drawn as ring segments (transparent hole).
  if (hole && !farthest) {
    ctx.beginPath()
    ctx.arc(cx, cy, hole, 0, Math.PI * 2)
    ctx.fillStyle = '#FFFFFF'
    ctx.fill()
  }
}

/**
 * The cartesian families (bar, line, area, scatter) drawn on the axis graph (spec 5.1): every group against its
 * own bound axes, each value axis with its own scale (primary and secondary), crossings, orientation, display
 * units, gridlines and authored tick marks. With one value axis and no authored placement it draws exactly as
 * the engine did before A1a (the same layout arithmetic and drawing order).
 */
function paintContour(ctx: CanvasRenderingContext2D, model: DrawModel, group: DrawGroup, box: Box): void {
  const rows = group.series.length
  const cols = Math.max(1, model.categories.length || Math.max(...group.series.map(s => s.source.values.length), 0))
  if (!rows || !cols) return
  // Grid cap 256x256.
  const rCount = Math.min(rows, 256), cCount = Math.min(cols, 256)
  // Value axis major unit for banding.
  const axisId = group.binding.kind === 'surface' ? group.binding.value : undefined
  const axis = axisId ? model.axes.find(a => a.id === axisId) : model.axes.find(a => a.kind === 'value')
  const major = (axis && (axis as any).majorUnit) || 1
  // Collect z values.
  const grid: Array<Array<number | undefined>> = []
  let min = Infinity, max = -Infinity
  for (let r = 0; r < rCount; r++) {
    const row: Array<number | undefined> = []
    for (let c = 0; c < cCount; c++) {
      const v = (group.series[r]?.source.values as any)?.[c]
      row.push(v)
      if (v !== undefined && Number.isFinite(v)) { min = Math.min(min, v); max = Math.max(max, v) }
    }
    grid.push(row)
  }
  if (!Number.isFinite(min)) { min = 0; max = 1 }
  if (min === max) max = min + (major || 1)
  const bandCount = Math.max(1, Math.ceil((max - min) / (major || 1)))
  const cellW = box.w / cCount, cellH = box.h / rCount
  // wireframe?
  const wireframe = group.wireframe === true
  // Band colours: accent cycle fallback.
  const accents = ['4472C4','ED7D31','A5A5A5','FFC000','5B9BD5','70AD47']
  const bandColor = (band: number) => (group.bandColors && group.bandColors[band]) ? group.bandColors[band]! : accents[band % accents.length]
  for (let r = 0; r < rCount; r++) {
    for (let c = 0; c < cCount; c++) {
      const z = grid[r][c]
      if (z === undefined || !Number.isFinite(z)) continue
      const band = Math.max(0, Math.min(bandCount - 1, Math.floor((z - min) / (major || 1))))
      const x = box.x + c * cellW, y = box.y + r * cellH
      ctx.fillStyle = '#' + bandColor(band)
      ctx.fillRect(x, y, cellW, cellH)
      if (wireframe) {
        ctx.strokeStyle = '#000000'; ctx.lineWidth = 1; ctx.strokeRect(x + 0.5, y + 0.5, cellW - 1, cellH - 1)
      }
    }
  }
  if (wireframe) {
    // boundaries only already drawn per cell; ensure outer border
    ctx.strokeStyle = '#000000'; ctx.lineWidth = 1; ctx.strokeRect(box.x + 0.5, box.y + 0.5, box.w - 1, box.h - 1)
  }
}

function paintCartesian(ctx: CanvasRenderingContext2D, model: DrawModel, graph: AxisGraph, box: Box, fontPx: number, family: string, familyCss: FamilyCss, layout: ChartLayout): void {
  const placed = graph.groups.filter((g): g is GroupAxes => Boolean(g))
  if (!placed.length) return
  const count = model.view.slots
  ctx.font = font(fontPx, family, familyCss)
  const used = [...graph.axes.values()].filter(a => placed.some(g => g.horizontal === a || g.vertical === a))
  let remaining = 262144
  const limited = () => { if (!layout.diagnostics.some(d => d.code === 'chart-path-limit')) layout.diagnostics.push({ code: 'chart-path-limit', reason: 'Later chart elements omitted at the 262,144 path-segment work bound' }) }
  const admitPath = (cost: number) => { if (cost > remaining) { limited(); return false }; remaining -= cost; return true }
  const drawableAxes = new Set(used.filter(a => admitPath(16 + 4 * (a.ticks.length + a.minorTicks.length + Math.min(1000, a.slots)))))
  const valueAxes = used.filter(a => placed.some(g => g.value === a))
  const baseAxes = used.filter(a => !valueAxes.includes(a))
  /** Category labels sit between ticks unless the plot authored crossBetween=midCat (legacy: band centres). */
  const labelsBetween = (a: ResolvedAxis) => { const g = placed.find(p => p.category === a); return !(g && (g.value.axis.crossBetween ?? a.axis.crossBetween) === 'midCat') }
  const shown = (a: ResolvedAxis) => !a.deleted && a.axis.tickLblPos !== 'none'
  const sideOf = (a: ResolvedAxis) => axisSide(a, graph)
  const dateLabels = (a: ResolvedAxis) => layout.texts.get(a.id)!
  const labelBand = (a: ResolvedAxis) => layout.axisBands.get(a.id) ?? 0
  const plot = layout.plot
  if (plot.w <= 4 || plot.h <= 4) return
  /** A coordinate along an axis's direction (x for horizontal axes, y for vertical ones) at fraction f. */
  const along = (a: ResolvedAxis, f: number) => a.orientation === 'h' ? plot.x + f * plot.w : plot.y + plot.h - f * plot.h
  /** The coordinate of an axis's line across the plot (its crossing on the axis it crosses). */
  const across = (a: ResolvedAxis) => a.orientation === 'h' ? plot.y + plot.h - a.crossesFraction * plot.h : plot.x + a.crossesFraction * plot.w
  const labelAnchor = (a: ResolvedAxis): number => {
    const place = labelPlace(a, graph)
    if (a.orientation === 'h') return place === 'near' ? plot.y + plot.h : place === 'far' ? plot.y : across(a)
    return place === 'near' ? plot.x : place === 'far' ? plot.x + plot.w : across(a)
  }
  const label = (a: ResolvedAxis, index: number, text: string, x: number, y: number, maxWidth?: number) => {
    if (!layout.labelIndices.get(a.id)?.has(index)) return
    const angle = textAngle(a.axis.text ?? model.text)
    ctx.font = font((a.axis.text?.sizePt ?? model.fontSizePt) * 96 / 72, family, familyCss)
    if (!angle) { if (maxWidth === undefined) ctx.fillText(text, x, y); else ctx.fillText(text, x, y, maxWidth) }
    else {
      const box = layout.labels.get(a.id)!.find(l => l.index === index)!
      if (a.orientation === 'h') y += (ctx.textBaseline === 'bottom' ? -1 : 1) * box.crossExtent / 2
      else x += (ctx.textAlign === 'right' ? -1 : 1) * box.crossExtent / 2
      ctx.save(); try { ctx.textAlign = 'center'; ctx.textBaseline = 'middle'; ctx.translate(x, y); ctx.rotate(angle); ctx.fillText(text, 0, 0) } finally { ctx.restore() }
    }
    ctx.font = font(fontPx, family, familyCss)
  }
  const tickMarks = (a: ResolvedAxis, values: readonly number[], position: (v: number) => number, mark: string | undefined) => {
    if (!mark || mark === 'none' || a.deleted) return
    const line = across(a), side = sideOf(a), outward = side === 'l' || side === 'b' ? 1 : -1
    // Outward and inward extents of the mark from the axis line.
    const from = mark === 'in' ? 0 : 4, to = mark === 'out' ? 0 : -4
    ctx.strokeStyle = '#BFBFBF'
    ctx.beginPath()
    for (const v of values) {
      const p = Math.round(position(v)) + 0.5
      if (a.orientation === 'h') { ctx.moveTo(p, line + from * outward); ctx.lineTo(p, line + to * outward) }
      else { ctx.moveTo(line - from * outward, p); ctx.lineTo(line - to * outward, p) }
    }
    ctx.stroke()
  }
  const gridline = (a: ResolvedAxis, f: number) => {
    ctx.beginPath()
    if (a.orientation === 'h') { const gx = Math.round(plot.x + f * plot.w) + 0.5; ctx.moveTo(gx, plot.y); ctx.lineTo(gx, plot.y + plot.h) }
    else { const gy = Math.round(plot.y + plot.h - f * plot.h) + 0.5; ctx.moveTo(plot.x, gy); ctx.lineTo(plot.x + plot.w, gy) }
    ctx.stroke()
  }

  // The plot-area fill and border (spec 5.4) sit under the gridlines and axes.
  if (model.plotAreaFill && fillWith(ctx, model.plotAreaFill, 'FFFFFF', plot)) ctx.fillRect(plot.x, plot.y, plot.w, plot.h)
  strokeBorder(ctx, model.plotAreaLine, '000000', plot, plot)
  // Value axes: gridlines (and the scale's lower bound) and tick labels.
  ctx.lineWidth = 1
  for (const a of valueAxes) {
    if (!drawableAxes.has(a)) continue
    if (a.axis.minorGridlines) { ctx.strokeStyle = '#F2F2F2'; for (const t of a.minorTicks) gridline(a, axisFraction(a, t)) }
    for (const [i, t] of a.ticks.entries()) {
      const f = axisFraction(a, t)
      ctx.strokeStyle = GRID
      if (a.axis.gridlines || t === a.min) gridline(a, f)
      if (!shown(a)) continue
      ctx.fillStyle = TEXT
      const anchor = labelAnchor(a), right = sideOf(a) === 'r', top = sideOf(a) === 't'
      // An authored outward tick mark pushes the labels clear of it.
      const pad = a.axis.majorTickMark === 'out' || a.axis.majorTickMark === 'cross' ? 4 : 0
      if (a.orientation === 'h') { ctx.textAlign = 'center'; ctx.textBaseline = top ? 'bottom' : 'top'; label(a, i, layout.texts.get(a.id)![i], plot.x + f * plot.w, top ? anchor - 3 - pad : anchor + 3 + pad) }
      else { ctx.textAlign = right ? 'left' : 'right'; ctx.textBaseline = 'middle'; label(a, i, layout.texts.get(a.id)![i], right ? anchor + 4 + pad : anchor - 4 - pad, plot.y + plot.h - f * plot.h) }
    }
    tickMarks(a, a.minorTicks, t => along(a, axisFraction(a, t)), a.axis.minorTickMark)
    tickMarks(a, a.ticks, t => along(a, axisFraction(a, t)), a.axis.majorTickMark)
  }
  // Category, date and x axes: the axis line at its crossing, then labels.
  for (const a of baseAxes) {
    if (!drawableAxes.has(a)) continue
    if (a.role === 'value') {
      // An x value axis (scatter): minor gridlines, then labels and gridlines per tick as before, then tick marks.
      const xAt = (t: number) => plot.x + axisFraction(a, t, false) * plot.w
      if (a.axis.minorGridlines) { ctx.strokeStyle = '#F2F2F2'; for (const t of a.minorTicks) gridline(a, axisFraction(a, t, false)) }
      ctx.textAlign = 'center'; ctx.textBaseline = 'top'
      for (const [i, t] of a.ticks.entries()) {
        const px = plot.x + axisFraction(a, t, false) * plot.w
        if (shown(a)) { ctx.fillStyle = TEXT; label(a, i, layout.texts.get(a.id)![i], px, labelAnchor(a) + 3) }
        if (a.axis.gridlines) { ctx.strokeStyle = GRID; ctx.beginPath(); ctx.moveTo(Math.round(px) + 0.5, plot.y); ctx.lineTo(Math.round(px) + 0.5, plot.y + plot.h); ctx.stroke() }
      }
      tickMarks(a, a.minorTicks, xAt, a.axis.minorTickMark)
      tickMarks(a, a.ticks, xAt, a.axis.majorTickMark)
      continue
    }
    const n = a.slots
    const between = labelsBetween(a)
    const len = a.orientation === 'h' ? plot.w : plot.h
    const band = len / n
    if (!a.deleted) {
      ctx.strokeStyle = '#BFBFBF'
      ctx.beginPath()
      const line = Math.round(across(a)) + 0.5
      if (a.orientation === 'h') { ctx.moveTo(plot.x, line); ctx.lineTo(plot.x + plot.w, line) }
      else { ctx.moveTo(line, plot.y); ctx.lineTo(line, plot.y + plot.h) }
      ctx.stroke()
    }
    // Date axes: gridlines and tick marks at the calendar ticks (major and minor), drawn whether or not labels are.
    const dateAt = (t: number) => categoryFraction(a, dateUnitIndex(t, a.dateBase!, graph.epoch) - a.dateStart!, between)
    if (a.role === 'date') {
      if (a.axis.minorGridlines) { ctx.strokeStyle = '#F2F2F2'; for (const t of a.minorTicks) gridline(a, dateAt(t)) }
      if (a.axis.gridlines) { ctx.strokeStyle = GRID; for (const t of a.ticks) gridline(a, dateAt(t)) }
      tickMarks(a, a.minorTicks, t => along(a, dateAt(t)), a.axis.minorTickMark)
      tickMarks(a, a.ticks, t => along(a, dateAt(t)), a.axis.majorTickMark)
    }
    if (a.role === 'category') {
      const authored = a.axis.tickMarkSkip ?? 1, skip = Math.max(authored, Math.ceil((between ? n + 1 : n) / 1000))
      const marks = Array.from({ length: Math.ceil((between ? n + 1 : n) / skip) }, (_, k) => k * skip)
      tickMarks(a, marks, k => along(a, between ? (a.reversed ? 1 - k / n : k / n) : categoryFraction(a, k, false)), a.axis.majorTickMark)
    }
    if (!shown(a)) continue
    const anchor = labelAnchor(a), place = sideOf(a)
    ctx.fillStyle = TEXT
    if (a.role === 'date') {
      // Date axes: labels at the major ticks (calendar steps), thinned in screen order (whatever the axis direction)
      // so they never overlap: a label is kept when its extent clears the last kept one.
      const texts = dateLabels(a)
      const items = a.ticks.map((t, k) => ({ k, p: along(a, dateAt(t)) }))
      items.sort((p, q) => p.p - q.p)
      for (const { k, p } of items) {
        if (!layout.labelIndices.get(a.id)?.has(k)) continue
        if (a.orientation === 'h') { ctx.textAlign = 'center'; ctx.textBaseline = place === 't' ? 'bottom' : 'top'; label(a, k, texts[k], p, place === 't' ? anchor - 3 : anchor + 3) }
        else { ctx.textAlign = place === 'r' ? 'left' : 'right'; ctx.textBaseline = 'middle'; label(a, k, texts[k], place === 'r' ? anchor + 4 : anchor - 4, p) }
      }
      continue
    }
    if (model.view.category(0) === undefined) continue
    // Legacy positions (band centres) when unreversed and between; otherwise the axis fraction.
    const at = (i: number) => between && !a.reversed ? (a.orientation === 'h' ? plot.x + band * (i + 0.5) : plot.y + plot.h - band * (i + 0.5)) : along(a, categoryFraction(a, i, between))
    for (let i = 0; i < count; i++) {
      const c = model.view.category(i)
      if (c === undefined) continue
      if (a.orientation === 'v') { ctx.textAlign = place === 'r' ? 'left' : 'right'; ctx.textBaseline = 'middle'; label(a, i, c, place === 'r' ? anchor + 4 : anchor - 4, at(i)) }
      else { ctx.textAlign = 'center'; ctx.textBaseline = place === 't' ? 'bottom' : 'top'; label(a, i, c, at(i), place === 't' ? anchor - 3 : anchor + 3) }
    }

  }
  // Axis titles and display-unit labels. Every save is balanced in finally, so a throw never leaks state (spec 6.5 containment).
  ctx.fillStyle = TEXT
  const rotated = (text: string, right: boolean) => { ctx.save(); try { ctx.translate(right ? box.x + box.w - fontPx : box.x + fontPx, plot.y + plot.h / 2); ctx.rotate(-Math.PI / 2); ctx.fillText(text, 0, 0) } finally { ctx.restore() } }
  const titled = [...baseAxes, ...valueAxes]
  for (const a of titled) {
    if (!a.axis.title) continue
    const side = sideOf(a)
    ctx.textAlign = 'center'; ctx.textBaseline = side === 'r' || side === 't' ? 'top' : 'bottom'
    if (a.orientation === 'v') rotated(a.axis.title, side === 'r')
    else ctx.fillText(a.axis.title, plot.x + plot.w / 2, side === 't' ? box.y : box.y + box.h)
  }
  for (const a of valueAxes) {
    const unit = a.display?.label
    if (!unit || a.deleted) continue
    const side = sideOf(a)
    if (a.orientation === 'v') {
      ctx.save()
      try { ctx.textAlign = 'center'; ctx.textBaseline = side === 'r' ? 'top' : 'bottom'; ctx.translate(side === 'r' ? labelAnchor(a) + labelBand(a) : labelAnchor(a) - labelBand(a), plot.y + plot.h / 2); ctx.rotate(-Math.PI / 2); ctx.fillText(unit, 0, 0) } finally { ctx.restore() }
    } else { ctx.textAlign = 'right'; ctx.textBaseline = side === 't' ? 'bottom' : 'top'; ctx.fillText(unit, plot.x + plot.w, side === 't' ? labelAnchor(a) - labelBand(a) - 3 : labelAnchor(a) + labelBand(a) + 3) }
  }

  ctx.save()
  try {
  ctx.beginPath(); ctx.rect(plot.x, plot.y, plot.w, plot.h); ctx.clip()
  const barGroups = model.groups.filter(g => g.family === 'bar')
  model.groups.forEach((group, g) => {
    const axes = graph.groups[g]
    if (!axes) return
    const { value: V, horizontal } = axes
    const B = axes.category ?? axes.x!
    /** The coordinate of a value on the group's value axis. */
    const at = (v: number) => along(V, axisFraction(V, v))
    const plotted = (v: number | undefined) => v !== undefined && V.plottable(v) ? v : undefined
    const crossing = B.crossesValue ?? Math.max(V.min, Math.min(V.max, 0))
    const stacked = group.grouping === 'stacked' || group.grouping === 'percentStacked' ? (() => {
      const pos = new Float64Array(count), neg = new Float64Array(count), total = new Float64Array(count)
      for (const s of group.series) for (let i = 0; i < count; i++) { const v = s.value(i); if (v === undefined) continue; if (v >= 0) pos[i] += v; else neg[i] += v; total[i] += Math.abs(v) }
      return { pos, neg, total }
    })() : undefined
    const share = (v: number, i: number) => group.grouping === 'percentStacked' && stacked ? (stacked.total[i] ? v / stacked.total[i] : 0) : v
    if (group.family === 'bar' && axes.category) {
      const C = axes.category, vertical = C.orientation === 'v'
      const geometry = barGeometry(group, C, axes.between, plot)
      const width = geometry.width
      const posBase = new Float64Array(count), negBase = new Float64Array(count)
      group.series.forEach((s, si) => {
        if (!admitPath(count * 4)) return
        const src = s.source
        for (let i = 0; i < count; i++) {
          const sample = s.plot(i), raw = plotted(typeof sample==='number'?sample:undefined)
          if (raw === undefined) continue
          const idx = graph.categoryIndex(C, i)
          if (!Number.isFinite(idx)) continue
          const v = share(raw, i)
          let from = crossing, to = v
          if (stacked) { const base = v >= 0 ? posBase : negBase; from = base[i]; to = base[i] + v; base[i] = to }
          const categoryStart = geometry.start(idx, si)
          const a = at(from), b = at(to)
          const srcIdx = model.view.source(i), point = src.dPt?.get(srcIdx)
          const bar = vertical ? { x: Math.min(a, b), y: categoryStart, w: Math.abs(b - a), h: width } : { x: categoryStart, y: Math.min(a, b), w: width, h: Math.abs(b - a) }
          // A negative point inverts when its dPt says so, else when its series does (spec 5.8): Office fills it with
          // the c14 invertSolidFillFmt colour, else white.
          const inverted = v < 0 && (point?.invertIfNegative ?? src.invertIfNegative) === true
          // The point colour (dPt, or varyColors in a lone single-series bar group, as before A4a), else the series fill.
          const color = s.pointColors?.[srcIdx]
          const pointColor = color && (point?.color || (barGroups.length === 1 && group.series.length === 1)) ? color : undefined
          const fill = inverted ? undefined : point?.fill ?? (pointColor ? undefined : src.fill)
          if (fillWith(ctx, fill, inverted ? src.invertColor ?? 'FFFFFF' : pointColor ?? s.color, bar)) ctx.fillRect(bar.x, bar.y, bar.w, bar.h)
          strokeBorder(ctx, point?.line ?? src.lineStyle, s.color, bar, bar)
        }
      })
      // Series lines (stacked bars, spec 5.8): from each series' segment end in one category to the same series'
      // in the next, across the gap between the two bars; admitted against the path work bound.
      const lines = group.serLines
      if (lines && !lines.noFill && stacked && count > 1 && group.series.length && admitPath((count - 1) * group.series.length * 2)) {
        ctx.save()
        try {
          ctx.strokeStyle = GRID; ctx.lineWidth = 1
          if (applyLineStyle(ctx, lines, plot)) {
            ctx.beginPath()
            const valueAt = (s: DrawSeries, i: number) => { const sample = s.plot(i), raw = plotted(typeof sample === 'number' ? sample : undefined); return raw === undefined ? 0 : share(raw, i) }
            for (let i = 0; i + 1 < count; i++) {
              const p = graph.categoryIndex(C, i), q = graph.categoryIndex(C, i + 1)
              if (!Number.isFinite(p) || !Number.isFinite(q)) continue
              const sp = geometry.start(p, 0), sq = geometry.start(q, 0)
              // The facing edges of the two bars along the category direction.
              const from = sq > sp ? sp + width : sp, to = sq > sp ? sq : sq + width
              // Running positive and negative stacks of both categories, as the bars stack them (no per-slot storage).
              let pos1 = 0, neg1 = 0, pos2 = 0, neg2 = 0
              for (const s of group.series) {
                const v1 = valueAt(s, i), v2 = valueAt(s, i + 1)
                const e1 = v1 >= 0 ? (pos1 += v1) : (neg1 += v1), e2 = v2 >= 0 ? (pos2 += v2) : (neg2 += v2)
                const y1 = at(e1), y2 = at(e2)
                if (vertical) { ctx.moveTo(y1, from); ctx.lineTo(y2, to) } else { ctx.moveTo(from, y1); ctx.lineTo(to, y2) }
              }
            }
            ctx.stroke()
          }
        } finally { ctx.restore() }
      }
    } else if ((group.family === 'line' || group.family === 'area') && horizontal === B) {
      // Streamed per series over three reused per-slot buffers (the running stack base, this series' bottom and
      // top; NaN marks no point), so the construction does not grow with the series count. Each series is drawn
      // as soon as it is computed, in the same order as before.
      const base = new Float64Array(count), bottom = new Float64Array(count), top = new Float64Array(count)
      const edge = !axes.between
      const band = plot.w / count
      const xAt = (i: number) => {
        const idx = axes.category ? graph.categoryIndex(B, i) : i
        if (!axes.category) return plot.x + axisFraction(B, i + 1, false) * plot.w
        if (B.role === 'date' || B.reversed) return along(B, categoryFraction(B, idx, axes.between))
        return edge ? plot.x + (count > 1 ? i * plot.w / (count - 1) : plot.w / 2) : plot.x + band * (i + 0.5)
      }
      const has = (i: number) => !Number.isNaN(top[i])
      // Drop lines (spec 5.8) span, per category, every plotted point of the group to the category axis.
      const drops = group.dropLines && !group.dropLines.noFill ? { lo: new Float64Array(count).fill(NaN), hi: new Float64Array(count).fill(NaN) } : undefined
      for (const s of group.series) {
        if (!admitPath(count * (group.family === 'area' ? 2 : s.markers ? 2 : 1) + 1)) continue
        const src = s.source
        bottom.set(base)
        for (let i = 0; i < count; i++) {
          const sample = s.plot(i), raw = plotted(typeof sample==='number'?sample:undefined)
          if (raw === undefined) { top[i] = stacked && group.family === 'area' ? base[i] : NaN; continue }
          if (!Number.isFinite(xAt(i))) { top[i] = NaN; continue }
          const v = share(raw, i)
          if (stacked) { base[i] += v; top[i] = base[i] } else top[i] = v
        }
        if (drops) for (let i = 0; i < count; i++) if (has(i)) { drops.lo[i] = Number.isNaN(drops.lo[i]) ? top[i] : Math.min(drops.lo[i], top[i]); drops.hi[i] = Number.isNaN(drops.hi[i]) ? top[i] : Math.max(drops.hi[i], top[i]) }
        ctx.strokeStyle = ctx.fillStyle = `#${s.color}`
        if (group.family === 'area') {
          ctx.beginPath()
          let any = false
          for(let start=0;start<count;){
            if(!has(start)){start++;continue}
            let end=start
            while(end+1<count&&has(end+1))end++
            ctx.moveTo(xAt(start),at(top[start]));any=true
            for(let i=start+1;i<=end;i++)ctx.lineTo(xAt(i),at(top[i]))
            for(let i=end;i>=start;i--)ctx.lineTo(xAt(i),at(stacked?bottom[i]:crossing))
            ctx.closePath();start=end+1
          }
          if(!any)continue
          ctx.globalAlpha = stacked ? 1 : 0.85
          if (fillWith(ctx, src.fill, s.color, plot)) ctx.fill()
          ctx.globalAlpha = 1
          strokeBorder(ctx, src.lineStyle, s.color, plot)
          continue
        }
        if (s.line !== false) {
          // The engine's series stroke (2.25 px, round joins, the series colour); an authored a:ln replaces only what it states.
          ctx.save()
          try {
            ctx.lineWidth = 2.25; ctx.lineJoin = 'round'
            if (applyLineStyle(ctx, src.lineStyle, plot)) {
              ctx.beginPath()
              let pen = false
              for (let i = 0; i < count; i++) { if (!has(i)) { if(s.plot(i)!=='span')pen = false; continue } const px = xAt(i), py = at(top[i]); if (!pen) { ctx.moveTo(px, py); pen = true } else ctx.lineTo(px, py) }
              ctx.stroke()
            }
          } finally { ctx.restore() }
        }
        if (s.markers) for (let i = 0; i < count; i++) { if (!has(i)) continue; drawMarker(ctx, s, model.view.source(i), xAt(i), at(top[i]), 2.5) }
      }
      if (drops && admitPath(count * 2 + 1)) {
        ctx.save()
        try {
          ctx.strokeStyle = GRID; ctx.lineWidth = 1
          if (applyLineStyle(ctx, group.dropLines, plot)) {
            const axis = at(crossing)
            ctx.beginPath()
            for (let i = 0; i < count; i++) {
              if (Number.isNaN(drops.hi[i])) continue
              const px = xAt(i), y1 = Math.min(at(drops.hi[i]), axis), y2 = Math.max(at(drops.lo[i]), axis)
              ctx.moveTo(px, y1); ctx.lineTo(px, y2)
            }
            ctx.stroke()
          }
        } finally { ctx.restore() }
      }
    } else if (group.family === 'scatter' && axes.x) {
      // Points are recomputed for each pass (line, markers) rather than kept: nothing per point outlives the pass.
      const X = axes.x
      const xp = (v: number) => plot.x + axisFraction(X, v, false) * plot.w
      let px = 0, py = 0
      const point = (s: DrawSeries, i: number): boolean => {
        const sample=s.plot(i), xs=s.hasX?s.plot(i,true):i+1
        const v = plotted(typeof sample==='number'?sample:undefined), xv = typeof xs==='number'?xs:undefined
        if (v === undefined || xv === undefined || !X.plottable(xv)) return false
        px = xp(xv); py = at(v)
        return true
      }
      for (const s of group.series) {
        if (!admitPath(count * 2 + 1)) continue
        ctx.strokeStyle = ctx.fillStyle = `#${s.color}`
        if (s.line) {
          ctx.save()
          try {
            ctx.lineWidth = 2
            if (applyLineStyle(ctx, s.source.lineStyle, plot)) {
              ctx.beginPath()
              let pen = false
              for (let i = 0; i < count; i++) {
                if (!point(s, i)) {
                  const y = s.plot(i), x = s.hasX ? s.plot(i, true) : i + 1
                  const rejectedNumber = (typeof y === 'number' && !V.plottable(y)) || (typeof x === 'number' && !X.plottable(x))
                  if (rejectedNumber || y === 'gap' || x === 'gap' || (y !== 'span' && x !== 'span')) pen = false
                  continue
                }
                if (!pen) { ctx.moveTo(px, py); pen = true }
                else ctx.lineTo(px, py)
              }
              ctx.stroke()
            }
          } finally { ctx.restore() }
        }
        if (s.markers !== false) for (let i = 0; i < count; i++) if (point(s, i)) drawMarker(ctx, s, model.view.source(i), px, py, 3)
      }
    }
  })
  } finally { ctx.restore() }
}

/** Paint a chart into a w×h box at the origin of `ctx` (worksheet CSS px). */
export function paintXlsxChart(source: ChartModel, ctx: CanvasRenderingContext2D, w: number, h: number, familyCss: FamilyCss = JSON.stringify, chartContext?: ChartDocumentContext): void {
  if (source.unsupported || !source.groups.length) { paintChartPlaceholder(source, ctx, w, h, familyCss); return }
  const ownerWorkspace = ownedChartWorkspace(source)
  const workspace = chartContext?.workspace ?? ownerWorkspace ?? createChartWorkspace(CHART_STANDALONE_PAINT_PROFILE)
  const standalone = !chartContext && !ownerWorkspace
  const context = chartContext ?? { profile: 'mobile' as const, workspace }
  const admission = { workspace, tag: `chart-paint#${++paints}` }
  let layout: ChartLayout
  let releaseLayout = () => {}
  try {
    const physicalDpr = ctx.canvas?.ownerDocument?.defaultView?.devicePixelRatio ?? globalThis.devicePixelRatio
    const dpr = Number.isFinite(physicalDpr) && physicalDpr > 0 ? physicalDpr : 1
    layout = layoutChart(source, w, h, dpr, fontGeneration(ctx), {
      width: (text, css) => { const previous = ctx.font; try { ctx.font = css; return ctx.measureText(text).width } finally { ctx.font = previous } },
      lineHeight: () => source.fontSizePt * 96 / 72,
    }, context, familyCss(source.fontFamily))
    releaseLayout = holdChartLayout(layout)
    const paintBytes = paintWorkspaceBytes(source, layout.view)
    if (!reserveChartLayoutWorkspace(workspace, paintBytes, admission.tag)) throw new ChartRefusal()
  } catch (error) {
    workspace.release(admission.tag)
    releaseLayout()
    if (standalone) releaseWorkspaceLayouts(workspace)
    if (!(error instanceof ChartRefusal)) throw error
    paintChartPlaceholder({ ...source, unsupported: 'This chart needs more memory than is available for charts' }, ctx, w, h, familyCss)
    return
  }
  const model = drawModel(source, layout.view), graph = layout.graph
  ctx.save()
  try {
    ctx.beginPath(); ctx.rect(0, 0, w, h); ctx.clip()
    // The chart area (spec 5.4): chartSpace/spPr (noFill: transparent), else the c:style default, else white.
    // Paper rule: the viewer's theme never reaches these pixels.
    const whole = { x: 0, y: 0, w, h }
    if (fillWith(ctx, model.chartAreaFill, model.chartAreaDefault ?? 'FFFFFF', whole)) ctx.fillRect(0, 0, w, h)
    if (!model.chartAreaLine) { ctx.strokeStyle = GRID; ctx.lineWidth = 1; ctx.strokeRect(0.5, 0.5, w - 1, h - 1) }
    else if (!model.chartAreaLine.noFill) {
      ctx.save()
      try { ctx.strokeStyle = GRID; ctx.lineWidth = 1; applyLineStyle(ctx, model.chartAreaLine, whole); ctx.strokeRect(0.5, 0.5, w - 1, h - 1) } finally { ctx.restore() }
    }
    const fontPx = model.fontSizePt * 96 / 72
    const pad = Math.min(10, w * 0.04, h * 0.04)
    let area: Box = { x: pad, y: pad, w: w - pad * 2, h: h - pad * 2 }
    if (model.title) {
      if (model.titleFill && fillWith(ctx, model.titleFill, 'FFFFFF', layout.title)) ctx.fillRect(layout.title.x, layout.title.y, layout.title.w, layout.title.h)
      strokeBorder(ctx, model.titleLine, '000000', layout.title, layout.title)
      const titlePx = Math.min(96, fontPx * 1.4)
      ctx.font = font(titlePx, model.fontFamily, familyCss)
      ctx.fillStyle = TEXT; ctx.textAlign = 'center'; ctx.textBaseline = 'top'
      ctx.fillText(model.title, layout.title.x + layout.title.w / 2, layout.title.y, layout.title.w)
      area = { ...area, y: area.y + titlePx + 6, h: area.h - titlePx - 6 }
    }
    paintLegend(ctx, model, layout.legend, fontPx, model.fontFamily, familyCss, layout)
    area = layout.area
    if (area.w <= 8 || area.h <= 8) return
    const group = model.groups[0]
    if (group.family === 'pie' || group.family === 'doughnut') paintPie(ctx, model, group, area)
    else if (group.family === 'contour') paintContour(ctx, model, group, area)
    else if (graph) paintCartesian(ctx, model, graph, area, fontPx, model.fontFamily, familyCss, layout)
    paintDataLabels(source,ctx,layout,familyCss(source.fontFamily))
  } finally { ctx.restore(); admission.workspace.release(admission.tag); releaseLayout(); if (standalone) releaseWorkspaceLayouts(workspace) }
}
let paints = 0
/**
 * The profile of the bounded per-call workspace a synchronous paint uses when the model has no owner, no chart
 * context is passed and no receiving document was associated (a headless `renderPages`/`renderSlide` of a
 * transferred model): a fresh 32 MiB (mobile-limit) chart workspace per call, released with the paint.
 * (The receiving preparation's context, when there is one, is passed per paint instead.)
 */
export const CHART_STANDALONE_PAINT_PROFILE: ChartProfile = 'mobile'
/** What the receiving side knows: its document, the viewer profile, its live parent budget, or a context it already holds. */
export interface ChartReceiver { owner: object; profile?: ChartProfile; context?: ChartDocumentContext; parent?: ChartParentLedger }
/** A receiving context held by one preparation or asset lease; `release` ends the hold (idempotent). */
export interface ChartReceiverLease { context: ChartDocumentContext; release(): void }
interface Held { context: ChartDocumentContext; refs: number }
const workspaceLeases = new WeakMap<ChartWorkspace, number>()
function holdLayouts(workspace: ChartWorkspace): () => void {
  workspaceLeases.set(workspace, (workspaceLeases.get(workspace) ?? 0) + 1)
  let released = false
  return () => {
    if (released) return; released = true
    const n = (workspaceLeases.get(workspace) ?? 1) - 1
    if (n) workspaceLeases.set(workspace, n)
    else { workspaceLeases.delete(workspace); releaseWorkspaceLayouts(workspace) }
  }
}
/** Receiving contexts per document, then per parent ledger (null: none), then per profile. */
const documents = new WeakMap<object, Map<ChartParentLedger | null, Map<ChartProfile, Held>>>()
/**
 * The receiving side of a transfer (spec 6.4): the chart context a preparation (or asset lease) paints its
 * transferred models with, carried per paint (`ContentPaintAssets.chartContext`). A supplied context is used as is.
 * Otherwise there is one context per (document, parent ledger, profile): profile-sized (mobile: 32 MiB) and parented
 * to that live budget, shared by the leases that hold it and dropped when the last one is released. A later parent
 * or profile never reuses an earlier receiver's context, and simultaneous receivers keep their own.
 */
export function acquireChartContext(receiver: ChartReceiver): ChartReceiverLease {
  if (receiver.context) return { context: receiver.context, release: holdLayouts(receiver.context.workspace) }
  const profile = receiver.profile ?? 'desktop', parent = receiver.parent ?? null
  let byParent = documents.get(receiver.owner)
  if (!byParent) documents.set(receiver.owner, byParent = new Map())
  let byProfile = byParent.get(parent)
  if (!byProfile) byParent.set(parent, byProfile = new Map())
  let held = byProfile.get(profile)
  if (!held) byProfile.set(profile, held = { context: documentChartContext(profile, undefined, receiver.parent), refs: 0 })
  held.refs++
  const entry = held, releaseLayouts = holdLayouts(held.context.workspace)
  let released = false
  return {
    context: entry.context,
    release: () => {
      if (released) return
      released = true
      releaseLayouts()
      if (--entry.refs > 0) return
      byProfile!.delete(profile)
      if (!byProfile!.size) byParent!.delete(parent)
    },
  }
}

// FontFaceSet is realm-local. Embedded-font aliases also participate in the layout key.
const fontSets = new WeakMap<FontFaceSet, { generation: number }>()
function fontGeneration(ctx: CanvasRenderingContext2D): number {
  const fonts = ctx.canvas?.ownerDocument?.fonts ?? (globalThis as { fonts?: FontFaceSet }).fonts
  if (!fonts) return 0
  let state = fontSets.get(fonts)
  if (!state) { state = { generation: 0 }; const current = state; fonts.addEventListener('loadingdone', () => { current.generation++; invalidateChartFonts() }); fontSets.set(fonts, state) }
  return state.generation
}
