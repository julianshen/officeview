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
import { ChartRefusal, chartModelOwner } from './own'
import { createChartWorkspace, type ChartParentLedger, type ChartProfile, type ChartWorkspace } from './workspace'
import { documentChartContext, type ChartDocumentContext } from './context'
import { dateUnitIndex } from './dates'
import { layoutChart, invalidateChartFonts, holdChartLayout, paintWorkspaceBytes, releaseWorkspaceLayouts, type ChartLayout, textAngle } from './layout'

type FamilyCss = (family: string) => string
/**
 * What the painter draws: the display view (plotVisOnly per role, points
 * capped), read slot by slot from the compact storage; no per-point array is
 * copied. Every non-value state (blank, error, unavailable) is a gap here, as
 * before D0; the per-family point matrix (spec 5.8) is A2's.
 */
interface DrawSeries extends Pick<ChartSeries, 'name' | 'color' | 'pointColors' | 'line' | 'markers'> {
  value(slot: number): number | undefined
  x(slot: number): number | undefined
  hasX: boolean
}
interface DrawGroup extends Omit<ChartGroup, 'series'> { series: DrawSeries[] }
interface DrawModel extends Omit<ChartModel, 'groups'> { groups: DrawGroup[]; view: DisplayView; source: ChartModel }
function drawModel(model: ChartModel, view = displayView(model)): DrawModel {
  return {
    ...model,
    view,
    source: model,
    groups: model.groups.map((group, g) => ({ ...group, series: view.groupSeries(g).map(({ series: s, value, x, hasX }): DrawSeries => ({ name: s.name, color: s.color, pointColors: s.pointColors, line: s.line, markers: s.markers, value, x, hasX })) })),
  }
}
const font = (sizePx: number, family: string, familyCss: FamilyCss, bold = false) => `${bold ? 'bold ' : ''}${sizePx}px ${familyCss(family)}`
const TEXT = '#595959', GRID = '#D9D9D9'

interface Box { x: number; y: number; w: number; h: number }
function legendEntries(model: DrawModel): Array<{ name: string; color: string }> {
  const group = model.groups[0]
  if (group && (group.family === 'pie' || group.family === 'doughnut')) {
    const s = group.series[0], out: Array<{ name: string; color: string }> = []
    if (s) for (let slot = 0; slot < Math.min(model.view.slots, 64); slot++) { const i = model.view.source(slot); out.push({ name: model.view.category(slot) ?? String(i + 1), color: s.pointColors?.[i] ?? s.color }) }
    return out
  }
  return model.groups.flatMap(g => g.series).map(s => ({ name: s.name, color: s.color }))
}

function paintLegend(ctx: CanvasRenderingContext2D, model: DrawModel, area: Box, fontPx: number, family: string, familyCss: FamilyCss, layout: ChartLayout): Box {
  const entries = legendEntries(model)
  if (!model.legend || !entries.length) return area
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
  const values: number[] = []
  for (let slot = 0; slot < model.view.slots; slot++) { const v = series.value(slot); values.push(v !== undefined && v > 0 ? v : 0) }
  const total = values.reduce((a, b) => a + b, 0)
  if (!(total > 0)) return
  const r = Math.max(1, Math.min(box.w, box.h) / 2 - 4), cx = box.x + box.w / 2, cy = box.y + box.h / 2
  let angle = ((group.firstSliceAngle ?? 0) - 90) * Math.PI / 180
  values.forEach((value, i) => {
    if (!value) return
    const sweep = value / total * Math.PI * 2
    ctx.beginPath()
    ctx.moveTo(cx, cy)
    ctx.arc(cx, cy, r, angle, angle + sweep)
    ctx.closePath()
    ctx.fillStyle = `#${series.pointColors?.[model.view.source(i)] ?? series.color}`
    ctx.fill()
    ctx.strokeStyle = '#FFFFFF'; ctx.lineWidth = 1; ctx.stroke()
    angle += sweep
  })
  if (group.family === 'doughnut') {
    ctx.beginPath()
    ctx.arc(cx, cy, r * (group.holeSize ?? 50) / 100, 0, Math.PI * 2)
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
  const crossOf = (a: ResolvedAxis) => a.crosses !== undefined ? graph.axes.get(a.crosses) : undefined
  /** Category labels sit between ticks unless the plot authored crossBetween=midCat (legacy: band centres). */
  const labelsBetween = (a: ResolvedAxis) => { const g = placed.find(p => p.category === a); return !(g && (g.value.axis.crossBetween ?? a.axis.crossBetween) === 'midCat') }
  const shown = (a: ResolvedAxis) => !a.deleted && a.axis.tickLblPos !== 'none'
  /**
   * Where an axis's tick labels go: a plot edge, or next to the axis line. Without an authored tickLblPos the
   * engine's legacy placement is kept (the near edge), except that an axis authored to cross at the maximum (a
   * secondary axis) or across a reversed axis takes the far edge.
   */
  const labelPlace = (a: ResolvedAxis): 'near' | 'far' | 'axis' => {
    const pos = a.axis.tickLblPos
    const reversedCross = crossOf(a)?.reversed === true
    if (pos === undefined) return (a.axis.crosses === 'max') !== reversedCross ? 'far' : 'near'
    if (pos === 'low') return reversedCross ? 'far' : 'near'
    if (pos === 'high') return reversedCross ? 'near' : 'far'
    return 'axis'
  }
  const sideOf = (a: ResolvedAxis): 'l' | 'r' | 'b' | 't' => {
    const place = labelPlace(a)
    if (place === 'axis') return a.side
    return a.orientation === 'v' ? (place === 'near' ? 'l' : 'r') : (place === 'near' ? 'b' : 't')
  }
  const dateLabels = (a: ResolvedAxis) => layout.texts.get(a.id)!
  /** The thickness of an axis's tick labels on its side (the engine's legacy measures). */
  const labelBand = (a: ResolvedAxis): number => {
    if (a.deleted) return 0
    if (a.axis.tickLblPos === 'none') return 0
    if (a.orientation === 'h') return fontPx + 6
    if (a.role === 'value') return Math.max(...layout.texts.get(a.id)!.map(text => layout.width(text, font(fontPx, family, familyCss)))) + 6 + (a.axis.majorTickMark === 'out' || a.axis.majorTickMark === 'cross' ? 4 : 0)
    let widest = 0
    if (a.role === 'date') for (const text of dateLabels(a)) widest = Math.max(widest, layout.width(text, font(fontPx, family, familyCss)))
    else for (let slot = 0; slot < count; slot++) { const c = model.view.category(slot); if (c !== undefined) widest = Math.max(widest, layout.width(c, font(fontPx, family, familyCss))) }
    return Math.min(box.w * 0.35, widest + 8)
  }
  const plot = layout.plot
  if (plot.w <= 4 || plot.h <= 4) return
  /** A coordinate along an axis's direction (x for horizontal axes, y for vertical ones) at fraction f. */
  const along = (a: ResolvedAxis, f: number) => a.orientation === 'h' ? plot.x + f * plot.w : plot.y + plot.h - f * plot.h
  /** The coordinate of an axis's line across the plot (its crossing on the axis it crosses). */
  const across = (a: ResolvedAxis) => a.orientation === 'h' ? plot.y + plot.h - a.crossesFraction * plot.h : plot.x + a.crossesFraction * plot.w
  const labelAnchor = (a: ResolvedAxis): number => {
    const place = labelPlace(a)
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
      const items = a.ticks.map((t, k) => ({ k, p: along(a, dateAt(t)), extent: a.orientation === 'h' ? layout.width(texts[k], font(fontPx, family, familyCss)) : fontPx }))
      items.sort((p, q) => p.p - q.p)
      for (const { k, p } of items) {
        if (!layout.labelIndices.get(a.id)?.has(k)) continue
        if (a.orientation === 'h') { ctx.textAlign = 'center'; ctx.textBaseline = place === 't' ? 'bottom' : 'top'; label(a, k, texts[k], p, place === 't' ? anchor - 3 : anchor + 3) }
        else { ctx.textAlign = place === 'r' ? 'left' : 'right'; ctx.textBaseline = 'middle'; label(a, k, texts[k], place === 'r' ? anchor + 4 : anchor - 4, p) }
      }
      continue
    }
    if (model.view.category(0) === undefined) continue
    let widest = 0
    for (let slot = 0; slot < count; slot++) { const c = model.view.category(slot); if (c !== undefined) widest = Math.max(widest, layout.width(c, font(fontPx, family, familyCss))) }
    const skip = a.axis.tickLblSkip ?? (a.orientation === 'v' ? Math.max(1, Math.ceil((fontPx + 2) / band)) : Math.max(1, Math.ceil((widest + 6) / band)))
    // Legacy positions (band centres) when unreversed and between; otherwise the axis fraction.
    const at = (i: number) => between && !a.reversed ? (a.orientation === 'h' ? plot.x + band * (i + 0.5) : plot.y + plot.h - band * (i + 0.5)) : along(a, categoryFraction(a, i, between))
    for (let i = 0; i < count; i++) {
      const c = model.view.category(i)
      if (c === undefined) continue
      if (a.orientation === 'v') { ctx.textAlign = place === 'r' ? 'left' : 'right'; ctx.textBaseline = 'middle'; label(a, i, c, place === 'r' ? anchor + 4 : anchor - 4, at(i)) }
      else { ctx.textAlign = 'center'; ctx.textBaseline = place === 't' ? 'bottom' : 'top'; label(a, i, c, at(i), place === 't' ? anchor - 3 : anchor + 3, skip * band) }
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
      const n = group.grouping === 'clustered' ? Math.max(1, group.series.length) : 1
      const stride = 1 - group.overlap / 100
      const span = 1 + (n - 1) * stride
      const len = vertical ? plot.h : plot.w
      const band = axes.between ? len / C.slots : len / Math.max(1, C.slots - 1)
      const width = band / (span + group.gapWidth / 100)
      /** The start of a category's band (legacy arithmetic when between and unreversed). */
      const slotStart = (idx: number): number => {
        if (!axes.between) return along(C, categoryFraction(C, idx, false)) - band / 2
        if (vertical) return C.reversed ? plot.y + band * idx : plot.y + plot.h - band * (idx + 1)
        return C.reversed ? plot.x + plot.w - band * (idx + 1) : plot.x + band * idx
      }
      const posBase = new Float64Array(count), negBase = new Float64Array(count)
      group.series.forEach((s, si) => {
        if (!admitPath(count * 4)) return
        for (let i = 0; i < count; i++) {
          const raw = plotted(s.value(i))
          if (raw === undefined) continue
          const idx = graph.categoryIndex(C, i)
          if (!Number.isFinite(idx)) continue
          const v = share(raw, i)
          let from = crossing, to = v
          if (stacked) { const base = v >= 0 ? posBase : negBase; from = base[i]; to = base[i] + v; base[i] = to }
          const slot = group.grouping === 'clustered' ? si : 0
          const offset = (band - width * span) / 2 + width * slot * stride
          const a = at(from), b = at(to)
          const color = s.pointColors?.[model.view.source(i)]
          ctx.fillStyle = `#${color && barGroups.length === 1 && group.series.length === 1 ? color : s.color}`
          if (vertical) ctx.fillRect(Math.min(a, b), slotStart(idx) + offset, Math.abs(b - a), width)
          else ctx.fillRect(slotStart(idx) + offset, Math.min(a, b), width, Math.abs(b - a))
        }
      })
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
      for (const s of group.series) {
        if (!admitPath(count * (group.family === 'area' ? 2 : s.markers ? 2 : 1) + 1)) continue
        bottom.set(base)
        for (let i = 0; i < count; i++) {
          const raw = plotted(s.value(i))
          if (raw === undefined) { top[i] = stacked && group.family === 'area' ? base[i] : NaN; continue }
          if (!Number.isFinite(xAt(i))) { top[i] = NaN; continue }
          const v = share(raw, i)
          if (stacked) { base[i] += v; top[i] = base[i] } else top[i] = v
        }
        ctx.strokeStyle = ctx.fillStyle = `#${s.color}`
        if (group.family === 'area') {
          ctx.beginPath()
          let started = false
          for (let i = 0; i < count; i++) { if (!has(i)) continue; const px = xAt(i), py = at(top[i]); if (!started) { ctx.moveTo(px, py); started = true } else ctx.lineTo(px, py) }
          if (!started) continue
          for (let i = count - 1; i >= 0; i--) if (has(i)) ctx.lineTo(xAt(i), at(stacked ? bottom[i] : crossing))
          ctx.closePath()
          ctx.globalAlpha = stacked ? 1 : 0.85
          ctx.fill()
          ctx.globalAlpha = 1
          continue
        }
        if (s.line !== false) {
          ctx.lineWidth = 2.25; ctx.lineJoin = 'round'
          ctx.beginPath()
          let pen = false
          for (let i = 0; i < count; i++) { if (!has(i)) { pen = false; continue } const px = xAt(i), py = at(top[i]); if (!pen) { ctx.moveTo(px, py); pen = true } else ctx.lineTo(px, py) }
          ctx.stroke()
        }
        if (s.markers) for (let i = 0; i < count; i++) { if (!has(i)) continue; ctx.beginPath(); ctx.arc(xAt(i), at(top[i]), 2.5, 0, Math.PI * 2); ctx.fill() }
      }
    } else if (group.family === 'scatter' && axes.x) {
      // Points are recomputed for each pass (line, markers) rather than kept: nothing per point outlives the pass.
      const X = axes.x
      const xp = (v: number) => plot.x + axisFraction(X, v, false) * plot.w
      let px = 0, py = 0
      const point = (s: DrawSeries, i: number): boolean => {
        const v = plotted(s.value(i)), xv = s.hasX ? s.x(i) : i + 1
        if (v === undefined || xv === undefined || !X.plottable(xv)) return false
        px = xp(xv); py = at(v)
        return true
      }
      for (const s of group.series) {
        if (!admitPath(count * 2 + 1)) continue
        ctx.strokeStyle = ctx.fillStyle = `#${s.color}`
        if (s.line) {
          ctx.lineWidth = 2; ctx.beginPath()
          let pen = false
          for (let i = 0; i < count; i++) { if (!point(s, i)) { pen = false; continue } if (!pen) { ctx.moveTo(px, py); pen = true } else ctx.lineTo(px, py) }
          ctx.stroke()
        }
        if (s.markers !== false) for (let i = 0; i < count; i++) if (point(s, i)) { ctx.beginPath(); ctx.arc(px, py, 3, 0, Math.PI * 2); ctx.fill() }
      }
    }
  })
  } finally { ctx.restore() }
}

/** Paint a chart into a w×h box at the origin of `ctx` (worksheet CSS px). */
export function paintXlsxChart(source: ChartModel, ctx: CanvasRenderingContext2D, w: number, h: number, familyCss: FamilyCss = JSON.stringify, chartContext?: ChartDocumentContext): void {
  if (source.unsupported || !source.groups.length) { paintChartPlaceholder(source, ctx, w, h, familyCss); return }
  const owner = chartModelOwner(source)
  const ownerWorkspace = owner && ((owner.workspace as { facet?: ChartWorkspace }).facet ?? owner.workspace)
  const workspace = chartContext?.workspace ?? ownerWorkspace ?? createChartWorkspace(CHART_STANDALONE_PAINT_PROFILE)
  const standalone = !chartContext && !ownerWorkspace
  const context = chartContext ?? { profile: 'mobile' as const, workspace }
  const admission = { workspace, tag: `chart-paint#${++paints}` }
  let layout: ChartLayout
  let releaseLayout = () => {}
  try {
    const transform = ctx.getTransform?.()
    const dpr = transform ? Math.max(.001, Math.hypot(transform.a, transform.b)) : 1
    layout = layoutChart(source, w, h, dpr, fontGeneration(ctx), {
      width: (text, css) => { const previous = ctx.font; try { ctx.font = css; return ctx.measureText(text).width } finally { ctx.font = previous } },
      lineHeight: () => source.fontSizePt * 96 / 72,
    }, context, familyCss(source.fontFamily))
    releaseLayout = holdChartLayout(layout)
    const paintBytes = paintWorkspaceBytes(source, layout.view)
    if (!workspace.reserve(paintBytes, admission.tag)) throw new ChartRefusal()
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
    ctx.fillStyle = '#FFFFFF'; ctx.fillRect(0, 0, w, h)
    ctx.strokeStyle = GRID; ctx.lineWidth = 1; ctx.strokeRect(0.5, 0.5, w - 1, h - 1)
    const fontPx = model.fontSizePt * 96 / 72
    const pad = Math.min(10, w * 0.04, h * 0.04)
    let area: Box = { x: pad, y: pad, w: w - pad * 2, h: h - pad * 2 }
    if (model.title) {
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
    else if (graph) paintCartesian(ctx, model, graph, area, fontPx, model.fontFamily, familyCss, layout)
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
  const fonts = ctx.canvas?.ownerDocument?.fonts
  if (!fonts) return 0
  let state = fontSets.get(fonts)
  if (!state) { state = { generation: 0 }; const current = state; fonts.addEventListener('loadingdone', () => { current.generation++; invalidateChartFonts() }); fontSets.set(fonts, state) }
  return state.generation
}
