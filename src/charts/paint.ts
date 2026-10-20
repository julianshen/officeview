/**
 * Chart painter (lazy chunk): the common chart types (bar/column
 * clustered/stacked/100%, line, area, pie, doughnut, scatter). Anything else
 * paints the eager placeholder card. Moved from `src/xlsx/chart.ts` (A0a);
 * since A1a the cartesian families draw on the axis graph (`charts/axes.ts`).
 */
import { paintChartPlaceholder } from '../drawing/content-paint'
import { formatNumber } from '../xlsx/chartDescription'
import { displayView, type DisplayView } from './display'
import type { ChartGroup, ChartModel, ChartSeries } from './model'
import { axisFraction, categoryFraction, resolveAxes, type GroupAxes, type ResolvedAxis } from './axes'
import { dateTickLabel, dateUnitIndex } from './dates'

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
function drawModel(model: ChartModel): DrawModel {
  const view = displayView(model)
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

function paintLegend(ctx: CanvasRenderingContext2D, model: DrawModel, area: Box, fontPx: number, family: string, familyCss: FamilyCss): Box {
  const entries = legendEntries(model)
  if (!model.legend || !entries.length) return area
  ctx.font = font(fontPx, family, familyCss)
  const sw = fontPx * 0.7, gap = 6
  const widths = entries.map(e => ctx.measureText(e.name).width + sw + gap * 2)
  const line = fontPx + 6
  if (model.legend === 'r' || model.legend === 'l' || model.legend === 'tr') {
    const width = Math.min(area.w * 0.4, Math.max(...widths) + gap)
    const total = Math.min(entries.length, Math.max(1, Math.floor(area.h / line))) * line
    const x = model.legend === 'l' ? area.x : area.x + area.w - width
    let y = model.legend === 'tr' ? area.y : area.y + (area.h - total) / 2
    ctx.textBaseline = 'middle'; ctx.textAlign = 'left'
    for (const entry of entries) {
      if (y + line > area.y + area.h + 0.5) break
      ctx.fillStyle = `#${entry.color}`
      ctx.fillRect(x + gap, y + (line - sw) / 2, sw, sw)
      ctx.fillStyle = TEXT
      ctx.fillText(entry.name, x + gap * 2 + sw, y + line / 2, width - sw - gap * 3)
      y += line
    }
    return model.legend === 'l' ? { x: area.x + width, y: area.y, w: area.w - width, h: area.h } : { ...area, w: area.w - width }
  }
  const total = widths.reduce((a, b) => a + b, 0)
  let x = area.x + Math.max(0, (area.w - total) / 2)
  const y = model.legend === 't' ? area.y : area.y + area.h - line
  ctx.textBaseline = 'middle'; ctx.textAlign = 'left'
  entries.forEach((entry, i) => {
    if (x + widths[i] > area.x + area.w + 0.5) return
    ctx.fillStyle = `#${entry.color}`
    ctx.fillRect(x + gap, y + (line - sw) / 2, sw, sw)
    ctx.fillStyle = TEXT
    ctx.fillText(entry.name, x + gap * 2 + sw, y + line / 2)
    x += widths[i]
  })
  return model.legend === 't' ? { x: area.x, y: area.y + line, w: area.w, h: area.h - line } : { ...area, h: area.h - line }
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
function paintCartesian(ctx: CanvasRenderingContext2D, model: DrawModel, box: Box, fontPx: number, family: string, familyCss: FamilyCss): void {
  const graph = resolveAxes(model.source, model.view)
  const placed = graph.groups.filter((g): g is GroupAxes => Boolean(g))
  if (!placed.length) return
  const count = model.view.slots
  ctx.font = font(fontPx, family, familyCss)
  const used = [...graph.axes.values()].filter(a => placed.some(g => g.horizontal === a || g.vertical === a))
  const valueAxes = used.filter(a => placed.some(g => g.value === a))
  const baseAxes = used.filter(a => !valueAxes.includes(a))
  const crossOf = (a: ResolvedAxis) => a.crosses !== undefined ? graph.axes.get(a.crosses) : undefined
  /** Category labels sit between ticks unless the plot authored crossBetween=midCat (legacy: band centres). */
  const labelsBetween = (a: ResolvedAxis) => { const g = placed.find(p => p.category === a); return !(g && (g.value.axis.crossBetween ?? a.axis.crossBetween) === 'midCat') }
  const valueText = (a: ResolvedAxis, t: number) => formatNumber(a.labelValue(t), a.percent)
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
  const titleBand = fontPx + 6
  const dateLabels = (a: ResolvedAxis) => a.ticks.map(t => dateTickLabel(t, a.dateBase ?? 'days', graph.epoch))
  /** The thickness of an axis's tick labels on its side (the engine's legacy measures). */
  const labelBand = (a: ResolvedAxis): number => {
    if (a.deleted) return 0
    if (a.axis.tickLblPos === 'none') return 0
    if (a.orientation === 'h') return fontPx + 6
    if (a.role === 'value') return Math.max(...a.ticks.map(t => ctx.measureText(valueText(a, t)).width)) + 6 + (a.axis.majorTickMark === 'out' || a.axis.majorTickMark === 'cross' ? 4 : 0)
    let widest = 0
    if (a.role === 'date') for (const text of dateLabels(a)) widest = Math.max(widest, ctx.measureText(text).width)
    else for (let slot = 0; slot < count; slot++) { const c = model.view.category(slot); if (c !== undefined) widest = Math.max(widest, ctx.measureText(c).width) }
    return Math.min(box.w * 0.35, widest + 8)
  }
  // Layout: each side's bands from the plot outward (tick labels, display-unit label, title), applied in the
  // engine's legacy order so a single-axis chart keeps its exact geometry.
  const bands: Record<'l' | 'r' | 'b' | 't', number[]> = { l: [], r: [], b: [], t: [] }
  for (const a of [...valueAxes, ...baseAxes].sort((p, q) => (p.orientation === 'v' ? 0 : 1) - (q.orientation === 'v' ? 0 : 1))) {
    const side = sideOf(a)
    bands[side].push(labelBand(a))
    if (a.display?.label && !a.deleted) bands[side].push(titleBand)
    if (a.axis.title) bands[side].push(titleBand)
  }
  // Legacy: one band list per side (vertical axes first); horizontal-bar charts reserve the category band first.
  let x = box.x, w = box.w
  for (const b of bands.l) { x += b; w -= b }
  w -= 8
  for (const b of bands.r) w -= b
  let h = box.h
  for (const b of bands.b) h -= b
  h -= 4
  let y = box.y + 4
  for (const b of bands.t) { y += b; h -= b }
  const plot: Box = { x, y, w, h }
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
    if (a.axis.minorGridlines) { ctx.strokeStyle = '#F2F2F2'; for (const t of a.minorTicks) gridline(a, axisFraction(a, t)) }
    for (const t of a.ticks) {
      const f = axisFraction(a, t)
      ctx.strokeStyle = GRID
      if (a.axis.gridlines || t === a.min) gridline(a, f)
      if (!shown(a)) continue
      ctx.fillStyle = TEXT
      const anchor = labelAnchor(a), right = sideOf(a) === 'r', top = sideOf(a) === 't'
      // An authored outward tick mark pushes the labels clear of it.
      const pad = a.axis.majorTickMark === 'out' || a.axis.majorTickMark === 'cross' ? 4 : 0
      if (a.orientation === 'h') { ctx.textAlign = 'center'; ctx.textBaseline = top ? 'bottom' : 'top'; ctx.fillText(valueText(a, t), plot.x + f * plot.w, top ? anchor - 3 - pad : anchor + 3 + pad) }
      else { ctx.textAlign = right ? 'left' : 'right'; ctx.textBaseline = 'middle'; ctx.fillText(valueText(a, t), right ? anchor + 4 + pad : anchor - 4 - pad, plot.y + plot.h - f * plot.h) }
    }
    tickMarks(a, a.minorTicks, t => along(a, axisFraction(a, t)), a.axis.minorTickMark)
    tickMarks(a, a.ticks, t => along(a, axisFraction(a, t)), a.axis.majorTickMark)
  }
  // Category, date and x axes: the axis line at its crossing, then labels.
  for (const a of baseAxes) {
    if (a.role === 'value') {
      // An x value axis (scatter): labels and gridlines per tick, as before.
      ctx.textAlign = 'center'; ctx.textBaseline = 'top'
      for (const t of a.ticks) {
        const px = plot.x + axisFraction(a, t, false) * plot.w
        if (shown(a)) { ctx.fillStyle = TEXT; ctx.fillText(formatNumber(a.labelValue(t), a.percent), px, labelAnchor(a) + 3) }
        if (a.axis.gridlines) { ctx.strokeStyle = GRID; ctx.beginPath(); ctx.moveTo(Math.round(px) + 0.5, plot.y); ctx.lineTo(Math.round(px) + 0.5, plot.y + plot.h); ctx.stroke() }
      }
      tickMarks(a, a.ticks, t => plot.x + axisFraction(a, t, false) * plot.w, a.axis.majorTickMark)
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
    if (!shown(a)) continue
    const anchor = labelAnchor(a), place = sideOf(a)
    ctx.fillStyle = TEXT
    if (a.role === 'date') {
      // Date axes: labels at the major ticks (calendar steps), thinned so they never overlap.
      const texts = dateLabels(a)
      let last = -Infinity
      a.ticks.forEach((t, k) => {
        const f = categoryFraction(a, dateUnitIndex(t, a.dateBase!, graph.epoch) - a.dateStart!, between)
        const p = along(a, f), width = a.orientation === 'h' ? ctx.measureText(texts[k]).width : fontPx
        if (p - width / 2 < last + 4) return
        last = p + width / 2
        if (a.orientation === 'h') { ctx.textAlign = 'center'; ctx.textBaseline = place === 't' ? 'bottom' : 'top'; ctx.fillText(texts[k], p, place === 't' ? anchor - 3 : anchor + 3) }
        else { ctx.textAlign = place === 'r' ? 'left' : 'right'; ctx.textBaseline = 'middle'; ctx.fillText(texts[k], place === 'r' ? anchor + 4 : anchor - 4, p) }
      })
      tickMarks(a, a.ticks, t => along(a, categoryFraction(a, dateUnitIndex(t, a.dateBase!, graph.epoch) - a.dateStart!, between)), a.axis.majorTickMark)
      continue
    }
    if (model.view.category(0) === undefined) continue
    let widest = 0
    for (let slot = 0; slot < count; slot++) { const c = model.view.category(slot); if (c !== undefined) widest = Math.max(widest, ctx.measureText(c).width) }
    const skip = a.axis.tickLblSkip ?? (a.orientation === 'v' ? Math.max(1, Math.ceil((fontPx + 2) / band)) : Math.max(1, Math.ceil((widest + 6) / band)))
    // Legacy positions (band centres) when unreversed and between; otherwise the axis fraction.
    const at = (i: number) => between && !a.reversed ? (a.orientation === 'h' ? plot.x + band * (i + 0.5) : plot.y + plot.h - band * (i + 0.5)) : along(a, categoryFraction(a, i, between))
    for (let i = 0; i < count; i++) {
      const c = model.view.category(i)
      if (c === undefined || i % skip) continue
      if (a.orientation === 'v') { ctx.textAlign = place === 'r' ? 'left' : 'right'; ctx.textBaseline = 'middle'; ctx.fillText(c, place === 'r' ? anchor + 4 : anchor - 4, at(i)) }
      else { ctx.textAlign = 'center'; ctx.textBaseline = place === 't' ? 'bottom' : 'top'; ctx.fillText(c, at(i), place === 't' ? anchor - 3 : anchor + 3, skip * band) }
    }
    const marks = Array.from({ length: between ? n + 1 : n }, (_, k) => k)
    tickMarks(a, marks.filter(k => k % (a.axis.tickMarkSkip ?? 1) === 0), k => along(a, between ? (a.reversed ? 1 - k / n : k / n) : categoryFraction(a, k, false)), a.axis.majorTickMark)
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
      const pos = new Array<number>(count).fill(0), neg = new Array<number>(count).fill(0), total = new Array<number>(count).fill(0)
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
      const posBase = new Array<number>(count).fill(0), negBase = new Array<number>(count).fill(0)
      group.series.forEach((s, si) => {
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
      const base = new Array<number>(count).fill(0)
      const edge = !axes.between
      const band = plot.w / count
      const xAt = (i: number) => {
        const idx = axes.category ? graph.categoryIndex(B, i) : i
        if (!axes.category) return plot.x + axisFraction(B, i + 1, false) * plot.w
        if (B.role === 'date' || B.reversed) return along(B, categoryFraction(B, idx, axes.between))
        return edge ? plot.x + (count > 1 ? i * plot.w / (count - 1) : plot.w / 2) : plot.x + band * (i + 0.5)
      }
      const layers: Array<{ s: DrawSeries; top: Array<number | undefined>; bottom: number[] }> = []
      for (const s of group.series) {
        const bottom = base.slice()
        const top = Array.from({ length: count }, (_, i) => {
          const raw = plotted(s.value(i))
          if (raw === undefined) return stacked && group.family === 'area' ? base[i] : undefined
          if (!Number.isFinite(xAt(i))) return undefined
          const v = share(raw, i)
          if (stacked) { base[i] += v; return base[i] }
          return v
        })
        layers.push({ s, top, bottom })
      }
      for (const { s, top, bottom } of layers) {
        ctx.strokeStyle = ctx.fillStyle = `#${s.color}`
        if (group.family === 'area') {
          ctx.beginPath()
          let started = false
          top.forEach((v, i) => { if (v === undefined) return; const px = xAt(i), py = at(v); if (!started) { ctx.moveTo(px, py); started = true } else ctx.lineTo(px, py) })
          if (!started) continue
          for (let i = count - 1; i >= 0; i--) if (top[i] !== undefined) ctx.lineTo(xAt(i), at(stacked ? bottom[i] : crossing))
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
          top.forEach((v, i) => { if (v === undefined) { pen = false; return } const px = xAt(i), py = at(v); if (!pen) { ctx.moveTo(px, py); pen = true } else ctx.lineTo(px, py) })
          ctx.stroke()
        }
        if (s.markers) top.forEach((v, i) => { if (v === undefined) return; ctx.beginPath(); ctx.arc(xAt(i), at(v), 2.5, 0, Math.PI * 2); ctx.fill() })
      }
    } else if (group.family === 'scatter' && axes.x) {
      const X = axes.x
      const xp = (v: number) => plot.x + axisFraction(X, v, false) * plot.w
      for (const s of group.series) {
        const pts: Array<readonly [number, number] | undefined> = []
        for (let i = 0; i < count; i++) {
          const v = plotted(s.value(i)), xv = s.hasX ? s.x(i) : i + 1
          pts.push(v === undefined || xv === undefined || !X.plottable(xv) ? undefined : [xp(xv), at(v)] as const)
        }
        ctx.strokeStyle = ctx.fillStyle = `#${s.color}`
        if (s.line) {
          ctx.lineWidth = 2; ctx.beginPath()
          let pen = false
          for (const p of pts) { if (!p) { pen = false; continue } if (!pen) { ctx.moveTo(p[0], p[1]); pen = true } else ctx.lineTo(p[0], p[1]) }
          ctx.stroke()
        }
        if (s.markers !== false) for (const p of pts) if (p) { ctx.beginPath(); ctx.arc(p[0], p[1], 3, 0, Math.PI * 2); ctx.fill() }
      }
    }
  })
  } finally { ctx.restore() }
}

/** Paint a chart into a w×h box at the origin of `ctx` (worksheet CSS px). */
export function paintXlsxChart(source: ChartModel, ctx: CanvasRenderingContext2D, w: number, h: number, familyCss: FamilyCss = JSON.stringify): void {
  if (source.unsupported || !source.groups.length) { paintChartPlaceholder(source, ctx, w, h, familyCss); return }
  const model = drawModel(source)
  ctx.save()
  try {
    ctx.beginPath(); ctx.rect(0, 0, w, h); ctx.clip()
    ctx.fillStyle = '#FFFFFF'; ctx.fillRect(0, 0, w, h)
    ctx.strokeStyle = GRID; ctx.lineWidth = 1; ctx.strokeRect(0.5, 0.5, w - 1, h - 1)
    const fontPx = model.fontSizePt * 96 / 72
    const pad = Math.min(10, w * 0.04, h * 0.04)
    let area: Box = { x: pad, y: pad, w: w - pad * 2, h: h - pad * 2 }
    if (model.title) {
      const titlePx = fontPx * 1.4
      ctx.font = font(titlePx, model.fontFamily, familyCss)
      ctx.fillStyle = TEXT; ctx.textAlign = 'center'; ctx.textBaseline = 'top'
      ctx.fillText(model.title, w / 2, area.y, area.w)
      area = { ...area, y: area.y + titlePx + 6, h: area.h - titlePx - 6 }
    }
    area = paintLegend(ctx, model, area, fontPx, model.fontFamily, familyCss)
    if (area.w <= 8 || area.h <= 8) return
    const group = model.groups[0]
    if (group.family === 'pie' || group.family === 'doughnut') paintPie(ctx, model, group, area)
    else paintCartesian(ctx, model, area, fontPx, model.fontFamily, familyCss)
  } finally { ctx.restore() }
}
