/**
 * Chart painter (lazy chunk): the common chart types (bar/column
 * clustered/stacked/100%, line, area, pie, doughnut, scatter). Anything else
 * paints the eager placeholder card. Moved from `src/xlsx/chart.ts` (A0a)
 * behaviour-identically.
 */
import { paintChartPlaceholder } from '../drawing/content-paint'
import { formatNumber } from '../xlsx/chartDescription'
import type { XlsxChartAxis, XlsxChartGroup, XlsxChartModel, XlsxChartSeries } from '../xlsx/chartModel'

type FamilyCss = (family: string) => string
function niceScale(lo: number, hi: number, axis: XlsxChartAxis, percent: boolean): { min: number; max: number; step: number } {
  if (percent) return { min: axis.min ?? 0, max: axis.max ?? 1, step: axis.majorUnit ?? 0.2 }
  let min = axis.min ?? Math.min(0, lo), max = axis.max ?? Math.max(0, hi)
  if (!(max > min)) max = min + 1
  const raw = (max - min) / 5
  const base = 10 ** Math.floor(Math.log10(raw))
  const step = axis.majorUnit && axis.majorUnit > 0 ? axis.majorUnit : [1, 2, 5, 10].map(m => m * base).find(m => m >= raw) ?? base * 10
  if (axis.min === undefined) min = Math.floor(min / step) * step
  if (axis.max === undefined) max = Math.ceil(max / step) * step
  if (!(max > min)) max = min + step
  return { min, max, step }
}
const font = (sizePx: number, family: string, familyCss: FamilyCss, bold = false) => `${bold ? 'bold ' : ''}${sizePx}px ${familyCss(family)}`
const TEXT = '#595959', GRID = '#D9D9D9'

interface Box { x: number; y: number; w: number; h: number }
function legendEntries(model: XlsxChartModel): Array<{ name: string; color: string }> {
  const group = model.groups[0]
  if (group && (group.type === 'pie' || group.type === 'doughnut')) {
    const s = group.series[0]
    return (s?.values ?? []).map((_, i) => ({ name: model.categories[i] ?? String(i + 1), color: s.pointColors?.[i] ?? s.color })).slice(0, 64)
  }
  return model.groups.flatMap(g => g.series).map(s => ({ name: s.name, color: s.color }))
}

function paintLegend(ctx: CanvasRenderingContext2D, model: XlsxChartModel, area: Box, fontPx: number, family: string, familyCss: FamilyCss): Box {
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

function paintPie(ctx: CanvasRenderingContext2D, model: XlsxChartModel, group: XlsxChartGroup, box: Box): void {
  const series = group.series[0]
  if (!series) return
  const values = series.values.map(v => v !== undefined && v > 0 ? v : 0)
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
    ctx.fillStyle = `#${series.pointColors?.[i] ?? series.color}`
    ctx.fill()
    ctx.strokeStyle = '#FFFFFF'; ctx.lineWidth = 1; ctx.stroke()
    angle += sweep
  })
  if (group.type === 'doughnut') {
    ctx.beginPath()
    ctx.arc(cx, cy, r * (group.holeSize ?? 50) / 100, 0, Math.PI * 2)
    ctx.fillStyle = '#FFFFFF'
    ctx.fill()
  }
  void model
}

function paintCartesian(ctx: CanvasRenderingContext2D, model: XlsxChartModel, box: Box, fontPx: number, family: string, familyCss: FamilyCss): void {
  const scatter = model.groups.every(g => g.type === 'scatter')
  const horizontal = model.groups.some(g => g.type === 'bar' && g.barDir === 'bar')
  const percent = model.groups.some(g => g.grouping === 'percentStacked')
  const count = Math.max(model.categories.length, ...model.groups.flatMap(g => g.series.map(s => s.values.length)), 1)
  // Value extents, honouring stacking.
  let lo = Infinity, hi = -Infinity
  const stackedValues = new Map<XlsxChartGroup, { pos: number[]; neg: number[]; total: number[] }>()
  for (const group of model.groups) {
    if (group.grouping === 'stacked' || group.grouping === 'percentStacked') {
      const pos = new Array<number>(count).fill(0), neg = new Array<number>(count).fill(0), total = new Array<number>(count).fill(0)
      for (const s of group.series) s.values.forEach((v, i) => { if (v === undefined || i >= count) return; if (v >= 0) pos[i] += v; else neg[i] += v; total[i] += Math.abs(v) })
      stackedValues.set(group, { pos, neg, total })
      if (group.grouping === 'stacked') for (let i = 0; i < count; i++) { lo = Math.min(lo, neg[i]); hi = Math.max(hi, pos[i]) }
    } else for (const s of group.series) for (const v of s.values) if (v !== undefined) { lo = Math.min(lo, v); hi = Math.max(hi, v) }
  }
  if (!Number.isFinite(lo)) { lo = 0; hi = 1 }
  const scale = niceScale(lo, hi, model.valueAxis, percent)
  let xs: { min: number; max: number; step: number } | undefined
  if (scatter) {
    let xlo = Infinity, xhi = -Infinity
    for (const s of model.groups.flatMap(g => g.series)) (s.xValues?.length ? s.xValues : s.values.map((_, i) => i + 1)).forEach(v => { if (v !== undefined) { xlo = Math.min(xlo, v); xhi = Math.max(xhi, v) } })
    if (!Number.isFinite(xlo)) { xlo = 0; xhi = 1 }
    xs = niceScale(xlo, xhi, model.categoryAxis, false)
  }
  ctx.font = font(fontPx, family, familyCss)
  const ticks: number[] = []
  for (let v = scale.min, n = 0; v <= scale.max + scale.step / 1e6 && n < 200; v += scale.step, n++) ticks.push(Number(v.toPrecision(12)))
  const label = (v: number) => formatNumber(v, percent)
  const valueLabelWidth = model.valueAxis.deleted ? 0 : Math.max(...ticks.map(t => ctx.measureText(label(t)).width)) + 6
  const catLabelHeight = model.categoryAxis.deleted ? 0 : fontPx + 6
  const valueTitle = model.valueAxis.title, catTitle = model.categoryAxis.title
  const titleBand = fontPx + 6
  let plot: Box = { ...box }
  if (horizontal) {
    const catWidth = model.categoryAxis.deleted ? 0 : Math.min(box.w * 0.35, Math.max(0, ...model.categories.map(c => ctx.measureText(c).width)) + 8)
    plot = { x: box.x + catWidth + (catTitle ? titleBand : 0), y: box.y + 4, w: box.w - catWidth - (catTitle ? titleBand : 0) - 8, h: box.h - catLabelHeight - (valueTitle ? titleBand : 0) - 4 }
  } else {
    plot = { x: box.x + valueLabelWidth + (valueTitle ? titleBand : 0), y: box.y + 4, w: box.w - valueLabelWidth - (valueTitle ? titleBand : 0) - 8, h: box.h - catLabelHeight - (catTitle ? titleBand : 0) - 4 }
  }
  if (plot.w <= 4 || plot.h <= 4) return
  const vpos = (v: number) => (Math.min(scale.max, Math.max(scale.min, v)) - scale.min) / (scale.max - scale.min)
  // Value-axis gridlines and labels.
  ctx.lineWidth = 1
  for (const t of ticks) {
    const f = vpos(t)
    ctx.strokeStyle = GRID
    if (model.valueAxis.gridlines || t === scale.min) {
      ctx.beginPath()
      if (horizontal) { const x = Math.round(plot.x + f * plot.w) + 0.5; ctx.moveTo(x, plot.y); ctx.lineTo(x, plot.y + plot.h) }
      else { const y = Math.round(plot.y + plot.h - f * plot.h) + 0.5; ctx.moveTo(plot.x, y); ctx.lineTo(plot.x + plot.w, y) }
      ctx.stroke()
    }
    if (model.valueAxis.deleted) continue
    ctx.fillStyle = TEXT
    if (horizontal) { ctx.textAlign = 'center'; ctx.textBaseline = 'top'; ctx.fillText(label(t), plot.x + f * plot.w, plot.y + plot.h + 3) }
    else { ctx.textAlign = 'right'; ctx.textBaseline = 'middle'; ctx.fillText(label(t), plot.x - 4, plot.y + plot.h - f * plot.h) }
  }
  const zero = vpos(Math.max(scale.min, Math.min(scale.max, 0)))
  // Category axis line and labels.
  const band = (horizontal ? plot.h : plot.w) / count
  if (!scatter) {
    ctx.strokeStyle = '#BFBFBF'
    ctx.beginPath()
    if (horizontal) { const x = Math.round(plot.x + zero * plot.w) + 0.5; ctx.moveTo(x, plot.y); ctx.lineTo(x, plot.y + plot.h) }
    else { const y = Math.round(plot.y + plot.h - zero * plot.h) + 0.5; ctx.moveTo(plot.x, y); ctx.lineTo(plot.x + plot.w, y) }
    ctx.stroke()
    if (!model.categoryAxis.deleted && model.categories.length) {
      ctx.fillStyle = TEXT
      const widest = Math.max(...model.categories.map(c => ctx.measureText(c).width))
      const skip = horizontal ? Math.max(1, Math.ceil((fontPx + 2) / band)) : Math.max(1, Math.ceil((widest + 6) / band))
      model.categories.forEach((c, i) => {
        if (i % skip) return
        // Horizontal bars list categories top-down in reverse (Excel's default orientation).
        if (horizontal) { ctx.textAlign = 'right'; ctx.textBaseline = 'middle'; ctx.fillText(c, plot.x - 4, plot.y + plot.h - band * (i + 0.5)) }
        else { ctx.textAlign = 'center'; ctx.textBaseline = 'top'; ctx.fillText(c, plot.x + band * (i + 0.5), plot.y + plot.h + 3, skip * band) }
      })
    }
  } else if (xs && !model.categoryAxis.deleted) {
    ctx.fillStyle = TEXT; ctx.textAlign = 'center'; ctx.textBaseline = 'top'
    for (let v = xs.min, n = 0; v <= xs.max + xs.step / 1e6 && n < 200; v += xs.step, n++) {
      const x = plot.x + (v - xs.min) / (xs.max - xs.min) * plot.w
      ctx.fillText(formatNumber(Number(v.toPrecision(12))), x, plot.y + plot.h + 3)
      if (model.categoryAxis.gridlines) { ctx.strokeStyle = GRID; ctx.beginPath(); ctx.moveTo(Math.round(x) + 0.5, plot.y); ctx.lineTo(Math.round(x) + 0.5, plot.y + plot.h); ctx.stroke() }
    }
  }
  // Axis titles. Every save is balanced in finally, so a throw never leaks state (spec 6.5 containment).
  ctx.fillStyle = TEXT
  const rotated = (text: string) => { ctx.save(); try { ctx.translate(box.x + fontPx, plot.y + plot.h / 2); ctx.rotate(-Math.PI / 2); ctx.fillText(text, 0, 0) } finally { ctx.restore() } }
  if (catTitle) { ctx.textAlign = 'center'; ctx.textBaseline = 'bottom'; if (horizontal) rotated(catTitle); else ctx.fillText(catTitle, plot.x + plot.w / 2, box.y + box.h) }
  if (valueTitle) { ctx.textAlign = 'center'; ctx.textBaseline = 'bottom'; if (horizontal) ctx.fillText(valueTitle, plot.x + plot.w / 2, box.y + box.h); else rotated(valueTitle) }

  ctx.save()
  try {
  ctx.beginPath(); ctx.rect(plot.x, plot.y, plot.w, plot.h); ctx.clip()
  const at = (f: number) => horizontal ? plot.x + f * plot.w : plot.y + plot.h - f * plot.h
  const barGroups = model.groups.filter(g => g.type === 'bar')
  for (const group of model.groups) {
    const stacked = stackedValues.get(group)
    const share = (v: number, i: number) => group.grouping === 'percentStacked' && stacked ? (stacked.total[i] ? v / stacked.total[i] : 0) : v
    if (group.type === 'bar') {
      const n = group.grouping === 'clustered' ? Math.max(1, group.series.length) : 1
      const stride = 1 - group.overlap / 100
      const span = 1 + (n - 1) * stride
      const width = band / (span + group.gapWidth / 100)
      const posBase = new Array<number>(count).fill(0), negBase = new Array<number>(count).fill(0)
      group.series.forEach((s, si) => {
        s.values.forEach((raw, i) => {
          if (raw === undefined || i >= count) return
          const v = share(raw, i)
          let from = 0, to = v
          if (stacked) { const base = v >= 0 ? posBase : negBase; from = base[i]; to = base[i] + v; base[i] = to }
          const slot = group.grouping === 'clustered' ? si : 0
          const offset = (band - width * span) / 2 + width * slot * stride
          const a = at(vpos(from)), b = at(vpos(to))
          ctx.fillStyle = `#${s.pointColors?.[i] && barGroups.length === 1 && group.series.length === 1 ? s.pointColors[i] : s.color}`
          if (horizontal) ctx.fillRect(Math.min(a, b), plot.y + plot.h - band * (i + 1) + offset, Math.abs(b - a), width)
          else ctx.fillRect(plot.x + band * i + offset, Math.min(a, b), width, Math.abs(b - a))
        })
      })
    } else if (group.type === 'line' || group.type === 'area') {
      const base = new Array<number>(count).fill(0)
      const edge = group.type === 'area'
      const xAt = (i: number) => edge ? plot.x + (count > 1 ? i * plot.w / (count - 1) : plot.w / 2) : plot.x + band * (i + 0.5)
      const layers: Array<{ s: XlsxChartSeries; top: Array<number | undefined>; bottom: number[] }> = []
      for (const s of group.series) {
        const bottom = base.slice()
        const top = Array.from({ length: count }, (_, i) => {
          const raw = s.values[i]
          if (raw === undefined) return stacked && edge ? base[i] : undefined
          const v = share(raw, i)
          if (stacked) { base[i] += v; return base[i] }
          return v
        })
        layers.push({ s, top, bottom })
      }
      for (const { s, top, bottom } of layers) {
        ctx.strokeStyle = ctx.fillStyle = `#${s.color}`
        if (edge) {
          ctx.beginPath()
          let started = false
          top.forEach((v, i) => { if (v === undefined) return; const x = xAt(i), y = at(vpos(v)); if (!started) { ctx.moveTo(x, y); started = true } else ctx.lineTo(x, y) })
          if (!started) continue
          for (let i = count - 1; i >= 0; i--) if (top[i] !== undefined) ctx.lineTo(xAt(i), at(vpos(stacked ? bottom[i] : Math.max(scale.min, Math.min(scale.max, 0)))))
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
          top.forEach((v, i) => { if (v === undefined) { pen = false; return } const x = xAt(i), y = at(vpos(v)); if (!pen) { ctx.moveTo(x, y); pen = true } else ctx.lineTo(x, y) })
          ctx.stroke()
        }
        if (s.markers) top.forEach((v, i) => { if (v === undefined) return; ctx.beginPath(); ctx.arc(xAt(i), at(vpos(v)), 2.5, 0, Math.PI * 2); ctx.fill() })
      }
    } else if (group.type === 'scatter' && xs) {
      const xp = (v: number) => plot.x + (v - xs!.min) / (xs!.max - xs!.min) * plot.w
      for (const s of group.series) {
        const pts = s.values.map((v, i) => { const x = s.xValues?.length ? s.xValues[i] : i + 1; return v === undefined || x === undefined ? undefined : [xp(x), at(vpos(v))] as const })
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
  }
  } finally { ctx.restore() }
}

/** Paint a chart into a w×h box at the origin of `ctx` (worksheet CSS px). */
export function paintXlsxChart(model: XlsxChartModel, ctx: CanvasRenderingContext2D, w: number, h: number, familyCss: FamilyCss = JSON.stringify): void {
  if (model.unsupported || !model.groups.length) { paintChartPlaceholder(model, ctx, w, h, familyCss); return }
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
    if (group.type === 'pie' || group.type === 'doughnut') paintPie(ctx, model, group, area)
    else paintCartesian(ctx, model, area, fontPx, model.fontFamily, familyCss)
  } finally { ctx.restore() }
}
