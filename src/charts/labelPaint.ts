/** Placement of prepared data-label strings. Formatting never happens during paint. */
import { barGeometry } from './barGeometry'
import type { ChartModel } from './model'
import type { ChartLayout } from './layout'
import { labelFont } from './textStyle'
import { axisFraction, categoryFraction } from './axes'
import { valueAt, hiddenAt } from './points'
import { ofPieDisplayValue, ofPiePlan, ofPiePointPlacement, type OfPiePlan } from './families/ofPie'
export function paintDataLabels(model: ChartModel, ctx: CanvasRenderingContext2D, layout: ChartLayout, family: string): void {
  const labels = model.preparedAxes?.dataLabels
  if (!labels?.length) return
  const { view, plot } = layout, slots = new Map<number, number>()
  for (let i = 0; i < view.slots; i++) slots.set(view.source(i), i)
  const along = (a: NonNullable<(typeof layout.graph.groups)[number]>['value'], v: number) => a.orientation === 'h' ? plot.x + axisFraction(a, v) * plot.w : plot.y + plot.h - axisFraction(a, v) * plot.h
  let pieSeries: object | undefined, pieBefore = 0, pieIndex = 0, pieTotal = 0
  let ofPieState: { g: object; s: object; plan: OfPiePlan } | undefined
  for (const label of labels) {
    const g = model.groups[label.group], s = g.series[label.series]
    // The ofPie Other pseudo label's point (n) is not a displayed source, so it bypasses the slot
    // lookup and the hidden check entirely.
    const other = label.other === true, slot = other ? -1 : slots.get(label.point)
    if (slot === undefined || (!other && model.plotVisOnly !== false && (hiddenAt(s.values.hidden, label.point) || hiddenAt(s.xValues?.hidden, label.point)))) continue
    const raw = valueAt(s.values, label.point) ?? 0
    let x = 0, y = 0, anchorX = 0, anchorY = 0, outside = false
    if (g.family === 'pie' || g.family === 'doughnut') {
      if (pieSeries !== s) {
        pieSeries = s
        pieBefore = 0
        pieIndex = 0
        pieTotal = 0
        for (let i = 0; i < view.slots; i++) pieTotal += Math.max(0, valueAt(s.values, view.source(i)) ?? 0)
      }
      while (pieIndex < slot) { pieBefore += Math.max(0, valueAt(s.values, view.source(pieIndex++)) ?? 0) }
      if (!(pieTotal > 0) || raw <= 0) continue
      const box = layout.area, r = Math.max(1, Math.min(box.w, box.h) / 2 - 4), cx = box.x + box.w / 2, cy = box.y + box.h / 2
      const angle = ((g.firstSliceAngle ?? 0) - 90) * Math.PI / 180 + (pieBefore + raw / 2) / pieTotal * Math.PI * 2
      outside = label.position === 'outEnd' || (label.position === 'bestFit' && raw / pieTotal < .12)
      anchorX = cx + Math.cos(angle) * r; anchorY = cy + Math.sin(angle) * r
      x = cx + Math.cos(angle) * (outside ? r + 12 : r * .65); y = cy + Math.sin(angle) * (outside ? r + 12 : r * .65)
    } else if (g.family === 'ofPie') {
      // ofPie placement shares the painter's plan (the split, geometry and slice angles, computed
      // once per group and series, like the pie's running totals). A label's point is its source
      // index in the display; the pseudo point n (the synthetic Other slice) places on the primary
      // region via its -1 marker.
      let state = ofPieState
      if (state === undefined || state.g !== g || state.s !== s) {
        const ds = view.groupSeries(label.group).find(d => d.series === s)
        const values: number[] = []
        for (let i = 0; i < view.slots; i++) values.push(ofPieDisplayValue(ds?.value(i)))
        state = ofPieState = { g: g, s: s, plan: ofPiePlan(g, values, layout.area, slot => view.source(slot)) }
      }
      const placement = ofPiePointPlacement(state.plan, slot)
      if (!placement || !(placement.fraction > 0)) continue
      outside = label.position === 'outEnd' || (label.position === 'bestFit' && placement.fraction < .12)
      if (placement.bar) {
        anchorX = placement.cx - placement.r; anchorY = placement.cy
        x = placement.cx + placement.r * .3; y = placement.cy
      } else {
        anchorX = placement.cx + Math.cos(placement.angle) * placement.r
        anchorY = placement.cy + Math.sin(placement.angle) * placement.r
        x = placement.cx + Math.cos(placement.angle) * (outside ? placement.r + 12 : placement.r * .65)
        y = placement.cy + Math.sin(placement.angle) * (outside ? placement.r + 12 : placement.r * .65)
      }
    } else {
      const a = layout.graph.groups[label.group]
      if (!a) continue
      let v = raw, base = 0
      if (g.grouping === 'stacked' || g.grouping === 'percentStacked') {
        let total = 0
        for (let j = 0; j < g.series.length; j++) {
          const n = valueAt(g.series[j].values, label.point) ?? 0
          total += Math.abs(n)
          if (j < label.series && Math.sign(n) === Math.sign(raw)) base += n
        }
        if (g.grouping === 'percentStacked') { v = total ? v / total : 0; base = total ? base / total : 0 }
      }
      const fraction = a.category ? categoryFraction(a.category, layout.graph.categoryIndex(a.category, slot), a.between) : axisFraction(a.x!, valueAt(s.xValues, label.point) ?? slot + 1)
      const category = a.horizontal === a.value ? plot.y + plot.h - fraction * plot.h : plot.x + fraction * plot.w
      const start = along(a.value, base), end = along(a.value, base + v)
      let position = end
      if (label.position === 'ctr') position = (start + end) / 2
      else if (label.position === 'inBase') position = start + (end >= start ? 6 : -6)
      else if (label.position === 'inEnd') position = end + (end >= start ? -6 : 6)
      else if (label.position === 'outEnd') position = end + (end >= start ? 8 : -8)
      if (a.value.orientation === 'h') { x = position; y = category } else { x = category; y = position }
      if (g.family === 'bar' && a.category) {
        const geometry = barGeometry(g, a.category, a.between, plot)
        const center = geometry.start(layout.graph.categoryIndex(a.category, slot), label.series) + geometry.width / 2
        if (a.category.orientation === 'v') y = center
        else x = center
      }
      if (label.position === 't') y -= 8
      if (label.position === 'b') y += 8
      if (label.position === 'l') x -= 8
      if (label.position === 'r') x += 8
    }
    ctx.save()
    if (outside && label.leaderLines) {
      ctx.strokeStyle = '#595959'
      ctx.lineWidth = 1
      ctx.beginPath()
      ctx.moveTo(anchorX, anchorY)
      ctx.lineTo(x, y)
      ctx.stroke()
    }
    ctx.textBaseline = 'middle'; ctx.textAlign = 'left'
    const runs = label.runs.length ? label.runs : [{ text: label.text }]
    let width = label.legendKey ? 12 : 0
    for (const run of runs) width += layout.width(run.text, labelFont(run, layout.fontPx, family))
    let at = x - width / 2
    if (label.legendKey) { ctx.fillStyle = '#' + s.color; ctx.fillRect(at, y - 4, 8, 8); at += 12 }
    for (const run of runs) {
      ctx.font = labelFont(run, layout.fontPx, family)
      ctx.fillStyle = run.color ?? '#595959'
      ctx.fillText(run.text, at, y)
      at += layout.width(run.text, ctx.font)
    }
    ctx.restore()
  }
}
