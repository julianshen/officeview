/**
 * Radar charts (spec 4.1; lazy chart chunk): `c:radarStyle` standard (lines),
 * marker (lines and markers; a series marker `symbol=none` hides them, as
 * Excel writes its plain radar) and filled (filled polygons). Categories run
 * clockwise from 12 o'clock; the value axis is radial, from the centre (its
 * minimum) to the outer ring (its maximum), with web gridlines at its ticks
 * and its labels up the 12 o'clock spoke. The category axis' gridlines are the
 * spokes. Non-value points follow the line row of the point matrix (spec 5.8):
 * a gap breaks the polygon, a span connects across, zero sits at the centre.
 *
 * Nothing per point is allocated; every path is admitted against the
 * painter's path work bound. Data labels on radar points are not drawn yet.
 */
import { attrs, getChildren, type XmlNode } from '../../core/xml'
import type { ChartFillStyle } from '../fills'
import type { ChartGroup, ChartModel, ChartSeries } from '../model'
import type { DisplayView } from '../display'
import { applyLineStyle, type ChartLineStyle } from '../style'
import { axisFraction, type AxisGraph, type ResolvedAxis } from '../axes'

const STYLES = ['standard', 'marker', 'filled'] as const
/** `c:radarStyle` (schema default standard). */
export function readRadarStyle(node: XmlNode): NonNullable<ChartGroup['radarStyle']> {
  const raw = attrs(getChildren(node, 'radarStyle')[0]).val
  return (STYLES as readonly string[]).includes(raw ?? '') ? raw as NonNullable<ChartGroup['radarStyle']> : 'standard'
}

/** The radial value axis and the category axis of a radar group. */
function radarAxes(group: Pick<ChartGroup, 'binding'>, graph: AxisGraph): { V: ResolvedAxis; C?: ResolvedAxis; categoryKey: string } | undefined {
  const binding = group.binding
  if (binding.kind !== 'radial') return undefined
  const V = graph.axes.get(binding.value)
  if (!V) return undefined
  const C = binding.category !== undefined ? graph.axes.get(binding.category) : undefined
  return { V, ...(C ? { C } : {}), categoryKey: binding.category ?? `${binding.value}:categories` }
}
const categoryLabel = (view: DisplayView, slot: number) => view.category(slot) ?? String(view.source(slot) + 1)

/**
 * The strings a radar draws, keyed like axis texts (layout measures them): the radial axis' tick labels
 * (`valueLabels`, the prepared or formatted ones) and the category labels around the web.
 */
export function radarTexts(model: ChartModel, graph: AxisGraph, view: DisplayView, valueLabels: (axis: ResolvedAxis) => string[]): Array<[string, string[]]> {
  const out: Array<[string, string[]]> = []
  for (const group of model.groups) {
    if (group.family !== 'radar') continue
    const axes = radarAxes(group, graph)
    if (!axes) continue
    out.push([axes.V.id, valueLabels(axes.V)])
    out.push([axes.categoryKey, Array.from({ length: view.slots }, (_, i) => categoryLabel(view, i))])
  }
  return out
}

interface Box { x: number; y: number; w: number; h: number }
/** What the radar painter reads of a drawn series (the painter's display series). */
export interface RadarSeries {
  source: ChartSeries
  color: string
  line?: boolean
  markers?: boolean
  plot(slot: number): number | 'gap' | 'span'
}
/** The painter's shared helpers, passed in (the family never imports the painter). */
export interface RadarPaintKit {
  plot: Box
  slots: number
  source(slot: number): number
  view: DisplayView
  fontSizePt: number
  font(px: number): string
  texts: Map<string, string[]>
  width(text: string, font: string): number
  admitPath(cost: number): boolean
  fillWith(ctx: CanvasRenderingContext2D, fill: ChartFillStyle | undefined, color: string, box: Box): boolean
  strokeBorder(ctx: CanvasRenderingContext2D, line: ChartLineStyle | undefined, color: string, box: Box): void
  drawMarker(ctx: CanvasRenderingContext2D, s: RadarSeries, sourceIndex: number, x: number, y: number, legacyRadius: number): void
}
const TEXT = '#595959', GRID = '#D9D9D9', AXIS = '#BFBFBF'

/** Paints one radar group: the web (gridlines, spokes, labels), then each series in order. */
export function paintRadar(ctx: CanvasRenderingContext2D, group: Pick<ChartGroup, 'binding' | 'radarStyle'>, series: readonly RadarSeries[], graph: AxisGraph, kit: RadarPaintKit): void {
  const axes = radarAxes(group, graph)
  if (!axes) return
  const { V, C, categoryKey } = axes
  const n = kit.slots, plot = kit.plot
  const categoryPx = (C?.axis.text?.sizePt ?? kit.fontSizePt) * 96 / 72, valuePx = (V.axis.text?.sizePt ?? kit.fontSizePt) * 96 / 72
  const categories = kit.texts.get(categoryKey) ?? [], ticks = kit.texts.get(V.id) ?? []
  const showCategories = !(C?.deleted ?? false) && C?.axis.tickLblPos !== 'none'
  const showValues = !V.deleted && V.axis.tickLblPos !== 'none'
  // The web leaves room for the category labels around it.
  let labelWidth = 0
  if (showCategories) for (const text of categories) labelWidth = Math.max(labelWidth, kit.width(text, kit.font(categoryPx)))
  const R = Math.max(1, Math.min(plot.w / 2 - (showCategories ? labelWidth + 8 : 4), plot.h / 2 - (showCategories ? categoryPx + 6 : 4)))
  const cx = plot.x + plot.w / 2, cy = plot.y + plot.h / 2
  const angle = (slot: number) => -Math.PI / 2 + 2 * Math.PI * slot / n
  const at = (slot: number, f: number): [number, number] => [cx + f * R * Math.cos(angle(slot)), cy + f * R * Math.sin(angle(slot))]
  const ring = (f: number) => {
    ctx.beginPath()
    if (n < 3) { ctx.arc(cx, cy, f * R, 0, Math.PI * 2); return }
    for (let i = 0; i < n; i++) { const [x, y] = at(i, f); if (i) ctx.lineTo(x, y); else ctx.moveTo(x, y) }
    ctx.closePath()
  }

  ctx.save()
  try {
    ctx.lineWidth = 1
    // Value-axis gridlines: a web ring at every major tick.
    if (V.axis.gridlines && kit.admitPath(V.ticks.length * (n + 1))) {
      ctx.strokeStyle = GRID
      for (const t of V.ticks) { ring(axisFraction(V, t)); ctx.stroke() }
    }
    // Category gridlines: the spokes; the radial axis line up the 12 o'clock spoke.
    if (C?.axis.gridlines && kit.admitPath(n * 2)) {
      ctx.strokeStyle = GRID
      ctx.beginPath()
      for (let i = 0; i < n; i++) { const [x, y] = at(i, 1); ctx.moveTo(cx, cy); ctx.lineTo(x, y) }
      ctx.stroke()
    }
    if (!V.deleted) { ctx.strokeStyle = AXIS; ctx.beginPath(); ctx.moveTo(cx, cy); ctx.lineTo(cx, cy - R); ctx.stroke() }
    ctx.fillStyle = TEXT
    if (showCategories) {
      ctx.font = kit.font(categoryPx)
      for (let i = 0; i < n; i++) {
        const text = categories[i]
        if (!text) continue
        const cos = Math.cos(angle(i)), sin = Math.sin(angle(i))
        ctx.textAlign = cos > 0.1 ? 'left' : cos < -0.1 ? 'right' : 'center'
        ctx.textBaseline = sin < -0.1 ? 'bottom' : sin > 0.1 ? 'top' : 'middle'
        ctx.fillText(text, cx + (R + 6) * cos, cy + (R + 6) * sin)
      }
    }
  } finally { ctx.restore() }

  const box = { x: cx - R, y: cy - R, w: 2 * R, h: 2 * R }
  /** A point's position, or its non-value plotting decision (a value off a log axis is a gap). */
  const point = (s: RadarSeries, slot: number): [number, number] | 'gap' | 'span' => {
    const v = s.plot(slot)
    if (typeof v !== 'number') return v
    return V.plottable(v) ? at(slot, axisFraction(V, v)) : 'gap'
  }
  for (const s of series) {
    if (!kit.admitPath(n * 2 + 2)) continue
    const src = s.source
    if (group.radarStyle === 'filled') {
      ctx.beginPath()
      let any = false, gapAt = -1
      for (let i = 0; i < n && gapAt < 0; i++) if (point(s, i) === 'gap') gapAt = i
      if (gapAt < 0) {
        // No gap: one polygon through every plotted point (a span joins across, as the line does).
        for (let i = 0; i < n; i++) { const p = point(s, i); if (typeof p === 'string') continue; if (any) ctx.lineTo(p[0], p[1]); else ctx.moveTo(p[0], p[1]); any = true }
        if (!any) continue
        ctx.closePath()
      } else {
        // A gap breaks the area: each run between gaps (walking once round from the first gap, so a run may wrap
        // past category 0) is its own region from the centre, never joined across the gap. A one-point run has no area.
        let first: [number, number] | undefined, open = false
        const end = () => { if (open) { ctx.closePath(); any = true } first = undefined; open = false }
        for (let k = 1; k <= n; k++) {
          const p = point(s, (gapAt + k) % n)
          if (p === 'gap') { end(); continue }
          if (p === 'span') continue
          if (!first) { first = p; continue }
          if (!open) { ctx.moveTo(cx, cy); ctx.lineTo(first[0], first[1]); open = true }
          ctx.lineTo(p[0], p[1])
        }
        end()
        if (!any) continue
      }
      if (kit.fillWith(ctx, src.fill, s.color, box)) ctx.fill()
      kit.strokeBorder(ctx, src.lineStyle, s.color, box)
      continue
    }
    if (s.line !== false) {
      ctx.save()
      try {
        ctx.strokeStyle = `#${s.color}`; ctx.lineWidth = 2.25; ctx.lineJoin = 'round'
        if (applyLineStyle(ctx, src.lineStyle, box)) {
          ctx.beginPath()
          let pen = false
          for (let i = 0; i < n; i++) {
            const p = point(s, i)
            if (p === 'gap') { pen = false; continue }
            if (p === 'span') continue
            if (pen) ctx.lineTo(p[0], p[1]); else { ctx.moveTo(p[0], p[1]); pen = true }
          }
          // Close the polygon: from the last run back to category 0, across spans, unless a gap intervenes.
          if (pen && n > 1) for (let i = 0; i < n; i++) {
            const p = point(s, i)
            if (p === 'gap') break
            if (p === 'span') continue
            ctx.lineTo(p[0], p[1]); break
          }
          ctx.stroke()
        }
      } finally { ctx.restore() }
    }
    if (s.markers) for (let i = 0; i < n; i++) { const p = point(s, i); if (typeof p !== 'string') kit.drawMarker(ctx, s, kit.source(i), p[0], p[1], 2.5) }
  }
  // The radial tick labels sit above the series (a filled radar never hides them).
  if (showValues) {
    ctx.save()
    try {
      ctx.fillStyle = TEXT; ctx.font = kit.font(valuePx); ctx.textAlign = 'right'; ctx.textBaseline = 'middle'
      V.ticks.forEach((t, i) => { if (ticks[i]) ctx.fillText(ticks[i], cx - 4, cy - axisFraction(V, t) * R) })
    } finally { ctx.restore() }
  }
}
