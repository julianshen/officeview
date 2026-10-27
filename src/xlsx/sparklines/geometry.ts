/**
 * Sparkline line geometry (plan Task 5 of 2026-10-27-xlsx-sparklines): translates one ResolvedSparkline plus its host
 * cell rectangle and scale into drawing primitives, in logical px. Pure: no canvas, no DOM, and nothing eager may
 * import this module — Task 6's painter loads it lazily like the chart engine.
 *
 * Every rule cites its row of the "Measured constants" table in docs/validation/xlsx-sparklines.md. All measured
 * lengths are POINTS at 1 px/pt; the viewer's logical px are CSS px (1 pt = 4/3 CSS px), so a pt constant converts by
 * ×4/3 into logical px, and device px = logical px × zoom × dpr.
 */
import type { ResolvedSparkline } from './resolve'

export interface SparkCell { x: number; y: number; w: number; h: number }
export interface SparkScale {
  zoom: number
  dpr: number
  /** Internal painter mode: emit exact device-grid coordinates without a device→logical→device round trip. */
  output?: 'logical' | 'device'
}

/** `sizePx` is the half-diagonal for diamonds, the half-side for squares, the radius for circles. */
export type SparkPrimitive =
  | { kind: 'polyline'; points: Array<[number, number]>; color: string; widthPx: number; join: CanvasLineJoin; cap: CanvasLineCap }
  | { kind: 'marker'; x: number; y: number; sizePx: number; color: string; shape: 'square' | 'circle' | 'diamond' }
  | { kind: 'rect'; x: number; y: number; w: number; h: number; color: string }
  | { kind: 'axis'; x0: number; x1: number; y: number; color: string; widthPx: number }

// Plot box insets, pt from the cell edges (plotInsetLeft/Right 2.0 pt ±0.06; top 4.15 pt, bottom 3.5 pt — asymmetric:
// top is about 0.6 pt more than bottom). The insets' zoom dependence is unknowable from the 1 px/pt capture route;
// they are treated as fixed points.
const INSETS = { left: 2.0, right: 2.0, top: 4.15, bottom: 3.5 }
/** 1 pt in logical (CSS) px. */
const PT_PX = 4 / 3
/** colourFallback: colorSeries absent → black. Every other element without its colour is simply not drawn. */
const SERIES_FALLBACK = '000000'

interface Pt { x: number; y: number }

/**
 * Pure geometry for one resolved line sparkline (Task 8 adds column and stacked; other types return []).
 * Clipping to the host cell (§5.1) and to the current tile is the painter's concern; everything the measurements
 * pin about shape — slot positions, vertical mapping, pixel snapping, stroke width, plot-box row clipping,
 * marker placement and colour precedence — is decided here.
 */
export function sparklineGeometry(s: ResolvedSparkline, cell: SparkCell, scale: SparkScale): SparkPrimitive[] {
  const group = s.group
  if (group.type !== 'line') return []
  const zoom = scale.zoom > 0 && Number.isFinite(scale.zoom) ? scale.zoom : 1
  const dpr = scale.dpr > 0 && Number.isFinite(scale.dpr) ? scale.dpr : 1
  const devPerPx = zoom * dpr
  // `cell` is in sheet logical coordinates, so snap against the global sheet device grid.
  // Viewport painting subtracts its tile origin only in the canvas transform, after this step.
  const toDevX = (logical: number): number => logical * devPerPx
  const toDevY = (logical: number): number => logical * devPerPx
  const toOutputX = (dev: number): number => dev / (scale.output === 'device' ? 1 : devPerPx)
  const toOutputY = (dev: number): number => dev / (scale.output === 'device' ? 1 : devPerPx)

  // tinyCellRule: when the insets swallow the cell the plot box collapses and nothing draws, markers included.
  const left = cell.x + INSETS.left * PT_PX
  const right = cell.x + cell.w - INSETS.right * PT_PX
  const top = cell.y + INSETS.top * PT_PX
  const bottom = cell.y + cell.h - INSETS.bottom * PT_PX
  const plotW = right - left
  const plotH = bottom - top
  if (!(plotW > 0) || !(plotH > 0)) return []

  const n = s.values.length
  if (!n) return []

  // verticalMapping: y(v) = top + (max − v)/(max − min)·plotH, the single documented linear map (its rounding is
  // underdetermined at the ±1 px level; comparisons allow a ±1 px vertex tolerance). Flat series (flatSeriesRule):
  // the scale widens to [min(0, v), max(0, v)] — v ≠ 0 renders at the plot edge of its sign (v > 0 measured at the
  // top; v < 0 unmeasured and extrapolated to the bottom), v = 0 is centred.
  const flat = s.min >= s.max
  const yOf = (v: number): number => {
    if (!flat) return top + (s.max - v) / (s.max - s.min) * plotH
    if (v > 0) return top
    if (v < 0) return bottom
    return top + plotH / 2
  }

  // lineStrokeWidth: whole device px, ≥ 1, round half up of the lineWeight in device px (the row pins the behaviour
  // at 1 px/pt; the scale-relative rounding is its extension). lineJoin/lineCap: round, round.
  const strokeDev = Math.max(1, Math.floor(group.lineWeightPt * PT_PX * devPerPx + 0.5))
  const strokePx = strokeDev / (scale.output === 'device' ? 1 : devPerPx)
  const oddStroke = strokeDev % 2 === 1

  // vertexSnapRule: odd device width → pixel centre ⌊x⌋+½, ⌊y⌋+½; even width → pixel corner round(x), round(y).
  const snapVertex = (dev: number): number => oddStroke ? Math.floor(dev) + 0.5 : Math.round(dev)

  // Computable points: values carry numbers (Task 4); defensively treat non-finite as absent like a null slot.
  const has = (i: number): boolean => i >= 0 && i < n && s.values[i] !== null && Number.isFinite(s.values[i] as number)
  // lineEndpointRule: n points sit at the centres of n equal slots across the plot box, never at its edges.
  const xDevOf = (i: number): number => toDevX(left + (i + 0.5) * plotW / n)
  const yDevOf = (i: number): number => toDevY(yOf(s.values[i] as number))

  // Colour precedence (colourPrecedence): a point's colour is the last painted of markers → first/last → high → low
  // (low beats high and both beat first, last and markers). first vs last is undetermined — they coincide only when
  // n = 1, where low wins anyway — so the order below takes last over first. An element whose colour is missing is
  // not drawn at all (colourFallback), which demotes the point to the next decorated element.
  const drawable = (i: number): boolean => has(i)
  interface Decoration { color: string; special: boolean }
  const winners = new Map<number, Decoration>()
  const apply = (on: boolean, color: string | undefined, indices: number[], special: boolean): void => {
    if (!on || !color) return
    for (const i of indices) if (drawable(i)) winners.set(i, { color, special })
  }
  const all: number[] = []
  for (let i = 0; i < n; i++) if (drawable(i)) all.push(i)
  apply(group.markers, group.colors.markers, all, false)
  // firstIndex/lastIndex are data-order slots; a null slot (leading/trailing empty cell) has no vertex and no mark.
  apply(group.first, group.colors.first, [s.firstIndex], true)
  apply(group.last, group.colors.last, [s.lastIndex], true)
  apply(group.high, group.colors.high, [...s.highIndices], true)
  apply(group.low, group.colors.low, [...s.lowIndices], true)

  // Line runs: consecutive drawable points connect iff every null slot between them is a bridge (cellValueRule:
  // errors bridge even under gap; spanRule: interior spans bridge, leading/trailing empties do not; gapRule: gaps
  // break). Slot positions are kept, so bridged chords skip the empty slot without closing it up.
  const bridged = new Set(s.bridged)
  const firstDrawable = all[0], lastDrawable = all[all.length - 1]
  const isBridge = (i: number): boolean =>
    bridged.has(i) || (group.displayEmptyCellsAs === 'span' && i > firstDrawable && i < lastDrawable)
  const runs: number[][] = []
  let run: number[] = []
  for (let k = 0; k < all.length; k++) {
    const i = all[k]
    if (run.length && k > 0) {
      const previous = all[k - 1]
      let connected = true
      for (let j = previous + 1; j < i; j++) if (!isBridge(j)) { connected = false; break }
      if (!connected) { runs.push(run); run = [] }
    }
    run.push(i)
  }
  if (run.length) runs.push(run)

  // lineClipRule: strokes clip to the plot box rows [⌊top⌋, ⌊bottom⌋+1) in device px (markers are NOT clipped);
  // customClipRule: out-of-range values are not dropped — the polyline runs through them and is cut where it
  // crosses the band edges.
  const clipTop = Math.floor(toDevY(top))
  const clipBottom = Math.floor(toDevY(bottom)) + 1

  const primitives: SparkPrimitive[] = []
  const seriesColor = group.colors.series ?? SERIES_FALLBACK

  // paintOrder: axis → series line → markers → special points. Special points share the marker shape (the painter
  // draws one primitive per point), so the array order below preserves the paint order.
  // axisRule: drawn only when displayXAxis and min < 0 < max of the effective scale, at y(0), exactly 1 device px
  // wide (its width against lineWeight is underdetermined), across the whole plot box from the first to the last
  // device pixel column.
  if (group.displayXAxis && s.min < 0 && s.max > 0 && group.colors.axis) {
    const yDev = snapVertex(toDevY(yOf(0)))
    primitives.push({
      kind: 'axis',
      x0: toOutputX(Math.floor(toDevX(left)) + 0.5),
      x1: toOutputX(Math.floor(toDevX(right)) + 0.5),
      y: toOutputY(yDev),
      color: group.colors.axis,
      widthPx: scale.output === 'device' ? 1 : 1 / devPerPx,
    })
  }

  const samePt = (a: Pt, b: Pt): boolean => Math.abs(a.x - b.x) < 1e-9 && Math.abs(a.y - b.y) < 1e-9
  /** Clip the device-px segment against the plot box rows; returns the surviving endpoints, undefined if none. */
  const clipToBand = (a: Pt, b: Pt): [Pt, Pt] | undefined => {
    const dy = b.y - a.y
    let t0 = 0, t1 = 1
    if (Math.abs(dy) < 1e-9) {
      if (a.y < clipTop || a.y > clipBottom) return undefined
    } else {
      const lo = (clipTop - a.y) / dy, hi = (clipBottom - a.y) / dy
      const enter = Math.max(0, Math.min(lo, hi)), exit = Math.min(1, Math.max(lo, hi))
      if (enter > exit) return undefined
      t0 = enter; t1 = exit
    }
    const p = { x: a.x + (b.x - a.x) * t0, y: a.y + dy * t0 }
    const q = { x: a.x + (b.x - a.x) * t1, y: a.y + dy * t1 }
    return samePt(p, q) ? undefined : [p, q]
  }

  for (const indices of runs) {
    if (indices.length < 2) continue
    let current: Array<{ x: number; y: number }> = []
    const flush = (): void => {
      // Snapping and clipping happen on the device grid (clipToBand works in device px); anything reaching the
      // primitive converts back to logical px, exactly like the marker and axis primitives below.
      if (current.length >= 2) primitives.push({ kind: 'polyline', points: current.map((p): [number, number] => [toOutputX(p.x), toOutputY(p.y)]), color: seriesColor, widthPx: strokePx, join: 'round', cap: 'round' })
      current = []
    }
    for (let k = 0; k + 1 < indices.length; k++) {
      const a = { x: snapVertex(xDevOf(indices[k])), y: snapVertex(yDevOf(indices[k])) }
      const b = { x: snapVertex(xDevOf(indices[k + 1])), y: snapVertex(yDevOf(indices[k + 1])) }
      const piece = clipToBand(a, b)
      if (!piece) { flush(); continue }
      if (current.length && samePt(current[current.length - 1], piece[0])) current.push(piece[1])
      else { flush(); current = piece }
    }
    flush()
  }

  // markerSize: half-diagonal = (4/3)·lineWeight in pt (1 pt at 0.75, 3 pt at 2.25). Chosen over the
  // 2·strokePx − 1 device formula (both fit the two measured rows) for zoom stability; specialPointSize
  // confirms high/low/first/last use the same diamond.
  const markerHalfDiag = group.lineWeightPt * PT_PX * PT_PX
  const winnersSorted = [...winners.entries()].sort((a, b) => a[0] - b[0])
  for (const [i, decoration] of winnersSorted) {
    // vertexSnapRule: ordinary markers snap their centre to round(x), round(y); high/low/first/last points to
    // ⌊x⌋, ⌊y⌋ — 1 device px toward the start of the row/column versus an ordinary marker.
    const snap = decoration.special ? Math.floor : Math.round
    primitives.push({
      kind: 'marker',
      x: toOutputX(snap(xDevOf(i))),
      y: toOutputY(snap(yDevOf(i))),
      sizePx: markerHalfDiag * (scale.output === 'device' ? devPerPx : 1),
      color: decoration.color,
      shape: 'diamond',
    })
  }
  return primitives
}
