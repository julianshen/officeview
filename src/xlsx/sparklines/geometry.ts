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

/**
 * `sizePx` is the half-diagonal for diamonds, the half-side for squares, the radius for circles. A polyline's `band`
 * is the horizontal strip [y0, y1) its stroke is clipped to (lineClipRule).
 */
export type SparkPrimitive =
  | { kind: 'polyline'; points: Array<[number, number]>; color: string; widthPx: number; join: CanvasLineJoin; cap: CanvasLineCap; band?: { y0: number; y1: number } }
  | { kind: 'marker'; x: number; y: number; sizePx: number; color: string; shape: 'square' | 'circle' | 'diamond' }
  | { kind: 'rect'; x: number; y: number; w: number; h: number; color: string }
  | { kind: 'axis'; x0: number; x1: number; y: number; color: string; widthPx: number }

// Plot box insets, pt from the host rectangle's edges (docs/validation/xlsx-sparklines.md, `cellBoxRule`,
// `plotInsetLeft/Right/Top/Bottom`, Task 7 reference comparison). The host rectangle runs gridline to gridline and its
// first column/row is the cell's own left/top gridline, so the measured inner-edge insets gain that gridline: left
// 2.0 + 1 = 3.0 and top 4.08 + 1 = 5.08; right and bottom are measured from the next gridline, the rectangle's edge.
// Top 4.08 / bottom 3.06 (refit from 4.15 / 3.5 by the reference comparison): the only band in which max → row 5,
// min → row 56 and a symmetric midpoint → row 31 all hold (top ≥ 4, bottom > 3, top − bottom ≥ 1 from the inner
// edges). The insets' zoom dependence is unknowable from the 1 px/pt capture route; they are treated as fixed points.
/** @internal Exported for the Task 7 constant sweep; not part of the public API. */
export const INSETS = { left: 3.0, right: 2.0, top: 5.08, bottom: 3.06 }
/** 1 pt in logical (CSS) px. */
const PT_PX = 4 / 3
/** colourFallback: colorSeries absent → black. Every other element without its colour is simply not drawn. */
const SERIES_FALLBACK = '000000'

interface Pt { x: number; y: number }

/**
 * Pure geometry for one resolved sparkline: line paths plus the Stage 2 (plan Task 8) column and stacked (win/loss)
 * bars; other types return []. Clipping to the host cell (§5.1) and to the current tile is the painter's concern;
 * everything the measurements pin about shape — slot positions, vertical mapping, pixel rounding, bar widths, the
 * baseline and δ rules, stroke width, plot-box row clipping, marker placement and colour precedence — is decided here.
 */
export function sparklineGeometry(s: ResolvedSparkline, cell: SparkCell, scale: SparkScale): SparkPrimitive[] {
  const group = s.group
  const type = group.type
  // Stage 2 (Task 8): column and stacked (win/loss) bars; anything else → [] (spec §10: only staged types return
  // anything). Bars share the line's plot box, insets, tiny-cell collapse and the device-grid snap helpers
  // (barTinyCellRule = tinyCellRule: `col-tiny` measures 0 ink pixels).
  if (type !== 'line' && type !== 'column' && type !== 'stacked') return []
  const zoom = scale.zoom > 0 && Number.isFinite(scale.zoom) ? scale.zoom : 1
  const dpr = scale.dpr > 0 && Number.isFinite(scale.dpr) ? scale.dpr : 1
  const devPerPx = zoom * dpr
  // `cell` is in sheet logical coordinates, so snap against the global sheet device grid.
  // Viewport painting subtracts its tile origin only in the canvas transform, after this step.
  const toDevX = (logical: number): number => logical * devPerPx
  const toDevY = (logical: number): number => logical * devPerPx
  const toOutputX = (dev: number): number => dev / (scale.output === 'device' ? 1 : devPerPx)
  const toOutputY = (dev: number): number => dev / (scale.output === 'device' ? 1 : devPerPx)

  // tinyCellRule (and barTinyCellRule): when the insets swallow the cell the plot box collapses and nothing draws,
  // markers included.
  const left = cell.x + INSETS.left * PT_PX
  const right = cell.x + cell.w - INSETS.right * PT_PX
  const top = cell.y + INSETS.top * PT_PX
  const bottom = cell.y + cell.h - INSETS.bottom * PT_PX
  const plotW = right - left
  const plotH = bottom - top
  if (!(plotW > 0) || !(plotH > 0)) return []

  const n = s.values.length
  if (!n) return []

  /** colourFallback: colorSeries absent → black. Every other element without its colour is simply not drawn. */
  const seriesColor = group.colors.series ?? SERIES_FALLBACK
  // Computable points: values carry numbers (Task 4); defensively treat non-finite as absent like a null slot.
  const has = (i: number): boolean => i >= 0 && i < n && s.values[i] !== null && Number.isFinite(s.values[i] as number)
  // verticalMapping note: flat series (flatSeriesRule) widen the scale to [min(0, v), max(0, v)] — both paths use it.
  const flat = s.min >= s.max

  // lineStrokeWidth: whole device px, ≥ 1, round half up of the lineWeight in device px (the row pins the behaviour
  // at 1 px/pt; the scale-relative rounding is its extension). lineJoin/lineCap: round, round. The axis is 1 device
  // px (odd), so both paths snap like the line does.
  const strokeDev = Math.max(1, Math.floor(group.lineWeightPt * PT_PX * devPerPx + 0.5))
  const strokePx = strokeDev / (scale.output === 'device' ? 1 : devPerPx)
  const oddStroke = strokeDev % 2 === 1
  // `rtlRule` (Task 10): rightToLeft mirrors the x about the plot box BEFORE any snapping (x' = L + R − x), while
  // indices and colours stay in data order — a first point keeps its colour at the RIGHT edge of the box.
  const rtl = group.rightToLeft === true

  // vertexSnapRule: odd device width → pixel centre ⌊x⌋+½, ⌊y⌋+½; even width → pixel corner round(x), round(y).
  const snapVertex = (dev: number): number => oddStroke ? Math.floor(dev) + 0.5 : Math.round(dev)

  if (type !== 'line') {
    const stacked = type === 'stacked'
    const Ldev = toDevX(left), Rdev = toDevX(right), Tdev = toDevY(top), Bdev = toDevY(bottom)
    const Hdev = Bdev - Tdev
    const rowTop = Math.round(Tdev)
    const rowBottom = Math.round(Bdev)
    const midRow = Math.round(Tdev + Hdev / 2)
    // barSlotRule / barWidthRule: each of the n values owns one of n equal slots (lineEndpointRule's slots) and the
    // bar is centred on the slot centre; width = slot / 1.4 (a 40% gap between bars, half on each side); edges are
    // round(c − w/2) and round(c + w/2) in device px (measured: P = 199 → 17–18 px, P = 403 → 36 px exactly).
    // `dateSlotRule` (Task 10): with a resolved axis the slot is the group's day slot (axis.slot of the plot width)
    // and the centre is the dated day-slot centre `positions[i]` of the plot width; `rtlRule` mirrors the centre
    // about the plot box (measured edges reproduce the undated ones mirrored exactly; colours stay in data order).
    const slot = s.axis ? toDevX(plotW) * s.axis.slot : (Rdev - Ldev) / n
    const halfBar = slot / 1.4 / 2
    const centreDev = (i: number): number => {
      if (s.axis) {
        const frac = s.axis.positions[i] ?? (i + 0.5) / n
        return toDevX(rtl ? right - frac * plotW : left + frac * plotW)
      }
      const centre = Ldev + (i + 0.5) * slot
      return rtl ? Ldev + Rdev - centre : centre
    }

    // barColourPrecedence: each bar's colour is the last painted of series → negative → first/last → high → low
    // (high/low are the max/min values, win/loss too, not "all wins" — the resolve indices). Ties: every tied bar is
    // coloured. A missing element colour is not applied (colourFallback): colorNegative absent keeps the series
    // colour; colorSeries absent → black (shared seriesColor).
    const winners = new Map<number, string>()
    const colorize = (on: boolean, color: string | undefined, indices: Iterable<number>): void => {
      if (!on || !color) return
      for (const i of indices) if (has(i)) winners.set(i, color)
    }
    if (group.negative && group.colors.negative) {
      for (let i = 0; i < n; i++) {
        const v = s.values[i]
        if (v !== null && Number.isFinite(v) && (v as number) < 0) winners.set(i, group.colors.negative)
      }
    }
    // firstIndex/lastIndex are data-order slots; a null slot has no bar and no mark (Task 4 note).
    colorize(group.first, group.colors.first, [s.firstIndex])
    colorize(group.last, group.colors.last, [s.lastIndex])
    colorize(group.high, group.colors.high, [...s.highIndices])
    colorize(group.low, group.colors.low, [...s.lowIndices])

    // columnNoZeroRule (δ = 0.63 pt): when the effective scale excludes 0 the baseline is the plot edge on 0's side
    // and the scale's end nearest 0 maps δ inside it, so the smallest value keeps a sliver (1 px at 1 px/pt). δ is
    // the midpoint of the feasible sweep [0.45, 0.81] pt (unmeasured; underdetermined per the row's notes) and
    // applies for any scale type (ruling). columnScaleRule/verticalMapping otherwise: the single linear map with
    // H = plot height; columnBaselineRule: baseline row round(y(0)) when the scale includes 0.
    const deltaDev = 0.63 * PT_PX * devPerPx
    const yDevFor = (v: number, ymin: number, ymax: number): number => {
      if (ymin > 0) return Tdev + (ymax - v) / (ymax - ymin) * (Hdev - deltaDev)
      if (ymax < 0) return Tdev + deltaDev + (ymax - v) / (ymax - ymin) * (Hdev - deltaDev)
      return Tdev + (ymax - v) / (ymax - ymin) * Hdev
    }
    // Custom scales that cut bars (ruling 3): bar edges clamp to the plot box rows [round(T), round(B)).
    const clampRow = (row: number): number => Math.min(rowBottom, Math.max(rowTop, row))

    const primitives: SparkPrimitive[] = []
    for (let i = 0; i < n; i++) {
      // columnEmptyRule: gap, zero and span slots keep their positions and stay empty; a literal 0 sits at the
      // baseline → zero-height bar → nothing (the 0 takes part in min/max via the resolved scale, Task 4).
      if (!has(i)) continue
      const v = s.values[i] as number
      let barTop: number, barBottom: number
      if (stacked) {
        // winLossRule: a win bar runs [round(T), round(T + H/2)) and a loss bar [round(T + H/2), round(T + H)),
        // whatever the magnitudes. Exact sign test: v > 0 win, v < 0 loss, v = 0 nothing (±1e-16 draws full bars).
        if (v > 0) { barTop = rowTop; barBottom = midRow }
        else if (v < 0) { barTop = midRow; barBottom = rowBottom }
        else continue
      } else {
        // Flat scale (ruling 1, unmeasured): widen like flatSeriesRule to [min(0, v), max(0, v)] then apply the
        // baseline rules — v > 0 a full-height bar from the bottom, v < 0 full-height from the top, v = 0 nothing
        // (the widened scale [0, 0] collapses to the zero-height bar below).
        const ymin = flat ? Math.min(0, v) : s.min
        const ymax = flat ? Math.max(0, v) : s.max
        if (flat && ymin === ymax) continue
        const baseRow = ymin > 0 || ymax < 0
          ? (ymin > 0 ? rowBottom : rowTop)
          : Math.round(yDevFor(0, ymin, ymax))
        const valueRow = clampRow(Math.round(yDevFor(v, ymin, ymax)))
        barTop = Math.min(baseRow, valueRow)
        barBottom = Math.max(baseRow, valueRow)
      }
      if (!(barBottom > barTop)) continue // zero-height bar draws nothing (columnEmptyRule 'zero')
      const centre = centreDev(i)
      const barLeft = Math.round(centre - halfBar)
      const barRight = Math.round(centre + halfBar)
      if (!(barRight > barLeft)) continue
      primitives.push({
        kind: 'rect',
        x: toOutputX(barLeft),
        y: toOutputY(barTop),
        w: toOutputX(barRight - barLeft),
        h: toOutputY(barBottom - barTop),
        color: winners.get(i) ?? (group.negative && group.colors.negative && v < 0 ? group.colors.negative : seriesColor),
      })
    }

    // barAxisRule, painted OVER the bars (the line paints it under): the column axis is exactly the line's
    // (displayXAxis AND min < 0 < max of the effective scale — the all-positive column axis is unmeasured, ruling
    // 4); the win/loss axis is ALWAYS drawn when displayXAxis, at the middle row round(T + H/2), even all-zero.
    // 1 device px, colorAxis, floor(L) … floor(R) wide, built exactly like the line axis except its rounding: the
    // bar axis is always odd-width (1 device px) whatever the lineWeightPt, so its y snaps to the pixel centre
    // ⌊y⌋ + ½ independent of the line stroke's parity (the parity snap is the line path's own behaviour).
    if (group.displayXAxis && group.colors.axis && (stacked || (s.min < 0 && s.max > 0))) {
      const snapAxisY = (dev: number): number => Math.floor(dev) + 0.5
      const axisYDev = snapAxisY(stacked ? Tdev + Hdev / 2 : yDevFor(0, s.min, s.max))
      primitives.push({
        kind: 'axis',
        x0: toOutputX(Math.floor(Ldev) + 0.5),
        x1: toOutputX(Math.floor(Rdev) + 0.5),
        y: toOutputY(axisYDev),
        color: group.colors.axis,
        widthPx: scale.output === 'device' ? 1 : 1 / devPerPx,
      })
    }
    return primitives
  }

  // verticalMapping: y(v) = top + (max − v)/(max − min)·plotH, the single documented linear map (its rounding is
  // underdetermined at the ±1 px level; comparisons allow a ±1 px vertex tolerance). Flat series (flatSeriesRule):
  // the scale widens to [min(0, v), max(0, v)] — v ≠ 0 renders at the plot edge of its sign (v > 0 measured at the
  // top; v < 0 unmeasured and extrapolated to the bottom), v = 0 is centred.
  const yOf = (v: number): number => {
    if (!flat) return top + (s.max - v) / (s.max - s.min) * plotH
    if (v > 0) return top
    if (v < 0) return bottom
    return top + plotH / 2
  }

  // lineEndpointRule: n points sit at the centres of n equal slots across the plot box, never at its edges — the
  // undated slots keep the exact Task 5 expression (the pin at o.5 boundaries is knife-edge exact). `dateSlotRule`
  // (Task 10): with a resolved axis the x of point i is `s.axis.positions[i]` of the plot width — the centre of the
  // point's dated day slot — never evenly spaced; `rtlRule` mirrors the payoff side about the box in logical px.
  const xDevOf = (i: number): number => {
    const frac = s.axis ? s.axis.positions[i] ?? (i + 0.5) / n : undefined
    if (frac === undefined) {
      const slot = (i + 0.5) * plotW / n
      return toDevX(rtl ? right - slot : left + slot)
    }
    return toDevX(rtl ? right - frac * plotW : left + frac * plotW)
  }
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

  // lineClipRule / customClipRule, two levels (Task 7 reference comparison, line-scale-custom row 3 and
  // line-scale-custom-clip): out-of-range values are not dropped — the CENTRELINE runs through them and is cut where
  // it crosses the snapped rows of the plot top and bottom (the rows a max / min vertex snaps to: clip row 1 ink runs
  // exactly from the value-2 vertex to the value-6 vertex, and a segment leaving the min vertex downwards leaves no
  // ink beside it), and the STROKE is then clipped to the plot box rows [⌊top⌋, ⌊bottom⌋+1) in device px, which
  // cuts a wide stroke flat at the extremes (line-weight-2_25 row 4). Markers are NOT clipped.
  const clipTop = snapVertex(toDevY(top))
  const clipBottom = snapVertex(toDevY(bottom))
  const band = { y0: toOutputY(Math.floor(toDevY(top))), y1: toOutputY(Math.floor(toDevY(bottom)) + 1) }

  const primitives: SparkPrimitive[] = []

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
      if (current.length >= 2) primitives.push({ kind: 'polyline', points: current.map((p): [number, number] => [toOutputX(p.x), toOutputY(p.y)]), color: seriesColor, widthPx: strokePx, join: 'round', cap: 'round', band })
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
