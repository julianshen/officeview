// @vitest-environment node
/**
 * Sparkline line geometry (plan Task 5). Every expected number below is computed by hand from the referenced row of
 * the "Measured constants" table in docs/validation/xlsx-sparklines.md. The table is measured in POINTS at 1 px/pt;
 * the viewer's logical px are CSS px (1 pt = 4/3 CSS px), so a pt constant becomes ×4/3 in logical px.
 */
import { describe, expect, it } from 'vitest'
import { sparklineGeometry, type SparkPrimitive } from '../src/xlsx/sparklines/geometry'
import type { ResolvedSparkline } from '../src/xlsx/sparklines/resolve'
import type { XlsxSparklineGroup } from '../src/xlsx/sparklines/model'

/** A 100×40 logical px cell at (20, 30): 75×30 pt. */
const CELL = { x: 20, y: 30, w: 100, h: 40 }
const Z1 = { zoom: 1, dpr: 1 }
/** The stroke clip band of CELL at Z1 (lineClipRule): rows [⌊T⌋, ⌊B⌋+1). */
const BAND = { y0: 36, y1: 66 }

const groupOf = (over: Partial<XlsxSparklineGroup> = {}): XlsxSparklineGroup => ({
  type: 'line', lineWeightPt: 0.75, displayEmptyCellsAs: 'zero',
  markers: false, high: false, low: false, first: false, last: false, negative: false,
  displayXAxis: false, displayHidden: false, dateAxis: false, rightToLeft: false,
  minAxisType: 'individual', maxAxisType: 'individual', colors: {}, sparklines: [],
  ...over,
})

/** A ResolvedSparkline as Task 4 resolves it: high/low computed over non-null values (`tieRule`: all ties). */
const mk = (values: Array<number | null>, over: Partial<XlsxSparklineGroup> = {}, resolved: Partial<ResolvedSparkline> = {}): ResolvedSparkline => {
  const nonNull = values.filter((v): v is number => v !== null)
  const min = resolved.min ?? Math.min(...nonNull)
  const max = resolved.max ?? Math.max(...nonNull)
  const high: number[] = [], low: number[] = []
  values.forEach((v, i) => {
    if (v === null) return
    if (v === max) high.push(i)
    if (v === min) low.push(i)
  })
  return {
    host: { row: 0, col: 0 }, dataSheet: 'Sheet1', sourceRows: values.map((_, index) => index), group: groupOf(over), values, bridged: resolved.bridged ?? [],
    firstIndex: resolved.firstIndex ?? 0, lastIndex: resolved.lastIndex ?? values.length - 1,
    highIndices: resolved.highIndices ?? high, lowIndices: resolved.lowIndices ?? low, min, max,
  }
}

type Poly = Extract<SparkPrimitive, { kind: 'polyline' }>
type Marker = Extract<SparkPrimitive, { kind: 'marker' }>
type Axis = Extract<SparkPrimitive, { kind: 'axis' }>
const polylines = (p: SparkPrimitive[]): Poly[] => p.filter((x): x is Poly => x.kind === 'polyline')
const markers = (p: SparkPrimitive[]): Marker[] => p.filter((x): x is Marker => x.kind === 'marker')
const axes = (p: SparkPrimitive[]): Axis[] => p.filter((x): x is Axis => x.kind === 'axis')

/**
 * Plot box for CELL, from the gridline-to-gridline host rectangle (cellBoxRule) with plotInsetLeft 2.0 pt + the 1-pt
 * left gridline = 3.0 pt, plotInsetRight 2.0 pt, plotInsetTop 4.08 pt + the 1-pt top gridline = 5.08 pt and
 * plotInsetBottom 3.06 pt (refit by the Task 7 reference comparison), each ×4/3 into logical px:
 *   L = 20 + 4 = 24, R = 120 − 8/3 = 117.333…, plotW = 93.333… px (70 pt)
 *   T = 30 + 5.08·4/3 = 36.7733…, B = 70 − 3.06·4/3 = 65.92, plotH = 29.1466… px
 * Stroke clip band (lineClipRule): rows [⌊T⌋, ⌊B⌋+1) = [36, 66); centreline band [snap(T), snap(B)] = [36.5, 65.5].
 * Line endpoint x for n points (lineEndpointRule): x[i] = L + (i + ½)·plotW/n.
 * Vertex y (verticalMapping): y = T + (max − v)/(max − min)·plotH.
 * Snap (vertexSnapRule): 1 device px stroke (odd) → pixel centre ⌊x⌋+½.
 */
describe('line sparkline geometry', () => {
  it('plots value [1, 3, 2] through the slot centres of the measured plot box', () => {
    // x: L + (i+½)·93.3333/3 → 39.556, 70.667, 101.778 → snapped 39.5, 70.5, 101.5 (lineEndpointRule, vertexSnapRule)
    // y (min 1, max 3): [1] → 65.92 → 65.5; [3] → 36.773 → 36.5 (max → pixel row T); [2] → 51.347 → 51.5 (verticalMapping)
    expect(polylines(sparklineGeometry(mk([1, 3, 2], { colors: { series: 'FF0000FF' } }), CELL, Z1))).toEqual([
      { kind: 'polyline', points: [[39.5, 65.5], [70.5, 36.5], [101.5, 51.5]], color: 'FF0000FF', widthPx: 1, join: 'round', cap: 'round', band: BAND },
    ])
  })

  it('keeps polyline vertices at the same logical positions across zoom and device scales', () => {
    // Task 5 fix round 1: vertices are snapped on the device grid but emitted in logical px, so {zoom:2} and
    // {dpr:2} must reproduce the zoom-1 geometry within the ≤ 0.5 device px snapping slack (vertexSnapRule).
    // [1, 3] in ZCELL keeps every slot centre and mapped y off an exact device-pixel boundary — an
    // integral coordinate sits on the boundary where the parity modes (centre vs corner) legitimately differ.
    // ZCELL: CELL nudged by (0.25, 0.5) so no zoom-1 coordinate sits near a pixel boundary with the Task 7 insets
    // (5.08 pt top from the gridline puts CELL's vertices within 0.1 px of one).
    const ZCELL = { x: 20.25, y: 30.5, w: 100, h: 40 }
    const groups = { markers: true, colors: { markers: 'FFC000FF' } }
    const base = polylines(sparklineGeometry(mk([1, 3], groups), ZCELL, Z1))[0]
    for (const scale of [{ zoom: 2, dpr: 1 }, { zoom: 1, dpr: 2 }]) {
      const devPx = scale.zoom * scale.dpr
      const tol = 0.5 / devPx // 0.5 device px expressed in logical px
      const line = polylines(sparklineGeometry(mk([1, 3], groups), ZCELL, scale))[0]
      base.points.forEach(([x, y], k) => {
        expect(Math.abs(line.points[k][0] - x)).toBeLessThanOrEqual(tol)
        expect(Math.abs(line.points[k][1] - y)).toBeLessThanOrEqual(tol)
      })
      // The whole-device-px stroke keeps its logical width (lineStrokeWidth): 2 device px at 2× → 1 logical px.
      expect(line.widthPx).toBe(1)
      // Vertices line up with the marker centres within the snap slack.
      const ms = markers(sparklineGeometry(mk([1, 3], groups), ZCELL, scale))
      expect(ms).toHaveLength(2)
      ms.forEach((m, k) => {
        expect(Math.abs(m.x - line.points[k][0])).toBeLessThanOrEqual(tol)
        expect(Math.abs(m.y - line.points[k][1])).toBeLessThanOrEqual(tol)
      })
    }
  })

  it('insets the first and last points by P/(2n) instead of the plot edges (plotInsetLeft/Right, lineEndpointRule)', () => {
    // P/2n = 93.3333px/6 = 15.556 px = 11.667 pt from each edge: first x 39.556 px, NOT at the plot edge L = 24 px.
    // The snapped 39.5 and 101.5 bound both edges.
    expect(polylines(sparklineGeometry(mk([1, 3, 2]), CELL, Z1))[0].points[0]).toEqual([39.5, 65.5])
    expect(polylines(sparklineGeometry(mk([1, 3, 2]), CELL, Z1))[0].points[2]).toEqual([101.5, 51.5])
  })

  it('rounds lineWeight to whole device px of the current scale (lineStrokeWidth)', () => {
    // The row pins max(1, round-half-up(lineWeight)) in device px at 1 px/pt: 0.5 → 1, 0.75 → 1, 2.25 → 2, 4.5 → 5.
    // The viewer runs at 4/3 device px per pt at zoom 1: 0.75 pt → 1.0 device px → 1; 2.25 pt → 3.0 → 3;
    // 4.5 pt → 6.0 → 6. At zoom 2: 0.75 → 2 device px (1 logical), 2.25 → 6 (3 logical.
    const width = (scale: { zoom: number; dpr: number }, weight: number): number =>
      polylines(sparklineGeometry(mk([1, 3], { lineWeightPt: weight }), CELL, scale))[0].widthPx
    expect(width(Z1, 0.75)).toBe(1)
    expect(width(Z1, 0.5)).toBe(1)
    expect(width(Z1, 2.25)).toBe(3)
    expect(width(Z1, 4.5)).toBe(6)
    expect(width({ zoom: 2, dpr: 1 }, 0.75)).toBe(1)
    expect(width({ zoom: 2, dpr: 1 }, 2.25)).toBe(3)
    expect(width({ zoom: 1, dpr: 2 }, 2.25)).toBe(3)
  })

  it('marks every point with a round-snapped diamond of half-diagonal (4/3)·lineWeight (markerShape, markerSize, vertexSnapRule)', () => {
    // Ordinary markers snap their centre to round(x), round(y):
    //   round(39.556)=40, round(70.667)=71, round(101.778)=102; round(65.92)=66, round(36.773)=37, round(51.347)=51.
    // Half-diagonal: 4/3·0.75 pt = 1 pt = 4/3 logical px (markerSize).
    expect(markers(sparklineGeometry(mk([1, 3, 2], { markers: true, colors: { markers: 'FFC000FF' } }), CELL, Z1))).toEqual([
      { kind: 'marker', x: 40, y: 66, sizePx: 4 / 3, color: 'FFC000FF', shape: 'diamond' },
      { kind: 'marker', x: 71, y: 37, sizePx: 4 / 3, color: 'FFC000FF', shape: 'diamond' },
      { kind: 'marker', x: 102, y: 51, sizePx: 4 / 3, color: 'FFC000FF', shape: 'diamond' },
    ])
  })

  it('sizes markers from lineWeight: half-diagonal 3 pt at 2.25 (markerSize)', () => {
    // (4/3)·2.25 = 3 pt = 4 logical px.
    expect(markers(sparklineGeometry(mk([1, 2], { markers: true, lineWeightPt: 2.25, colors: { markers: 'FF000000' } }), CELL, Z1))[0].sizePx).toBe(4)
  })

  it('colour at a point that is first and high is the high colour (colourPrecedence)', () => {
    // Last painted of markers → first/last → high → low wins; high beats first (line-precedence row 1: idx 0 is high red).
    const p = sparklineGeometry(mk([3, 2, 1], {
      markers: true, first: true, last: true, high: true, low: true,
      colors: { markers: 'FF111111', first: 'FF222222', last: 'FF333333', high: 'FF444444', low: 'FF555555' },
    }), CELL, Z1)
    expect(markers(p).map(m => m.color)).toEqual(['FF444444', 'FF111111', 'FF555555'])
  })

  it('a flat series is drawn with every point low (colourPrecedence, line-flat row 3, tieRule)', () => {
    // Every point is both high and low; low wins, and every tie is coloured.
    const p = sparklineGeometry(mk([2, 2, 2], {
      markers: true, first: true, last: true, high: true, low: true,
      colors: { markers: 'FF111111', first: 'FF222222', last: 'FF333333', high: 'FF444444', low: 'FF555555' },
    }), CELL, Z1)
    expect(markers(p).map(m => m.color)).toEqual(['FF555555', 'FF555555', 'FF555555'])
  })

  it('colours every tied extreme (tieRule)', () => {
    // [2,9,4,9,1,6,1,9]: high at idx 1, 3, 7; low at idx 4, 6; no markers flag → no marker anywhere else.
    // Slot x (lineEndpointRule, n = 8): 29.833, 41.5, 53.167, 64.833, 76.5, 88.167, 99.833, 111.5.
    // Special points snap to ⌊x⌋, ⌊y⌋ (vertexSnapRule): high value 9 → y(9) = T = 36.773 → 36; low value 1 → y = 65.92 → 65.
    const p = markers(sparklineGeometry(mk([2, 9, 4, 9, 1, 6, 1, 9], {
      high: true, low: true, colors: { high: 'FF990000', low: 'FF009999' },
    }), CELL, Z1))
    expect(p).toEqual([
      { kind: 'marker', x: 41, y: 36, sizePx: 4 / 3, color: 'FF990000', shape: 'diamond' },
      { kind: 'marker', x: 64, y: 36, sizePx: 4 / 3, color: 'FF990000', shape: 'diamond' },
      { kind: 'marker', x: 76, y: 65, sizePx: 4 / 3, color: 'FF009999', shape: 'diamond' },
      { kind: 'marker', x: 99, y: 65, sizePx: 4 / 3, color: 'FF009999', shape: 'diamond' },
      { kind: 'marker', x: 111, y: 36, sizePx: 4 / 3, color: 'FF990000', shape: 'diamond' },
    ])
  })

  it('bridges error slots even under gap: the chord joins the neighbours, no vertex at the slot (cellValueRule, spanRule)', () => {
    // x of the null slot is dropped entirely (no vertex): straight 1 → 3 chord from 39.5,65.5 to 101.5,36.5.
    const values = [1, null, 3]
    expect(polylines(sparklineGeometry(mk(values, { displayEmptyCellsAs: 'gap' }, { bridged: [1] }), CELL, Z1))).toEqual([
      { kind: 'polyline', points: [[39.5, 65.5], [101.5, 36.5]], color: '000000', widthPx: 1, join: 'round', cap: 'round', band: BAND },
    ])
    expect(polylines(sparklineGeometry(mk(values, { displayEmptyCellsAs: 'span' }, { bridged: [1] }), CELL, Z1))).toEqual([
      { kind: 'polyline', points: [[39.5, 65.5], [101.5, 36.5]], color: '000000', widthPx: 1, join: 'round', cap: 'round', band: BAND },
    ])
  })

  it('gap breaks the polyline; span bridges interior gaps and keeps the slots (gapRule, spanRule)', () => {
    // n = 4 slot centres (lineEndpointRule): 35.667, 59, 82.333 (empty), 105.667 → snapped 35.5, 59.5, 82.5, 105.5.
    // [1, 2, _, 3] gap: runs [0,1] and [3]; the lone last point draws no line → one 2-point polyline.
    expect(polylines(sparklineGeometry(mk([1, 2, null, 3], { displayEmptyCellsAs: 'gap' }), CELL, Z1))).toEqual([
      { kind: 'polyline', points: [[35.5, 65.5], [59.5, 51.5]], color: '000000', widthPx: 1, join: 'round', cap: 'round', band: BAND },
    ])
    // span: interior null bridged by a straight segment; slots kept, so x does not close up.
    expect(polylines(sparklineGeometry(mk([1, 2, null, 3], { displayEmptyCellsAs: 'span' }), CELL, Z1))).toEqual([
      { kind: 'polyline', points: [[35.5, 65.5], [59.5, 51.5], [105.5, 36.5]], color: '000000', widthPx: 1, join: 'round', cap: 'round', band: BAND },
    ])
    // span: leading and trailing empties are NOT extended — the line starts/ends at the outermost data points.
    expect(polylines(sparklineGeometry(mk([null, 1, 3, null], { displayEmptyCellsAs: 'span' }), CELL, Z1))).toEqual([
      { kind: 'polyline', points: [[59.5, 65.5], [82.5, 36.5]], color: '000000', widthPx: 1, join: 'round', cap: 'round', band: BAND },
    ])
  })

  it('draws the axis only when displayXAxis and min < 0 < max, at y(0), 1 device px wide across the plot box (axisRule)', () => {
    // [−4, 4]: y(0) = T + 4/8·29.1467 = 51.347 → snapped 51.5 (verticalMapping); x from L pixel centre 24.5 to R pixel
    // centre 117.5 (the row pins x = L+½ … R+½ at the capture scale); 1 device px wide.
    const p = sparklineGeometry(mk([-4, 4], { displayXAxis: true, colors: { axis: 'FF8000FF' } }), CELL, Z1)
    expect(axes(p)).toEqual([{ kind: 'axis', x0: 24.5, x1: 117.5, y: 51.5, color: 'FF8000FF', widthPx: 1 }])
  })

  it('never draws the axis for all-positive or all-negative data or a missing colourAxis (axisRule, colourFallback)', () => {
    expect(axes(sparklineGeometry(mk([1, 2], { displayXAxis: true, colors: { axis: 'FF8000FF' } }), CELL, Z1))).toEqual([])
    expect(axes(sparklineGeometry(mk([-9, -2], { displayXAxis: true, colors: { axis: 'FF8000FF' } }), CELL, Z1))).toEqual([])
    expect(axes(sparklineGeometry(mk([-4, 4], { displayXAxis: true }), CELL, Z1))).toEqual([])
  })

  it('draws a flat positive series at the plot top and a flat zero series at the vertical middle (flatSeriesRule)', () => {
    // min = max = 5 ≠ 0: scale widens to [0, 5], the value is the max → plot top row (T → snapped 36.5).
    expect(polylines(sparklineGeometry(mk([5, 5, 5]), CELL, Z1)).map(l => l.points)).toEqual([
      [[39.5, 36.5], [70.5, 36.5], [101.5, 36.5]],
    ])
    // min = max = 0: vertical middle (T + plotH/2 = 51.347 → 51.5); n = 2 slot centres 47.333/94.667 → 47.5/94.5.
    expect(polylines(sparklineGeometry(mk([0, 0]), CELL, Z1)).map(l => l.points)).toEqual([
      [[47.5, 51.5], [94.5, 51.5]],
    ])
  })

  it('a single point draws nothing without flags and one diamond at the plot centre and top with flags (singlePointRule, specialPointSize, flatSeriesRule)', () => {
    // No point flags → nothing drawn at all.
    expect(sparklineGeometry(mk([5]), CELL, Z1)).toEqual([])
    // All flags on colourwise: one special-point diamond at the slot centre x and the flat-positive top y;
    // low wins the colour (colourPrecedence: single point with every flag is low).
    const p = sparklineGeometry(mk([5], {
      markers: true, first: true, last: true, high: true, low: true,
      colors: { markers: 'FF111111', first: 'FF222222', last: 'FF333333', high: 'FF444444', low: 'FF00A0A0' },
    }), CELL, Z1)
    expect(polylines(p)).toEqual([]) // no line for a single point
    // centre x = L + plotW/2 = 70.667 (slot centre, lineEndpointRule); y = top T = 36.773 → special points snap to
    // ⌊x⌋, ⌊y⌋ = (70, 36) (vertexSnapRule); the measured single-point diamond sits at cols 146–147, rows 4–5.
    expect(markers(p)).toEqual([{ kind: 'marker', x: 70, y: 36, sizePx: 4 / 3, color: 'FF00A0A0', shape: 'diamond' }])
  })

  it('draws nothing in a cell the insets swallow, markers included (tinyCellRule)', () => {
    // tinyCellRule host: inner 13 × 4 pt, smaller than the 7.65 pt of vertical insets. 17.34 × 5.34 logical px = 13 × 4 pt.
    expect(sparklineGeometry(mk([1, 2, 3], { markers: true, colors: { markers: 'FF000000' } }), { x: 0, y: 0, w: 17.34, h: 5.34 }, Z1)).toEqual([])
  })

  it('clips strokes to the plot box rows and cuts segments at the crossings, markers unclipped (lineClipRule, customClipRule)', () => {
    // [0, 6, 0] on custom scale [2, 6] (customScaleRule): y(6) = T → 36.5; y(0) = T + 1.5·plotH = 80.493 → 80.5,
    // below the plot box. The CENTRELINE is cut at the snapped bottom row snap(B) = 65.5 (lineClipRule, customClipRule:
    // the ink of line-scale-custom-clip runs exactly between the vertices of the band values): it re-enters at
    // x = 39.5 + (80.5−65.5)/44·31 = 50.0682 and leaves at 70.5 + (65.5−36.5)/44·31 = 90.9318. The STROKE is clipped
    // to the band rows [⌊T⌋, ⌊B⌋+1) = [36, 66).
    const p = sparklineGeometry(mk([0, 6, 0], { markers: true, colors: { markers: 'FF000000' } }, { min: 2, max: 6 }), CELL, Z1)
    const lines = polylines(p)
    expect(lines).toHaveLength(1)
    expect(lines[0].points[0][0]).toBeCloseTo(50.0682, 4)
    expect(lines[0].points[0][1]).toBeCloseTo(65.5, 7)
    expect(lines[0].points[1]).toEqual([70.5, 36.5])
    expect(lines[0].points[2][0]).toBeCloseTo(90.9318, 4)
    expect(lines[0].band).toEqual(BAND)
    // Markers are not clipped: the out-of-band point still carries its diamond (row 80 below the band).
    expect(markers(p)).toContainEqual({ kind: 'marker', x: 40, y: 80, sizePx: 4 / 3, color: 'FF000000', shape: 'diamond' })
  })

  it('skips the first/last mark when its slot is null (Task 4 note: firstIndex is 0 even with a leading empty)', () => {
    // Leading and trailing null under gap: the first slot has no vertex → no first marker; the last draws at
    // ⌊x⌋, ⌊y⌋ (vertexSnapRule): x = ⌊101.778⌋ = 101, y(3) = 36.773 → 36.
    const p = markers(sparklineGeometry(
      mk([null, 1, 3], { first: true, last: true, colors: { first: 'FF222222', last: 'FF333333' } }),
      CELL, Z1,
    ))
    expect(p).toEqual([{ kind: 'marker', x: 101, y: 36, sizePx: 4 / 3, color: 'FF333333', shape: 'diamond' }])
  })

  it('returns no primitives for column and stacked types in Stage 1 (spec §10: only line is staged)', () => {
    expect(sparklineGeometry(mk([1, 2], { type: 'column' }), CELL, Z1)).toEqual([])
    expect(sparklineGeometry(mk([1, 2], { type: 'stacked' }), CELL, Z1)).toEqual([])
  })

  it('falls back to black when colorSeries is absent, and skips elements with missing colours (colourFallback)', () => {
    expect(polylines(sparklineGeometry(mk([1, 2]), CELL, Z1))[0].color).toBe('000000')
    // colorMarkers absent → the markers flag is effectively off (while series still draws; the row's row-6 case).
    expect(markers(sparklineGeometry(mk([1, 2], { markers: true }), CELL, Z1))).toEqual([])
  })
})
