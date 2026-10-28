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
    highIndices: resolved.highIndices ?? high, lowIndices: resolved.lowIndices ?? low, min, max, axis: resolved.axis,
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

  it('falls back to black when colorSeries is absent, and skips elements with missing colours (colourFallback)', () => {
    expect(polylines(sparklineGeometry(mk([1, 2]), CELL, Z1))[0].color).toBe('000000')
    // colorMarkers absent → the markers flag is effectively off (while series still draws; the row's row-6 case).
    expect(markers(sparklineGeometry(mk([1, 2], { markers: true }), CELL, Z1))).toEqual([])
  })

  it('keeps the line path untouched by Stage 2: same [1, 3, 2] polyline (regression)', () => {
    expect(polylines(sparklineGeometry(mk([1, 3, 2], { colors: { series: 'FF0000FF' } }), CELL, Z1))).toEqual([
      { kind: 'polyline', points: [[39.5, 65.5], [70.5, 36.5], [101.5, 51.5]], color: 'FF0000FF', widthPx: 1, join: 'round', cap: 'round', band: BAND },
    ])
  })
})

describe('column and win/loss sparkline geometry (Stage 2, Task 8)', () => {
  type Bar = Extract<SparkPrimitive, { kind: 'rect' }>
  const bars = (p: SparkPrimitive[]): Bar[] => p.filter((x): x is Bar => x.kind === 'rect')
  // Capture hosts, gridline to gridline (cellBoxRule), in logical px; the capture scale 0.75 (zoom) × 1 (dpr) with
  // output 'device' is 1 device px per pt, so the doc's measured device numbers (L = 48, R = 247, T = 5.08,
  // B = 56.94, H = 51.86, P = 199; merged P = 403, R = 451) come out exactly in the emitted device coordinates.
  const HOST = { x: 45 * 4 / 3, y: 0, w: 204 * 4 / 3, h: 60 * 4 / 3 }
  const MERGED = { ...HOST, w: 408 * 4 / 3 }
  const CAP = { zoom: 0.75, dpr: 1, output: 'device' as const }
  const SER = 'FF4472C4'

  it('reproduces col-basic: n slots, 40% gap, rounded edges, the δ smallest-bar rule and exact rows', () => {
    // D = [4, 7, 3, 9, 2, 8, 5, 6] on the individual scale [2, 9]: scale excludes 0, so columnNoZeroRule (min > 0)
    // puts the baseline at the bottom plot edge (round(B) = 57) and maps δ = 0.63 pt inside it:
    // y(v) = T + (9−v)/7·(H−δ) with T = 5.08, H = 51.86 → round = tops 42/20/49/5/56/12/34/27 (value 2 keeps a
    // 1-px sliver at row 56). barSlotRule/barWidthRule: width w = slot/1.4 = 17.77, edges round(c ± w/2):
    // [52,69) [76,94) [101,119) [126,144) [151,169) [176,194) [201,219) [226,243).
    expect(bars(sparklineGeometry(mk([4, 7, 3, 9, 2, 8, 5, 6], { type: 'column', colors: { series: SER } }), HOST, CAP))).toEqual([
      { kind: 'rect', x: 52, y: 42, w: 17, h: 15, color: SER },
      { kind: 'rect', x: 76, y: 20, w: 18, h: 37, color: SER },
      { kind: 'rect', x: 101, y: 49, w: 18, h: 8, color: SER },
      { kind: 'rect', x: 126, y: 5, w: 18, h: 52, color: SER },
      { kind: 'rect', x: 151, y: 56, w: 18, h: 1, color: SER },
      { kind: 'rect', x: 176, y: 12, w: 18, h: 45, color: SER },
      { kind: 'rect', x: 201, y: 34, w: 18, h: 23, color: SER },
      { kind: 'rect', x: 226, y: 27, w: 17, h: 30, color: SER },
    ])
  })

  it('widens bars with the merged host: P = 403, every bar 36 px wide (col-merged-host, barWidthRule, mergedHostRule)', () => {
    const p = bars(sparklineGeometry(mk([4, 7, 3, 9, 2, 8, 5, 6], { type: 'column', colors: { series: SER } }), MERGED, CAP))
    expect(p[0]).toEqual({ kind: 'rect', x: 55, y: 42, w: 36, h: 15, color: SER })
    expect(p[4]).toEqual({ kind: 'rect', x: 257, y: 56, w: 36, h: 1, color: SER }) // measured columns 257–292
    expect(p.map(b => b.w)).toEqual([36, 36, 36, 36, 36, 36, 36, 36])
  })

  it('runs columns from the baseline row round(y(0)) = 31 both ways, negatives only orange under negative="1" (columnBaselineRule, barNegativeRule)', () => {
    // [−3, 2, −1, 4, −4, 3, 1, −2] on [−4, 4]: scale includes 0, plain verticalMapping; y(0) = 31.01 → baseline
    // row 31: positive bars end at row 30 (tops 5/12/18/25 for 4/3/2/1), negative bars start at row 31
    // (bottoms 37/44/50/57 for −1/−2/−3/−4). Same slots/widths as col-basic.
    const values = [-3, 2, -1, 4, -4, 3, 1, -2]
    const geom = (negative: boolean) =>
      bars(sparklineGeometry(mk(values, { type: 'column', negative, colors: { series: SER, negative: 'FF8000FF' } }), HOST, CAP))
    expect(geom(false).map(b => [b.x, b.y, b.w, b.h])).toEqual([
      [52, 31, 17, 19], [76, 18, 18, 13], [101, 31, 18, 6], [126, 5, 18, 26],
      [151, 31, 18, 26], [176, 12, 18, 19], [201, 25, 18, 6], [226, 31, 17, 13],
    ])
    expect(geom(false).map(b => b.color)).toEqual(Array(8).fill(SER))
    expect(geom(true).map(b => b.color)).toEqual(
      ['FF8000FF', SER, 'FF8000FF', SER, 'FF8000FF', SER, SER, 'FF8000FF'],
    )
    // colourFallback: a missing colorNegative is not applied — negative bars keep the series colour.
    expect(bars(sparklineGeometry(mk(values, { type: 'column', negative: true, colors: { series: SER } }), HOST, CAP)).filter(b => b.color === 'FF8000FF')).toEqual([])
  })

  it('hangs all-negative bars from the top edge, mirrored δ rule, heights 8,37,15,52,1,45,23,30 (columnNoZeroRule, max < 0)', () => {
    // Scale [−9, −2] excludes 0: baseline is the TOP plot edge (round(T) = 5) and the end nearest 0 (max −2) maps
    // δ = 0.63 pt inside: y(v) = T + δ + (−2−v)/7·(H−δ). min −9 → y 56.94 → row 57; max −2 → row 6, a 1-px bar.
    expect(bars(sparklineGeometry(mk([-3, -7, -4, -9, -2, -8, -5, -6], { type: 'column', negative: true, colors: { series: SER, negative: 'FF8000FF' } }), HOST, CAP)).map(b => [b.y, b.h, b.color])).toEqual([
      [5, 8, 'FF8000FF'], [5, 37, 'FF8000FF'], [5, 15, 'FF8000FF'], [5, 52, 'FF8000FF'],
      [5, 1, 'FF8000FF'], [5, 45, 'FF8000FF'], [5, 23, 'FF8000FF'], [5, 30, 'FF8000FF'],
    ])
  })

  it('draws win/loss bars full-height by exact sign, ±1e-16 included, zeros and empties excluded (winLossRule)', () => {
    // Wins [round(T), round(T + H/2)) = rows [5, 31), losses [31, round(B)) = [31, 57) — whatever the magnitudes;
    // v = 0 draws nothing. n = 5: slot 39.8, bar edges round(c ± 14.214): [54,82) [93,122) [133,162) [173,202) [213,241).
    expect(bars(sparklineGeometry(mk([1e-16, -1e-16, 0, 2, -3], { type: 'stacked', negative: true, colors: { series: SER, negative: 'FF8000FF' } }), HOST, CAP)).map(b => [b.x, b.y, b.w, b.h, b.color])).toEqual([
      [54, 5, 28, 26, SER], [93, 31, 29, 26, 'FF8000FF'], [173, 5, 29, 26, SER], [213, 31, 28, 26, 'FF8000FF'],
    ])
  })

  it('draws the win/loss axis at the middle row whenever displayXAxis, even all-zero, painted after the bars; the column axis keeps min < 0 < max (barAxisRule)', () => {
    // All-zero stacked: no bars; the axis alone, at round(T + H/2) row 31 (snap 31.5, 1 device px, floor(L)+½ …
    // floor(R)+½ = 48.5 … 247.5, colorAxis present).
    const p = sparklineGeometry(mk([0, 0, 0], { type: 'stacked', displayXAxis: true, colors: { series: SER, axis: 'FF8000FF' } }), HOST, CAP)
    expect(p).toEqual([{ kind: 'axis', x0: 48.5, x1: 247.5, y: 31.5, color: 'FF8000FF', widthPx: 1 }])
    // Win/loss axis colour missing → not drawn (colourFallback).
    expect(axes(sparklineGeometry(mk([0, 0], { type: 'stacked', displayXAxis: true, colors: { series: SER } }), HOST, CAP))).toEqual([])
    // Column axis: same y(0) as the line's (axisRule): [−2, 2] → y(0) = 31.01 → 31.5.
    expect(axes(sparklineGeometry(mk([-2, 2], { type: 'column', displayXAxis: true, colors: { axis: 'FF8000FF' } }), HOST, CAP))[0].y).toBe(31.5)
    // The all-positive column axis is not drawn (ruling; axisRule's min < 0 < max is unmeasured for columns).
    expect(axes(sparklineGeometry(mk([2, 9], { type: 'column', displayXAxis: true, colors: { axis: 'FF8000FF' } }), HOST, CAP))).toEqual([])
  })

  it('shows gap/zero/span empties and a literal 0 as an empty slot, slots keeping their positions (columnEmptyRule)', () => {
    // [2, _, 4] on [2, 4]: middle slot empty; the drawn bars keep slots 0 and 2 of n = 3 (slot 66.333, w 47.38:
    // c 81.17 → [57,105); c 213.83 → [190,238)): value 2 (the smallest) keeps a δ sliver at rows 56–57, value 4
    // tops at row 5.
    for (const mode of ['gap', 'zero', 'span'] as const) {
      expect(bars(sparklineGeometry(mk([2, null, 4], { type: 'column', displayEmptyCellsAs: mode, colors: { series: SER } }), HOST, CAP)).map(b => [b.x, b.x + b.w, b.y, b.h])).toEqual([
        [57, 105, 56, 1], [190, 238, 5, 52],
      ])
    }
    // A literal 0 under a zero-including scale sits at the baseline → zero-height bar → nothing.
    expect(bars(sparklineGeometry(mk([3, 0, 5], { type: 'column', colors: { series: SER } }), HOST, CAP))).toHaveLength(2)
  })

  it('draws a flat column full-height from the bottom (positive), from the top (negative), nothing (zero) — unmeasured ruling', () => {
    // flatSeriesRule widening to [min(0, v), max(0, v)] + columnBaselineRule: [0, 5] → baseline row 57, value at row 5.
    expect(bars(sparklineGeometry(mk([5, 5, 5], { type: 'column', colors: { series: SER } }), HOST, CAP)).map(b => [b.y, b.h])).toEqual([[5, 52], [5, 52], [5, 52]])
    // [−5, 0]: baseline row 5, value at row 57.
    expect(bars(sparklineGeometry(mk([-5, -5, -5], { type: 'column', colors: { series: SER } }), HOST, CAP)).map(b => [b.y, b.h])).toEqual([[5, 52], [5, 52], [5, 52]])
    expect(bars(sparklineGeometry(mk([0, 0, 0], { type: 'column', colors: { series: SER } }), HOST, CAP))).toEqual([])
  })

  it('draws a single value as one full-width slot bar (ruling: singlePointRule does not apply to bars)', () => {
    // slot = P = 199, w = 142.14, c = 147.5: round edges [76, 219); flat widened [0, 5] → full height rows [5, 57).
    expect(bars(sparklineGeometry(mk([5], { type: 'column', colors: { series: SER } }), HOST, CAP))).toEqual([{ kind: 'rect', x: 76, y: 5, w: 143, h: 52, color: SER }])
    expect(bars(sparklineGeometry(mk([-5], { type: 'column', colors: { series: SER } }), HOST, CAP))).toEqual([{ kind: 'rect', x: 76, y: 5, w: 143, h: 52, color: SER }])
  })

  it('clips bars cut by a custom scale to the plot box rows [round(T), round(B)) (ruling 3; customScaleRule)', () => {
    // [150, 50] on custom [0, 100]: y(150) = 5.08 − 0.5·51.86 → round −21 → clamped to row 5: full-height bar;
    // y(50) = 31.01 → rows [31, 57). Baseline y(0) = round(56.94) = 57. n = 2: slot 99.5, w 71: [62,133), [162,233).
    expect(bars(sparklineGeometry(mk([150, 50], { type: 'column' }, { min: 0, max: 100 }), HOST, CAP)).map(b => [b.x, b.y, b.w, b.h])).toEqual([
      [62, 5, 71, 52], [162, 31, 71, 26],
    ])
  })

  it('paints bar colours last of series → negative → first/last → high → low, every tie coloured (barColourPrecedence)', () => {
    // [−3, 1, −4, 2, 2] with everything on: negative orange on slots 0/2; first beats it on 0; low −4 (slot 2)
    // beats negative and most everything; high 2 at slots 3 AND 4 (tie) beats series and last; slot 1 stays series.
    const p = bars(sparklineGeometry(mk([-3, 1, -4, 2, 2], {
      type: 'column', negative: true, first: true, last: true, high: true, low: true,
      colors: { series: SER, negative: 'FF8000FF', first: 'FF22AA22', last: 'FF8000AA', high: 'FFCC0000', low: 'FF00AAAA' },
    }), HOST, CAP))
    expect(p.map(b => b.color)).toEqual(['FF22AA22', SER, 'FF00AAAA', 'FFCC0000', 'FFCC0000'])
    // Win/loss: same precedence over full-height sign bars (low −4 is the only low; high beats last on slot 4).
    const w = bars(sparklineGeometry(mk([-3, 1, -4, 2, 2], {
      type: 'stacked', negative: true, first: true, last: true, high: true, low: true,
      colors: { series: SER, negative: 'FF8000FF', first: 'FF22AA22', last: 'FF8000AA', high: 'FFCC0000', low: 'FF00AAAA' },
    }), HOST, CAP))
    expect(w.map(b => b.color)).toEqual(['FF22AA22', SER, 'FF00AAAA', 'FFCC0000', 'FFCC0000'])
  })

  it('ignores first/last/high/low colours whose flag is off, and colours only the flagged elements (Task 9 comparison)', () => {
    // The parser keeps every element colour whether or not the flag is on (the line path gates on the flag, so must bars).
    const colors = { series: SER, negative: 'FF8000FF', first: 'FF22AA22', last: 'FF8000AA', high: 'FFCC0000', low: 'FF00AAAA' }
    for (const type of ['column', 'stacked'] as const) {
      expect(bars(sparklineGeometry(mk([-3, 1, -4, 2, 2], { type, colors }), HOST, CAP)).map(b => b.color), `${type} flags off`).toEqual(Array(5).fill(SER))
      expect(bars(sparklineGeometry(mk([-3, 1, -4, 2, 2], { type, high: true, colors }), HOST, CAP)).map(b => b.color), `${type} high only`)
        .toEqual([SER, SER, SER, 'FFCC0000', 'FFCC0000'])
    }
  })

  it('snaps both bar axes to the odd 1-px centre row, independent of lineWeightPt (barAxisRule)', () => {
    // The bar axis is always 1 device px wide, so its y snaps to the pixel centre ⌊y⌋ + ½ like an odd stroke
    // whatever the line weight: the capture mid row T + H/2 = 31.01 → y 31.5 and the column's y(0) on [−2, 2] =
    // 31.01 → y 31.5, for the default 0.75 pt (1 device px) AND 1.5 pt (2 device px, which the line path's parity
    // snap would put on a corner).
    for (const weight of [0.75, 1.5] as const) {
      expect(axes(sparklineGeometry(mk([-2, 2], { type: 'column', displayXAxis: true, lineWeightPt: weight, colors: { axis: 'FF8000FF' } }), HOST, CAP)).map(a => [a.y, a.widthPx])).toEqual([[31.5, 1]])
      expect(axes(sparklineGeometry(mk([0, 0, -1], { type: 'stacked', displayXAxis: true, lineWeightPt: weight, colors: { axis: 'FF8000FF' } }), HOST, CAP)).map(a => [a.y, a.widthPx])).toEqual([[31.5, 1]])
    }
  })

  it('draws nothing in a cell too small for the plot box, both types (barTinyCellRule)', () => {
    expect(sparklineGeometry(mk([1, 2, 3], { type: 'column', colors: { series: SER } }), { x: 0, y: 0, w: 17.34, h: 5.34 }, Z1)).toEqual([])
    expect(sparklineGeometry(mk([1, -2], { type: 'stacked', displayXAxis: true, colors: { series: SER, axis: 'FF8000FF' } }), { x: 0, y: 0, w: 17.34, h: 5.34 }, Z1)).toEqual([])
  })

  it('keeps bar edges on whole device px and the 40% gap at 3.5 device px per logical px (zoom 1.75 dpr 2, barWidthRule)', () => {
    // Device edges stay integral after rounding; the rounded bar width stays within 1 device px of slot/1.4
    // (the plot width P = 265.333 logical px → 928.667 device px at 3.5×: slot 116.08, bar 82.92).
    const p = bars(sparklineGeometry(mk([4, 7, 3, 9, 2, 8, 5, 6], { type: 'column', colors: { series: SER } }), HOST, { zoom: 1.75, dpr: 2, output: 'device' }))
    const slot = (HOST.w - 3 * 4 / 3 - 2 * 4 / 3) * 1.75 * 2 / 8
    let i = 0
    for (const b of p) {
      for (const e of [b.x, b.y, b.w, b.h]) expect(Number.isInteger(e), `${i}:${e}`).toBe(true)
      expect(Math.abs(b.w - slot / 1.4), `width ${i}`).toBeLessThanOrEqual(1)
      i++
    }
    expect(p).toHaveLength(8)
  })

  it('passes colourFallback and the black series fallback through to bars (colourFallback)', () => {
    // Missing colorSeries → black; a flag without its colour (high) does not decorate anything.
    // n = 2: slot 99.5, w 71: edges [62,133), [162,233).
    const p = bars(sparklineGeometry(mk([2, 4], { type: 'column', high: true, low: true }), HOST, CAP))
    expect(p.map(b => [b.color, b.x])).toEqual([['000000', 62], ['000000', 162]])
  })
})

/**
 * Stage 3 (plan Task 10): the date axis and the right-to-left mirror. Resolve hands geometry the axis (`positions`
 * = each point's fraction of the plot width — dateSlotRule puts the point dated d at its day-slot centre — and
 * `slot` = the day-slot width as the same fraction); geometry multiplies by P and mirrors x about the plot box for
 * rightToLeft (rtlRule: x' = L + R − x BEFORE snapping; indices and colours stay in data order, so first draws at
 * the RIGHT). HOST/CAP like the Stage 2 describe: 1 device px per pt, so the doc's numbers (L = 48, R = 247,
 * P = 199 device px) come out exactly.
 */
describe('date axis and right-to-left geometry (Stage 3, Task 10)', () => {
  type Bar = Extract<SparkPrimitive, { kind: 'rect' }>
  const bars = (p: SparkPrimitive[]): Bar[] => p.filter((x): x is Bar => x.kind === 'rect')
  const HOST = { x: 45 * 4 / 3, y: 0, w: 204 * 4 / 3, h: 60 * 4 / 3 }
  const CAP = { zoom: 0.75, dpr: 1, output: 'device' as const }
  const SER = 'FFC00000', GREEN = 'FF22AA22', PURPLE = 'FF8000AA', HIGH = 'FFCC0000', LOW = 'FF00AAAA', MARK = 'FFC000FF'
  // The date-uneven dates: serials 45000 + [0,1,3,10,11,20,21,30] under [4,7,3,9,2,8,5,6]; span 45000–45030 →
  // 31 equal day slots of P/31 = 6.419 device px. The axis fractions dateSlotRule: ((d − 45000) + ½)/31.
  const OFF = [0, 1, 3, 10, 11, 20, 21, 30]
  const D = [4, 7, 3, 9, 2, 8, 5, 6]
  const axis = { positions: OFF.map(o => (o + 0.5) * (1 / 31)), slot: 1 / 31 }

  it('centres bars on their day slots and sizes them by day width (dateSlotRule, barWidthRule)', () => {
    // date-uneven column (individual scale [2, 9], columnNoZeroRule): slot 6.419, w = slot/1.4 = 4.585; bar edges
    // round(c ± w/2) of the day-slot centres: [49,54) [55,60) [68,73) [113,118) [120,124) [177,182) [184,188)
    // [241,246); tops/heights map exactly like col-basic's (42/20/49/5/56/12/34/27, baseline row 57).
    expect(bars(sparklineGeometry(mk(D, { type: 'column', colors: { series: SER } }, { axis }), HOST, CAP))).toEqual([
      { kind: 'rect', x: 49, y: 42, w: 5, h: 15, color: SER },
      { kind: 'rect', x: 55, y: 20, w: 5, h: 37, color: SER },
      { kind: 'rect', x: 68, y: 49, w: 5, h: 8, color: SER },
      { kind: 'rect', x: 113, y: 5, w: 5, h: 52, color: SER },
      { kind: 'rect', x: 120, y: 56, w: 4, h: 1, color: SER },
      { kind: 'rect', x: 177, y: 12, w: 5, h: 45, color: SER },
      { kind: 'rect', x: 184, y: 34, w: 4, h: 23, color: SER },
      { kind: 'rect', x: 241, y: 27, w: 5, h: 30, color: SER },
    ])
  })

  it('floors fractional serials to day slots while retaining same-day point order (dateFractionalRule)', () => {
    const fractionalAxis = { positions: [1 / 12, 1 / 12, 1 / 12, 1 / 4, 5 / 12, 7 / 12, 3 / 4, 11 / 12], slot: 1 / 6 }
    const barEdges = bars(sparklineGeometry(mk(D, { type: 'column', colors: { series: SER } }, { axis: fractionalAxis }), HOST, CAP))
      .map(bar => [bar.x, bar.x + bar.w])
    expect(barEdges).toEqual([
      [53, 76], [53, 76], [53, 76], [86, 110], [119, 143], [152, 176], [185, 209], [219, 242],
    ])

    const line = polylines(sparklineGeometry(mk(D, { markers: true }, { axis: fractionalAxis }), HOST, CAP))[0]
    expect(line.points).toEqual([
      [64.5, 42.5], [64.5, 19.5], [64.5, 49.5], [97.5, 5.5], [130.5, 56.5], [164.5, 12.5], [197.5, 34.5], [230.5, 27.5],
    ])
    const intradayAxis = { positions: Array(8).fill(0.5), slot: 1 }
    expect(polylines(sparklineGeometry(mk(D, {}, { axis: intradayAxis }), HOST, CAP))[0].points.map(point => point[0]))
      .toEqual(Array(8).fill(147.5))
  })

  it('moves line vertices off the even slots onto the dated day centres (dateSlotRule)', () => {
    // date-uneven line (individual scale [2, 9]): even-slot spacing would put vertices at 60.5/85.5/110.5…;
    // the dated slots put them at 51.5/57.5/70.5/115.5… (x = (o+½)·P/31 snapped ⌊·⌋+½ for the 1-px stroke).
    expect(polylines(sparklineGeometry(mk(D, {}, { axis }), HOST, CAP))[0].points).toEqual([
      [51.5, 42.5], [57.5, 19.5], [70.5, 49.5], [115.5, 5.5], [121.5, 56.5], [179.5, 12.5], [186.5, 34.5], [243.5, 27.5],
    ])
  })

  it('mirrors line x about the plot box, keeping indices and colours in data order (rtlRule)', () => {
    // rtl-line row 1 (even slots, no date axis): x'_i = R − x_i before snapping: vertex centres 59.5…234.5 (a
    // mirror of the even-slot centres; the low vertex at 135.5, not the dated 173.5); markers keep their per-index
    // snapshots (ordinary round, specials ⌊·⌋), so first keeps its colour at the RIGHTMOST slot (data order).
    const p = sparklineGeometry(mk(D, {
      rightToLeft: true, markers: true, first: true, last: true, high: true, low: true,
      colors: { series: SER, markers: MARK, first: GREEN, last: PURPLE, high: HIGH, low: LOW },
    }), HOST, CAP)
    expect(polylines(p)[0].points).toEqual([
      [234.5, 42.5], [209.5, 19.5], [184.5, 49.5], [159.5, 5.5], [135.5, 56.5], [110.5, 12.5], [85.5, 34.5], [60.5, 27.5],
    ])
    expect(markers(p).map(m => [m.x, m.y, m.color])).toEqual([
      [234, 42, GREEN], [210, 20, MARK], [185, 50, MARK], [159, 5, HIGH], [135, 56, LOW], [110, 12, MARK], [85, 35, MARK], [60, 27, PURPLE],
    ])
  })

  it('mirrors a date axis under rightToLeft: the exact tr mirror of the date positions (rtlRule, dateSlotRule)', () => {
    // rtl-line row 2: the dated vertices mirror about the plot box (51.5 ↔ 243.5, 57.5 ↔ 237.5 …) and the line
    // connects them in date order; data-order colours hold (first at the earliest date, now on the right).
    const p = sparklineGeometry(mk(D, { rightToLeft: true }, { axis }), HOST, CAP)
    expect(polylines(p)[0].points).toEqual([
      [243.5, 42.5], [237.5, 19.5], [224.5, 49.5], [179.5, 5.5], [173.5, 56.5], [115.5, 12.5], [108.5, 34.5], [51.5, 27.5],
    ])
  })

  it('mirrors column and win/loss bars with the same rounded edges, colours staying in data order (rtlRule)', () => {
    // rtl-column: rightToLeft mirrors each bar's edges about the plot box, reproducing col-basic's edges exactly
    // in reverse order; first/last colours stay on their data-order bars (first on the RIGHT, last on the left).
    const base = bars(sparklineGeometry(mk(D, { type: 'column', first: true, last: true, colors: { series: SER, first: GREEN, last: PURPLE } }), HOST, CAP))
    expect(base.map(b => b.color)).toEqual([GREEN, SER, SER, SER, SER, SER, SER, PURPLE])
    const rtl = bars(sparklineGeometry(mk(D, { type: 'column', first: true, last: true, rightToLeft: true, colors: { series: SER, first: GREEN, last: PURPLE } }), HOST, CAP))
    expect(rtl.map(b => [b.x, b.y, b.w, b.h])).toEqual([
      [226, 42, 17, 15], [201, 20, 18, 37], [176, 49, 18, 8], [151, 5, 18, 52],
      [126, 56, 18, 1], [101, 12, 18, 45], [76, 34, 18, 23], [52, 27, 17, 30],
    ])
    // The mirrored edges are col-basic's exactly reversed (y/heights keep each bar's own mapping).
    expect(rtl.map(b => [b.x, b.w])).toEqual([...base.map(b => [b.x, b.w])].reverse())
    expect(rtl.map(b => b.color)).toEqual([GREEN, SER, SER, SER, SER, SER, SER, PURPLE])
    // Win/loss over zero: [−3, 1, −4, 2, 2] with first/last — the mirrored edges reproduce the Stage 2 rows
    // ([54,82) [93,122) [133,162) [173,202) [213,241)) reversed, wins above / losses below row 31 as before.
    const wl = bars(sparklineGeometry(mk([-3, 1, -4, 2, 2], { type: 'stacked', negative: true, first: true, last: true, rightToLeft: true, colors: { series: SER, negative: 'FF8000FF', first: GREEN, last: PURPLE } }), HOST, CAP))
    expect(wl.map(b => [b.y, b.h, b.color])).toEqual([
      [31, 26, GREEN], [5, 26, SER], [31, 26, 'FF8000FF'], [5, 26, SER], [5, 26, PURPLE],
    ])
    expect(wl.map(b => b.x)).toEqual([213, 173, 133, 93, 54])
  })
})
