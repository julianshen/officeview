// @vitest-environment node
/**
 * Task A6a (spec 5.6, 7): trendlines. Numeric oracles are hand-computed (closed forms noted beside each case);
 * fixtures are Office-shaped synthetic chart parts, never Office evidence.
 */
import { describe, expect, it } from 'vitest'
import { createCanvas } from 'canvas'
import { parseXmlOrdered } from '../src/core/xml'
import { parseXlsxChart } from '../src/charts/parse'
import { paintXlsxChart } from '../src/charts/paint'
import { layoutChart } from '../src/charts/layout'
import { fitTrendline, TRENDLINE_SAMPLE_LIMIT, type ChartTrendline, type TrendlineFit } from '../src/charts/trendlines'
import type { ChartModel } from '../src/charts/model'
import { parseXlsx } from '../src/xlsx/parse'
import { chartWorkbook } from './helpers/xlsx-chart-fixture'

const parse = (xml: string) => parseXlsxChart(parseXmlOrdered(xml), undefined)!
const lit = (values: number[]) => `<numLit><ptCount val="${values.length}"/>${values.map((v, i) => `<pt idx="${i}"><v>${v}</v></pt>`).join('')}</numLit>`
const trend = (body: string) => `<trendline>${body}</trendline>`
const type = (t: string) => `<trendlineType val="${t}"/>`
const AX = '<axId val="1"/><axId val="2"/>'
const space = (plot: string) => `<chartSpace><chart><autoTitleDeleted val="1"/><plotArea>${plot}</plotArea></chart></chartSpace>`
const scatter = (xs: number[], ys: number[], trendlines: string, extra = '') => parse(space(
  `<scatterChart><scatterStyle val="lineMarker"/><ser><idx val="0"/><order val="0"/>${extra}${trendlines}<xVal>${lit(xs)}</xVal><yVal>${lit(ys)}</yVal></ser>${AX}</scatterChart>`
  + '<valAx><axId val="1"/><axPos val="b"/><crossAx val="2"/></valAx><valAx><axId val="2"/><axPos val="l"/><crossAx val="1"/></valAx>'))
const line = (ys: number[], trendlines: string, grouping = 'standard') => parse(space(
  `<lineChart><grouping val="${grouping}"/><ser><idx val="0"/><order val="0"/>${trendlines}<val>${lit(ys)}</val></ser>${AX}</lineChart>`
  + '<catAx><axId val="1"/><crossAx val="2"/></catAx><valAx><axId val="2"/><crossAx val="1"/></valAx>'))
const trendsOf = (model: ChartModel) => model.preparedAxes?.trendlines ?? []
const codes = (model: ChartModel) => (model.preparedAxes?.diagnostics ?? []).map(d => d.code)
const spec = (s: Partial<ChartTrendline> & Pick<ChartTrendline, 'type'>): ChartTrendline => s
const fit = (s: Partial<ChartTrendline> & Pick<ChartTrendline, 'type'>, xs: number[], ys: number[]): TrendlineFit => {
  const result = fitTrendline(spec(s), xs, ys)
  if (result.kind !== 'fit') throw new Error(`no fit: ${result.code} ${result.reason}`)
  return result
}
const close = (actual: readonly number[], expected: readonly number[], digits = 9) => {
  expect(actual.length).toBe(expected.length)
  actual.forEach((v, i) => expect(v).toBeCloseTo(expected[i], digits))
}
const range = (n: number, from = 1) => Array.from({ length: n }, (_, i) => i + from)

describe('charts-trendlines (Task A6a)', () => {
  it('linear, exponential, logarithmic, polynomial 2-6, power and moving average on known datasets', () => {
    // Linear, inexact: x 1..5, y 2,4,5,4,5. Sxy = 6, Sxx = 10: b = 0.6, a = 4 - 0.6*3 = 2.2; SSres 2.4, SStot 6: R² 0.6.
    const linear = fit({ type: 'linear' }, range(5), [2, 4, 5, 4, 5])
    close(linear.coefficients, [2.2, 0.6]); expect(linear.r2).toBeCloseTo(0.6, 12)
    // Linear, inexact: (1,1),(2,2),(3,2): Sxy = 1, Sxx = 2: b 0.5, a = 5/3 - 0.5*2 = 2/3; SSres 1/6, SStot 2/3: R² 0.75.
    const small = fit({ type: 'linear' }, [1, 2, 3], [1, 2, 2])
    close(small.coefficients, [2 / 3, 0.5]); expect(small.r2).toBeCloseTo(0.75, 12)
    // Exponential y = 3e^(0.5x) at x 0..4 (exact): c 3, b 0.5, R² 1 (of ln y).
    const exp = fit({ type: 'exp' }, range(5, 0), range(5, 0).map(x => 3 * Math.exp(0.5 * x)))
    close(exp.coefficients, [3, 0.5]); expect(exp.r2).toBeCloseTo(1, 12)
    // Exponential, inexact: x 0,1,2; ln y = 0,1,3 -> slope 1.5, ln c = 4/3 - 1.5 = -1/6; R² = 1 - (1/6)/(14/3) = 27/28.
    const expInexact = fit({ type: 'exp' }, [0, 1, 2], [1, Math.E, Math.exp(3)])
    close(expInexact.coefficients, [Math.exp(-1 / 6), 1.5]); expect(expInexact.r2).toBeCloseTo(27 / 28, 12)
    // Logarithmic y = 2 ln x + 1 at x 1,2,4,8 (exact): [c, b] = [2, 1].
    const log = fit({ type: 'log' }, [1, 2, 4, 8], [1, 2, 4, 8].map(x => 2 * Math.log(x) + 1))
    close(log.coefficients, [2, 1]); expect(log.r2).toBeCloseTo(1, 12)
    // Power y = 5x^1.5 at x 1,4,9 (exact): [5, 1.5].
    const power = fit({ type: 'power' }, [1, 4, 9], [5, 40, 135])
    close(power.coefficients, [5, 1.5]); expect(power.r2).toBeCloseTo(1, 12)
    // Polynomial order 2, inexact: x 0..3, y 0,1,1,0 -> by symmetry about 1.5, y = -0.5x² + 1.5x (+0): exact on these points.
    close(fit({ type: 'poly', order: 2 }, [0, 1, 2, 3], [0, 1, 1, 0]).coefficients, [0, 1.5, -0.5])
    // Polynomial order 2, inexact: x -1,0,1,2; y 1,0,0,3. Normal equations (Σ1=4, Σx=2, Σx²=6, Σx³=8, Σx⁴=18;
    // Σy=4, Σxy=5, Σx²y=13) eliminate to 5b1 + 5b2 = 3, 5b1 + 9b2 = 7: b2 = 1, b1 = -0.4, b0 = -0.3.
    close(fit({ type: 'poly', order: 2 }, [-1, 0, 1, 2], [1, 0, 0, 3]).coefficients, [-0.3, -0.4, 1])
    // Orders 3-6: exact polynomials with distinct coefficients, recovered exactly from 10 points.
    for (const order of [3, 4, 5, 6]) {
      const coefficients = Array.from({ length: order + 1 }, (_, j) => (j % 2 ? -1 : 1) * (j + 1) / 2)
      const xs = range(10, -4), ys = xs.map(x => coefficients.reduce((sum, c, j) => sum + c * x ** j, 0))
      const result = fit({ type: 'poly', order }, xs, ys)
      close(result.coefficients, coefficients, 6)
      expect(result.r2).toBeCloseTo(1, 9)
    }
    // Moving average, period 3 on 1..6: averages 2,3,4,5 at x 3..6 (each window's last point).
    const average = fitTrendline(spec({ type: 'movingAvg', period: 3 }), range(6), range(6))
    expect(average.kind).toBe('fit')
    if (average.kind === 'fit') expect(Array.from(average.averages!)).toEqual([3, 2, 4, 3, 5, 4, 6, 5])
    const pairs = fitTrendline(spec({ type: 'movingAvg', period: 2 }), range(4), [1, 2, 3, 4])
    if (pairs.kind === 'fit') expect(Array.from(pairs.averages!)).toEqual([2, 1.5, 3, 2.5, 4, 3.5])
  })

  it('XY uses x values; category charts use 1..n', () => {
    // Scatter: x 2,4,6 against y 1,2,3 -> slope 0.5, intercept 0. The same values on a line chart (x = 1..3) -> slope 1, intercept 0.
    const xy = scatter([2, 4, 6], [1, 2, 3], trend(type('linear')))
    expect(xy.groups[0].series[0].trendlines).toHaveLength(1)
    close(trendsOf(xy)[0].coefficients, [0, 0.5])
    const category = line([1, 2, 3], trend(type('linear')))
    close(trendsOf(category)[0].coefficients, [0, 1])
    // A blank keeps its category position: y 1,_,3 sits at x 1 and 3.
    const gap = parse(space(`<lineChart><grouping val="standard"/><ser><idx val="0"/><order val="0"/>${trend(type('linear'))}<val><numLit><ptCount val="3"/><pt idx="0"><v>1</v></pt><pt idx="2"><v>5</v></pt></numLit></val></ser>${AX}</lineChart><catAx><axId val="1"/><crossAx val="2"/></catAx><valAx><axId val="2"/><crossAx val="1"/></valAx>`))
    close(trendsOf(gap)[0].coefficients, [-1, 2])
  })

  it('non-positive exclusions per type with diagnostics', () => {
    // Exponential drops y <= 0: the remaining three points are exactly 3e^(0.5x).
    const exp = scatter([0, 1, 2, 3], [-1, 3 * Math.exp(0.5), 3 * Math.E, 3 * Math.exp(1.5)], trend(type('exp')))
    close(trendsOf(exp)[0].coefficients, [3, 0.5])
    expect(trendsOf(exp)[0].excluded).toBe(1)
    expect(exp.preparedAxes!.diagnostics.find(d => d.code === 'chart-trendline-excluded')?.reason).toMatch(/1 point/)
    // Logarithmic drops x <= 0 (two points), whatever their y.
    const log = scatter([-1, 0, 1, 2, 4], [9, 9, 1, 1 + 2 * Math.log(2), 1 + 2 * Math.log(4)], trend(type('log')))
    close(trendsOf(log)[0].coefficients, [2, 1])
    expect(trendsOf(log)[0].excluded).toBe(2)
    expect(codes(log)).toContain('chart-trendline-excluded')
    // Power drops both x <= 0 and y <= 0.
    const power = scatter([0, 1, 4, 9], [5, 5, 40, -3], trend(type('power')))
    close(trendsOf(power)[0].coefficients, [5, 1.5])
    expect(trendsOf(power)[0].excluded).toBe(2)
    // Linear and polynomial keep non-positive points: no exclusion diagnostic.
    const linear = scatter([-1, 0, 1], [-2, 0, 2], trend(type('linear')))
    close(trendsOf(linear)[0].coefficients, [0, 2])
    expect(codes(linear)).not.toContain('chart-trendline-excluded')
    // Fewer points left than unknowns: nothing drawn (singular).
    const none = scatter([1, 2, 3], [-1, -2, 4], trend(type('exp')))
    expect(trendsOf(none)).toHaveLength(0)
    expect(codes(none)).toEqual(expect.arrayContaining(['chart-trendline-excluded', 'chart-singular-trendline']))
  })

  it('a singular polynomial fit draws nothing with chart-singular-trendline', () => {
    // Three category points cannot determine an order-5 polynomial.
    const model = line([1, 4, 2], trend(type('poly') + '<order val="5"/>' + '<spPr><ln w="19050"><solidFill><srgbClr val="FF0000"/></solidFill></ln></spPr>'))
    expect(trendsOf(model)).toHaveLength(0)
    expect(codes(model)).toContain('chart-singular-trendline')
    expect(paint(model).strokes.filter(s => s.style === '#FF0000')).toHaveLength(0)
    // Every x equal: a linear fit is singular too.
    expect(fitTrendline(spec({ type: 'linear' }), [2, 2, 2], [1, 2, 3])).toMatchObject({ kind: 'none', code: 'chart-singular-trendline' })
    // Order 3 on four distinct x of which two repeat: three distinct x for four unknowns.
    expect(fitTrendline(spec({ type: 'poly', order: 3 }), [1, 1, 2, 3], [1, 2, 3, 4])).toMatchObject({ kind: 'none', code: 'chart-singular-trendline' })
    // Two distinct x (0.1, 0.3) for three unknowns: rounding leaves a tiny non-zero pivot (coefficients ~1e16
    // without the tolerance); the pivot test still calls it singular.
    expect(fitTrendline(spec({ type: 'poly', order: 2 }), [0.1, 0.1, 0.3, 0.3], [0.1, 1.1, 2.3, 0.3])).toMatchObject({ kind: 'none', code: 'chart-singular-trendline' })
  })

  it('intercept for linear, exponential and polynomial only', () => {
    // Linear through 0 (x 1..5, y 2,4,5,4,5): b = Σxy/Σx² = 66/55 = 1.2; R² uncentered (as LINEST const=FALSE): 1 - 6.8/86.
    const linear = fit({ type: 'linear', intercept: 0 }, range(5), [2, 4, 5, 4, 5])
    close(linear.coefficients, [0, 1.2]); expect(linear.r2).toBeCloseTo(1 - 6.8 / 86, 12)
    // Linear through 1: (y - 1) = b x -> b = (1+6+12+12+20)/55 = 51/55.
    close(fit({ type: 'linear', intercept: 1 }, range(5), [2, 4, 5, 4, 5]).coefficients, [1, 51 / 55])
    // Exponential with c = 1: ln y = b x at (1, e), (2, e³) -> b = (1 + 6)/5 = 1.4.
    close(fit({ type: 'exp', intercept: 1 }, [1, 2], [Math.E, Math.exp(3)]).coefficients, [1, 1.4])
    // Exponential with a non-positive intercept cannot be fitted.
    expect(fitTrendline(spec({ type: 'exp', intercept: 0 }), [1, 2], [1, 2])).toMatchObject({ kind: 'none' })
    // Polynomial order 2 through 5: y = 5 + 2x + 3x² recovered exactly; inexact data keeps b0 = 5.
    close(fit({ type: 'poly', order: 2, intercept: 5 }, range(4), range(4).map(x => 5 + 2 * x + 3 * x * x)).coefficients, [5, 2, 3])
    expect(fit({ type: 'poly', order: 2, intercept: 5 }, range(4), [1, 9, 3, 7]).coefficients[0]).toBe(5)
    // Logarithmic and power ignore the intercept: same coefficients as without, with a diagnostic.
    const logXs = [1, 2, 4, 8], logYs = [1, 3, 4, 8]
    close(fit({ type: 'log', intercept: 7 }, logXs, logYs).coefficients, fit({ type: 'log' }, logXs, logYs).coefficients, 12)
    close(fit({ type: 'power', intercept: 7 }, logXs, logYs).coefficients, fit({ type: 'power' }, logXs, logYs).coefficients, 12)
    const ignored = scatter([1, 2, 4, 8], [1, 3, 4, 8], trend(type('log') + '<intercept val="7"/>'))
    expect(codes(ignored)).toContain('chart-trendline-intercept-ignored')
    close(trendsOf(ignored)[0].coefficients, fit({ type: 'log' }, logXs, logYs).coefficients, 12)
    // Parsed from the part: linear through 0.
    close(trendsOf(scatter(range(5), [2, 4, 5, 4, 5], trend(type('linear') + '<intercept val="0"/>')))[0].coefficients, [0, 1.2])
  })

  it('extension bounded to 10x the data domain with a diagnostic', () => {
    // Data domain x 1..5 (4 units). forward 2, backward 1: drawn over [0, 7].
    const within = trendsOf(scatter(range(5), [2, 4, 5, 4, 5], trend(type('linear') + '<forward val="2"/><backward val="1"/>')))[0]
    expect(within.samples[0]).toBeCloseTo(0, 12)
    expect(within.samples[within.samples.length - 2]).toBeCloseTo(7, 12)
    // forward 100: the drawn domain is held to 10 x 4 = 40 units -> [1, 41], diagnosed.
    const far = scatter(range(5), [2, 4, 5, 4, 5], trend(type('linear') + '<forward val="100"/>'))
    const farLine = trendsOf(far)[0]
    expect(farLine.samples[0]).toBeCloseTo(1, 12)
    expect(farLine.samples[farLine.samples.length - 2]).toBeCloseTo(41, 12)
    expect(codes(far)).toContain('chart-trendline-extension-clamped')
    // Both sides scale together: forward 100, backward 20 -> 36 units shared 30 / 6 -> [-5, 35].
    const both = trendsOf(scatter(range(5), [2, 4, 5, 4, 5], trend(type('linear') + '<forward val="100"/><backward val="20"/>')))[0]
    expect(both.samples[0]).toBeCloseTo(-5, 12)
    expect(both.samples[both.samples.length - 2]).toBeCloseTo(35, 12)
    // Within bounds: no diagnostic. The y samples follow the fit (y = 2.2 + 0.6x).
    expect(codes(scatter(range(5), [2, 4, 5, 4, 5], trend(type('linear') + '<forward val="2"/>')))).not.toContain('chart-trendline-extension-clamped')
    expect(within.samples[1]).toBeCloseTo(2.2, 12)
    expect(within.samples[within.samples.length - 1]).toBeCloseTo(2.2 + 0.6 * 7, 12)
  })

  it('out-of-range order or period draws nothing, never clamped', () => {
    for (const order of ['1', '7', '2.5']) {
      const model = line([1, 4, 2, 5, 3, 6, 4, 8, 9], trend(type('poly') + `<order val="${order}"/>`))
      expect(trendsOf(model)).toHaveLength(0)
      expect(model.preparedAxes!.diagnostics.find(d => d.code === 'chart-trendline-parameter')?.reason).toMatch(/order/)
    }
    // Period 1 and period n (5 points): out of range; period n-1 = 4 draws.
    for (const period of ['1', '5']) {
      const model = line([1, 2, 3, 4, 5], trend(type('movingAvg') + `<period val="${period}"/>`))
      expect(trendsOf(model)).toHaveLength(0)
      expect(model.preparedAxes!.diagnostics.find(d => d.code === 'chart-trendline-parameter')?.reason).toMatch(/period/)
    }
    const ok = trendsOf(line([1, 2, 3, 4, 5], trend(type('movingAvg') + '<period val="4"/>')))[0]
    expect(Array.from(ok.samples)).toEqual([4, 2.5, 5, 3.5])
    expect(fitTrendline(spec({ type: 'poly', order: 7 }), range(10), range(10))).toMatchObject({ kind: 'none', code: 'chart-trendline-parameter' })
  })

  it('equation and R-squared labels: authored numFmt, else 4 significant digits', () => {
    const label = (body: string, xs = range(5), ys = [2, 4, 5, 4, 5]) => trendsOf(scatter(xs, ys, trend(body)))[0].label
    const shown = '<dispRSqr val="1"/><dispEq val="1"/>'
    expect(label(type('linear') + shown)).toBe('y = 0.6x + 2.2\nR² = 0.6')
    expect(label(type('linear') + shown + '<trendlineLbl><numFmt formatCode="0.00" sourceLinked="0"/></trendlineLbl>')).toBe('y = 0.60x + 2.20\nR² = 0.60')
    // 4 significant digits; a negative term reads "- 3", never "+ -3".
    expect(label(type('linear') + '<dispRSqr val="0"/><dispEq val="1"/>', [1, 2, 3], [1, 2, 2])).toBe('y = 0.5x + 0.6667')
    expect(label(type('linear') + '<dispRSqr val="0"/><dispEq val="1"/>', [1, 2, 3], [-1, 1, 3])).toBe('y = 2x - 3')
    expect(label(type('linear') + '<dispRSqr val="1"/><dispEq val="0"/>', [1, 2, 3], [1, 2, 2])).toBe('R² = 0.75')
    expect(label(type('poly') + '<order val="2"/><dispRSqr val="0"/><dispEq val="1"/>', range(4), range(4).map(x => 5 + 2 * x + 3 * x * x))).toBe('y = 3x² + 2x + 5')
    expect(label(type('exp') + '<dispRSqr val="0"/><dispEq val="1"/>', [0, 1, 2], [3, 3 * Math.exp(0.5), 3 * Math.E])).toBe('y = 3e^(0.5x)')
    expect(label(type('log') + '<dispRSqr val="0"/><dispEq val="1"/>', [1, 2, 4], [1, 1 + 2 * Math.log(2), 1 + 2 * Math.log(4)])).toBe('y = 2ln(x) + 1')
    expect(label(type('power') + '<dispRSqr val="0"/><dispEq val="1"/>', [1, 4, 9], [5, 40, 135])).toBe('y = 5x^1.5')
    // No label without dispEq/dispRSqr; moving averages have neither.
    expect(label(type('linear'))).toBeUndefined()
    expect(label(type('movingAvg') + '<period val="2"/>' + shown)).toBeUndefined()
    // Painted beside the line's end.
    const texts = paint(scatter(range(5), [2, 4, 5, 4, 5], trend(type('linear') + shown))).texts
    expect(texts).toEqual(expect.arrayContaining(['y = 0.6x + 2.2', 'R² = 0.6']))
  })

  it('at most 512 samples', () => {
    expect(TRENDLINE_SAMPLE_LIMIT).toBe(512)
    const red = '<spPr><ln w="19050"><solidFill><srgbClr val="FF0000"/></solidFill></ln></spPr>'
    const curve = scatter(range(2000), range(2000).map(x => Math.sin(x / 50) + x / 100), trend(type('poly') + '<order val="4"/>' + red))
    expect(trendsOf(curve)[0].samples.length).toBe(2 * 512)
    // A moving average of 2,000 points (1,999 averages) is decimated to 512, keeping its ends.
    const average = scatter(range(2000), range(2000).map(x => x % 7), trend(type('movingAvg') + '<period val="2"/>' + red))
    const samples = trendsOf(average)[0].samples
    expect(samples.length).toBe(2 * 512)
    expect(samples[0]).toBe(2); expect(samples[samples.length - 2]).toBe(2000)
    // The painted stroke has at most 511 segments, in the trendline's own colour, through the fit's ends.
    const painted = paint(curve)
    const strokes = painted.strokes.filter(s => s.style === '#FF0000')
    expect(strokes).toHaveLength(1)
    expect(strokes[0].segments.length).toBeLessThanOrEqual(511)
    expect(strokes[0].segments.length).toBeGreaterThan(400)
    expect(strokes[0].dash).toEqual([])
  })

  it('a 4097-point series fits on all points', () => {
    // 4,096 points on y = x, then one wild point. A display prefix (4,096) would fit slope exactly 1; the full
    // series gives the closed form b = Sxy / Sxx over x = 1..4097.
    const n = 4097, ys = range(n).map(x => x === n ? 1e6 : x)
    const model = line(ys, trend(type('linear')))
    expect(model.groups[0].series[0].analytic).toBe(true)
    expect(model.groups[0].series[0].values.length).toBe(n)
    const mean = (n + 1) / 2, yMean = ys.reduce((a, b) => a + b, 0) / n
    let sxy = 0, sxx = 0
    for (let i = 0; i < n; i++) { sxy += (i + 1 - mean) * (ys[i] - yMean); sxx += (i + 1 - mean) ** 2 }
    const slope = sxy / sxx
    const [a, b] = trendsOf(model)[0].coefficients
    expect(b).not.toBeCloseTo(1, 3)
    expect(b / slope).toBeCloseTo(1, 10)
    expect(a).toBeCloseTo(yMean - slope * mean, 6)
  })

  it('65537 points drop the trendline', () => {
    const model = line(range(65537), trend(type('linear')))
    expect(model.groups[0].series[0].analytic).toBe('refused')
    expect(trendsOf(model)).toHaveLength(0)
    expect(model.preparedAxes!.diagnostics.find(d => d.code === 'chart-trendline-dropped')?.reason).toMatch(/65,536/)
    // The series itself still draws (capped for display).
    expect(model.unsupported).toBeUndefined()
  })

  it('worksheet-filled series keep their trendlines and fit the filled cells', async () => {
    // No cache: the values come from Styled!A1:A5 (2,4,5,4,5) at fill time; the fit runs on the filled model.
    const xml = space(`<lineChart><grouping val="standard"/><ser><idx val="0"/><order val="0"/>${trend(type('linear') + '<dispEq val="1"/><dispRSqr val="1"/>')}<val><numRef><f>Styled!$A$1:$A$5</f></numRef></val></ser>${AX}</lineChart><catAx><axId val="1"/><crossAx val="2"/></catAx><valAx><axId val="2"/><crossAx val="1"/></valAx>`)
    const sheet = (await parseXlsx(await chartWorkbook(xml, { cells: [2, 4, 5, 4, 5].map((v, i) => ({ ref: `A${i + 1}`, v })) }))).sheets[0]
    const content = sheet.drawings?.[0].content
    expect(content?.kind).toBe('xlsxChart')
    const model = (content as { chart: ChartModel }).chart
    expect(model.groups[0].series[0].trendlines?.[0].type).toBe('linear')
    close(trendsOf(model)[0].coefficients, [2.2, 0.6])
    expect(trendsOf(model)[0].label).toBe('y = 0.6x + 2.2\nR² = 0.6')
  })

  it('charts without trendlines carry no trendline record; stacked and pie groups draw none', () => {
    const plain = line([1, 2, 3], '')
    expect(plain.groups[0].series[0].trendlines).toBeUndefined()
    expect(plain.preparedAxes && 'trendlines' in plain.preparedAxes).toBe(false)
    const stacked = line([1, 2, 3], trend(type('linear')), 'stacked')
    expect(trendsOf(stacked)).toHaveLength(0)
    expect(codes(stacked)).toContain('chart-trendline-dropped')
  })

  it('draws on category axes at category positions with the default trendline stroke', () => {
    // A line chart 2,4,6 (slope 2 through 0 at x = 0): the trendline coincides with the series at each category.
    const model = line([2, 4, 6], trend(type('linear')))
    const painted = paint(model)
    const color = `#${model.groups[0].series[0].color}`
    const strokes = painted.strokes.filter(s => s.style === color)
    // The series line (solid), then the trendline (dotted, 2 px; Excel 2013+ writes sysDot at 19050 EMU).
    const dotted = strokes.filter(s => s.dash.length > 0)
    expect(dotted).toHaveLength(1)
    const t = dotted[0]
    expect(t.width).toBe(2)
    const series = strokes[0].segments
    expect(t.segments[0].x1).toBeCloseTo(series[0].x1, 6)
    expect(t.segments[0].y1).toBeCloseTo(series[0].y1, 6)
    expect(t.segments[t.segments.length - 1].x2).toBeCloseTo(series[series.length - 1].x2, 6)
    expect(t.segments[t.segments.length - 1].y2).toBeCloseTo(series[series.length - 1].y2, 6)
  })
})

const W = 400, H = 260
function paint(model: ChartModel) {
  const canvas = createCanvas(W, H), ctx = canvas.getContext('2d') as unknown as CanvasRenderingContext2D
  const kind = (s: unknown) => typeof s === 'string' ? s.toUpperCase() : (s as object).constructor.name
  const segments: Array<{ x1: number; y1: number; x2: number; y2: number }> = []
  const strokes: Array<{ style: string; width: number; dash: number[]; segments: typeof segments }> = []
  const texts: string[] = []
  let path: typeof segments = [], pen: [number, number] = [0, 0]
  const o = { moveTo: ctx.moveTo.bind(ctx), lineTo: ctx.lineTo.bind(ctx), beginPath: ctx.beginPath.bind(ctx), stroke: ctx.stroke.bind(ctx), fillText: ctx.fillText.bind(ctx) }
  ctx.beginPath = () => { path = []; o.beginPath() }
  ctx.moveTo = (x, y) => { pen = [x, y]; o.moveTo(x, y) }
  ctx.lineTo = (x, y) => { path.push({ x1: pen[0], y1: pen[1], x2: x, y2: y }); pen = [x, y]; o.lineTo(x, y) }
  ctx.stroke = (...a: []) => { strokes.push({ style: kind(ctx.strokeStyle), width: ctx.lineWidth, dash: ctx.getLineDash(), segments: path }); o.stroke(...a) }
  ctx.fillText = (text, x, y, max) => { texts.push(text); o.fillText(text, x, y, max) }
  paintXlsxChart(model, ctx, W, H)
  const layout = layoutChart(model, W, H, 1, 0, { width: (text: string, css: string) => { ctx.font = css; return ctx.measureText(text).width }, lineHeight: () => model.fontSizePt * 96 / 72 })
  return { strokes, texts, layout }
}
