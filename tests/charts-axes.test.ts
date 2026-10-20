// @vitest-environment node
/**
 * A1a: the axis graph and scales (spec 5.1, 6.5 tick bound). Numeric model tests over the parsed chart and its
 * resolved axes, plus one painter check that the drawing follows the graph. Fixtures are Office-shaped
 * synthetic (the secondary-axis plot mirrors the shape Excel writes), not Office evidence.
 */
import { describe, expect, it } from 'vitest'
import { createCanvas } from 'canvas'
import { parseXmlOrdered } from '../src/core/xml'
import { axisFraction, categoryFraction, resolveAxes } from '../src/charts/axes'
import { chartModelOwner, paintXlsxChart, parseXlsxChart, pointBytes, releaseChartModel, stringBytes } from '../src/charts'
import { createChartWorkspace } from '../src/charts/workspace'
import type { ChartParseContext } from '../src/charts/context'
import type { ChartModel } from '../src/charts/model'

const C = 'http://schemas.openxmlformats.org/drawingml/2006/chart'
const A = 'http://schemas.openxmlformats.org/drawingml/2006/main'
const space = (plot: string) => `<c:chartSpace xmlns:c="${C}" xmlns:a="${A}"><c:chart><c:autoTitleDeleted val="1"/><c:plotArea>${plot}</c:plotArea></c:chart></c:chartSpace>`
const lit = (values: number[]) => `<c:numLit><c:ptCount val="${values.length}"/>${values.map((v, i) => `<c:pt idx="${i}"><c:v>${v}</c:v></c:pt>`).join('')}</c:numLit>`
const cats = (n: number) => `<c:cat><c:strLit><c:ptCount val="${n}"/>${Array.from({ length: n }, (_, i) => `<c:pt idx="${i}"><c:v>C${i + 1}</c:v></c:pt>`).join('')}</c:strLit></c:cat>`
const color = (c: string) => `<c:spPr><a:solidFill><a:srgbClr val="${c}"/></a:solidFill><a:ln><a:solidFill><a:srgbClr val="${c}"/></a:solidFill></a:ln></c:spPr>`
const ser = (i: number, values: number[], c = 'FF0000') => `<c:ser><c:idx val="${i}"/><c:order val="${i}"/>${color(c)}${cats(values.length)}<c:val>${lit(values)}</c:val></c:ser>`
const xy = (i: number, xs: number[], ys: number[]) => `<c:ser><c:idx val="${i}"/><c:order val="${i}"/><c:xVal>${lit(xs)}</c:xVal><c:yVal>${lit(ys)}</c:yVal></c:ser>`
const catAx = (id: string, cross: string, extra = '') => `<c:catAx><c:axId val="${id}"/>${extra}<c:crossAx val="${cross}"/></c:catAx>`
const valAx = (id: string, cross: string, extra = '') => `<c:valAx><c:axId val="${id}"/>${extra}<c:crossAx val="${cross}"/></c:valAx>`
const bar = (series: string, ids = ['1', '2'], dir = 'col', grouping = 'clustered') => `<c:barChart><c:barDir val="${dir}"/><c:grouping val="${grouping}"/>${series}${ids.map(id => `<c:axId val="${id}"/>`).join('')}</c:barChart>`
const line = (series: string, ids = ['1', '2']) => `<c:lineChart><c:grouping val="standard"/>${series}${ids.map(id => `<c:axId val="${id}"/>`).join('')}</c:lineChart>`
const parse = (plot: string): ChartModel => parseXlsxChart(parseXmlOrdered(space(plot)), undefined)!
/** Excel's secondary-axis shape: catAx 1 <-> valAx 2 (l); deleted catAx 3 <-> valAx 4 (r, crosses max). */
const SECONDARY = bar(ser(0, [2, 4, 6, 8, 10])) + line(ser(1, [100, 400, 250, 900, 1000], '0000FF'), ['3', '4'])
  + catAx('1', '2', '<c:axPos val="b"/>') + valAx('2', '1', '<c:axPos val="l"/><c:majorGridlines/><c:crossBetween val="between"/>')
  + catAx('3', '4', '<c:axPos val="b"/><c:delete val="1"/>') + valAx('4', '3', '<c:axPos val="r"/><c:crosses val="max"/><c:crossBetween val="between"/>')

describe('axis graph (spec 5.1)', () => {
  it('groups bind axes by axId; a secondary value axis has its own scale', () => {
    const model = parse(SECONDARY)
    expect(model.unsupported).toBeUndefined()
    expect(model.groups.map(g => g.binding)).toEqual([{ kind: 'cartesian', category: '1', value: '2' }, { kind: 'cartesian', category: '3', value: '4' }])
    const graph = resolveAxes(model)
    expect(graph.groups.map(g => [g!.category?.id, g!.value.id])).toEqual([['1', '2'], ['3', '4']])
    const primary = graph.axes.get('2')!, secondary = graph.axes.get('4')!
    expect([primary.min, primary.max, primary.major]).toEqual([0, 10, 2])
    expect([secondary.min, secondary.max, secondary.major]).toEqual([0, 1000, 200])
    // The secondary value axis crosses its (deleted) category axis at its maximum: drawn on the right.
    expect(secondary.side).toBe('r')
    expect(primary.side).toBe('l')
    expect(graph.axes.get('3')!.deleted).toBe(true)

    // The painter draws each group against its own value axis: the primary bars span the plot height
    // (a shared 0..1000 scale would make them a few pixels tall).
    const canvas = createCanvas(400, 260), ctx = canvas.getContext('2d')
    paintXlsxChart(model, ctx as unknown as CanvasRenderingContext2D, 400, 260)
    const { data } = ctx.getImageData(0, 0, 400, 260)
    let top = 260, bottom = 0
    for (let y = 0; y < 260; y++) for (let x = 0; x < 400; x++) {
      const k = (y * 400 + x) * 4
      if (data[k] > 200 && data[k + 1] < 60 && data[k + 2] < 60) { top = Math.min(top, y); bottom = Math.max(bottom, y) }
    }
    expect(bottom - top).toBeGreaterThan(150)
  })

  it('a deleted axis still controls the scale', () => {
    const model = parse(bar(ser(0, [5, 12, 30])) + catAx('1', '2', '<c:delete val="1"/>') + valAx('2', '1', '<c:delete val="1"/><c:scaling><c:orientation val="minMax"/><c:max val="50"/><c:min val="-10"/></c:scaling><c:majorUnit val="20"/>'))
    const value = resolveAxes(model).axes.get('2')!
    expect(value.deleted).toBe(true)
    expect([value.min, value.max, value.major]).toEqual([-10, 50, 20])
    expect(value.ticks).toEqual([-10, 10, 30, 50])
    expect(axisFraction(value, 20)).toBeCloseTo(0.5)
  })

  it('crosses autoZero, min, max and crossesAt place the crossing', () => {
    const at = (crossing: string, values = [-4, 6, 8]) => {
      const graph = resolveAxes(parse(bar(ser(0, values)) + catAx('1', '2', crossing) + valAx('2', '1')))
      return graph.axes.get('1')!.crossesValue
    }
    expect(at('<c:crosses val="autoZero"/>')).toBe(0)
    expect(at('')).toBe(0)
    // The scale of [-4, 6, 8] is -5..10 (step 5).
    expect(at('<c:crosses val="min"/>')).toBe(-5)
    expect(at('<c:crosses val="max"/>')).toBe(10)
    expect(at('<c:crossesAt val="5"/>')).toBe(5)
    // autoZero outside the range crosses at the nearer end; crossesAt is clamped into the scale.
    expect(at('<c:crosses val="autoZero"/>', [20, 40, 60])).toBe(0)
    expect(at('<c:crossesAt val="500"/>')).toBe(10)
    // The value axis crossing the category axis: min (left), max (right), crossesAt a category number.
    const valueAt = (crossing: string) => resolveAxes(parse(bar(ser(0, [1, 2, 3, 4])) + catAx('1', '2') + valAx('2', '1', crossing))).axes.get('2')!.crossesFraction
    expect(valueAt('<c:crosses val="autoZero"/>')).toBe(0)
    expect(valueAt('<c:crosses val="max"/>')).toBe(1)
    expect(valueAt('<c:crossesAt val="3"/>')).toBeCloseTo(2.5 / 4)
  })

  it('explicit min and max, majorUnit and minorUnit, minor gridlines and tick marks, tickLblPos nextTo, low, high and none', () => {
    const model = parse(bar(ser(0, [13, 47, 88])) + catAx('1', '2', '<c:majorTickMark val="out"/><c:minorTickMark val="none"/><c:tickLblPos val="low"/>')
      + valAx('2', '1', '<c:scaling><c:max val="100"/><c:min val="0"/></c:scaling><c:majorGridlines/><c:minorGridlines/><c:majorTickMark val="cross"/><c:minorTickMark val="in"/><c:tickLblPos val="nextTo"/><c:majorUnit val="25"/><c:minorUnit val="5"/>'))
    const graph = resolveAxes(model)
    const value = graph.axes.get('2')!, category = graph.axes.get('1')!
    expect(value.ticks).toEqual([0, 25, 50, 75, 100])
    expect(value.minorTicks).toEqual([5, 10, 15, 20, 30, 35, 40, 45, 55, 60, 65, 70, 80, 85, 90, 95])
    expect([value.axis.gridlines, value.axis.minorGridlines]).toEqual([true, true])
    expect([value.axis.majorTickMark, value.axis.minorTickMark, value.axis.tickLblPos]).toEqual(['cross', 'in', 'nextTo'])
    expect([category.axis.majorTickMark, category.axis.minorTickMark, category.axis.tickLblPos]).toEqual(['out', 'none', 'low'])
    for (const pos of ['high', 'none'] as const) expect(parse(bar(ser(0, [1])) + catAx('1', '2', `<c:tickLblPos val="${pos}"/>`) + valAx('2', '1')).axes[0].tickLblPos).toBe(pos)
    // Without a minorUnit the minor unit is a fifth of the major unit.
    const auto = resolveAxes(parse(bar(ser(0, [1, 9])) + catAx('1', '2') + valAx('2', '1', '<c:minorGridlines/>'))).axes.get('2')!
    expect(auto.minor).toBeCloseTo(auto.major / 5)
  })

  it('xy bindings keep independent x and y domains; a category group plus an XY group keep separate x domains', () => {
    const scatter = `<c:scatterChart><c:scatterStyle val="marker"/>${xy(0, [1000, 2000, 3000], [0.1, 0.2, 0.5])}<c:axId val="10"/><c:axId val="20"/></c:scatterChart>`
    const model = parse(scatter + valAx('10', '20', '<c:axPos val="b"/>') + valAx('20', '10', '<c:axPos val="l"/>'))
    expect(model.groups[0].binding).toEqual({ kind: 'xy', x: '10', y: '20' })
    const graph = resolveAxes(model)
    expect([graph.axes.get('10')!.min, graph.axes.get('10')!.max]).toEqual([0, 3000])
    expect([graph.axes.get('20')!.min, graph.axes.get('20')!.max]).toEqual([0, 0.5])
    expect(graph.axes.get('10')!.orientation).toBe('h')
    expect(graph.axes.get('20')!.orientation).toBe('v')
    // axPos decides which value axis is x, not the axId order.
    const swapped = parse(scatter.replace('<c:axId val="10"/><c:axId val="20"/>', '<c:axId val="20"/><c:axId val="10"/>') + valAx('10', '20', '<c:axPos val="b"/>') + valAx('20', '10', '<c:axPos val="l"/>'))
    expect(swapped.groups[0].binding).toEqual({ kind: 'xy', x: '10', y: '20' })

    const combo = parse(bar(ser(0, [1, 2, 3])) + catAx('1', '2', '<c:axPos val="b"/>') + valAx('2', '1', '<c:axPos val="l"/>')
      + scatter.replace('<c:axId val="10"/><c:axId val="20"/>', '<c:axId val="30"/><c:axId val="40"/>')
      + valAx('30', '40', '<c:axPos val="t"/><c:crosses val="max"/>') + valAx('40', '30', '<c:axPos val="r"/><c:crosses val="max"/>'))
    expect(combo.unsupported).toBeUndefined()
    const both = resolveAxes(combo)
    expect(both.groups[0]!.category!.role).toBe('category')
    expect(both.groups[0]!.category!.slots).toBe(3)
    expect(both.groups[1]!.x!.id).toBe('30')
    expect([both.axes.get('30')!.min, both.axes.get('30')!.max]).toEqual([0, 3000])
    expect([both.axes.get('2')!.max, both.axes.get('40')!.max]).toEqual([3, 0.5])
  })

  it('an unplaceable combination is chart-axis-unplaceable', () => {
    // A horizontal bar group and a line group on the same axes need the category axis both vertical and horizontal.
    const mixed = parse(bar(ser(0, [1, 2]), ['1', '2'], 'bar') + line(ser(1, [3, 4])) + catAx('1', '2') + valAx('2', '1'))
    expect(mixed.unsupported).toMatch(/axes/i)
    expect(mixed.diagnostics?.map(d => d.code)).toContain('chart-axis-unplaceable')
    // An axId that names no axis, or the wrong kind of axis.
    for (const plot of [bar(ser(0, [1]), ['1', '9']) + catAx('1', '2') + valAx('2', '1'), bar(ser(0, [1]), ['1', '2']) + catAx('1', '2') + catAx('2', '1')]) {
      const model = parse(plot)
      expect(model.unsupported).toBeDefined()
      expect(model.diagnostics?.map(d => d.code)).toContain('chart-axis-unplaceable')
    }
    // A bound pair whose crossAx name other axes.
    const crossed = parse(bar(ser(0, [1])) + catAx('1', '7') + valAx('2', '1') + valAx('7', '1'))
    expect(crossed.diagnostics?.map(d => d.code)).toContain('chart-axis-unplaceable')
    // A plot area without axes (and groups without axIds) keeps an implied primary pair.
    const implicit = parse(bar(ser(0, [1, 2]), []))
    expect(implicit.unsupported).toBeUndefined()
    expect(implicit.axes.map(axis => [axis.kind, axis.implicit])).toEqual([['category', true], ['value', true]])
  })

  it('crossBetween between vs midCat', () => {
    const graph = (between: string) => resolveAxes(parse(line(ser(0, [1, 2, 3, 4])) + catAx('1', '2') + valAx('2', '1', between)))
    const b = graph('<c:crossBetween val="between"/>'), m = graph('<c:crossBetween val="midCat"/>')
    expect(b.groups[0]!.between).toBe(true)
    expect(m.groups[0]!.between).toBe(false)
    const cb = b.axes.get('1')!, cm = m.axes.get('1')!
    expect([0, 3].map(i => categoryFraction(cb, i, true))).toEqual([0.125, 0.875])
    expect([0, 3].map(i => categoryFraction(cm, i, false))).toEqual([0, 1])
  })

  it('orientation maxMin', () => {
    const graph = resolveAxes(parse(bar(ser(0, [2, 4, 6, 8])) + catAx('1', '2', '<c:scaling><c:orientation val="maxMin"/></c:scaling>') + valAx('2', '1', '<c:scaling><c:orientation val="maxMin"/></c:scaling>')))
    const category = graph.axes.get('1')!, value = graph.axes.get('2')!
    expect([category.reversed, value.reversed]).toEqual([true, true])
    expect(axisFraction(value, value.max)).toBe(0)
    expect(axisFraction(value, value.min)).toBe(1)
    expect(categoryFraction(category, 0, true)).toBe(0.875)
    // A reversed category axis moves the value axis to the far side.
    expect(value.side).toBe('r')
  })

  it('logBase skips non-positive values with chart-invalid-log-value', () => {
    const graph = resolveAxes(parse(line(ser(0, [-1, 0, 10, 250, 1000])) + catAx('1', '2') + valAx('2', '1', '<c:scaling><c:logBase val="10"/></c:scaling>')))
    const value = graph.axes.get('2')!
    expect(value.log).toBe(10)
    expect([value.min, value.max]).toEqual([10, 1000])
    expect(value.ticks).toEqual([10, 100, 1000])
    expect(axisFraction(value, 100)).toBeCloseTo(0.5)
    expect([value.plottable(-1), value.plottable(0), value.plottable(10)]).toEqual([false, false, true])
    const diagnostics = graph.diagnostics.filter(d => d.code === 'chart-invalid-log-value')
    expect(diagnostics).toHaveLength(1)
    expect(diagnostics[0].reason).toMatch(/2 /)
    // Out of the 2..1000 range: the axis stays linear.
    expect(resolveAxes(parse(line(ser(0, [1, 2])) + catAx('1', '2') + valAx('2', '1', '<c:scaling><c:logBase val="1"/></c:scaling>'))).axes.get('2')!.log).toBeUndefined()
  })

  it('display units divide tick values and name the unit label', () => {
    const unit = (extra: string, values = [1.5e6, 4e6]) => resolveAxes(parse(bar(ser(0, values)) + catAx('1', '2') + valAx('2', '1', extra))).axes.get('2')!
    const millions = unit('<c:dispUnits><c:builtInUnit val="millions"/><c:dispUnitsLbl/></c:dispUnits>')
    expect(millions.display).toEqual({ divisor: 1e6, label: 'Millions' })
    expect(millions.ticks.map(t => millions.labelValue(t))).toEqual([0, 1, 2, 3, 4])
    // Without dispUnitsLbl the values are still divided, but no unit label is drawn.
    expect(unit('<c:dispUnits><c:builtInUnit val="thousands"/></c:dispUnits>').display).toEqual({ divisor: 1000 })
    for (const [name, divisor] of [['hundreds', 1e2], ['tenThousands', 1e4], ['hundredThousands', 1e5], ['tenMillions', 1e7], ['hundredMillions', 1e8], ['billions', 1e9], ['trillions', 1e12]] as const) {
      expect(unit(`<c:dispUnits><c:builtInUnit val="${name}"/></c:dispUnits>`).display?.divisor).toBe(divisor)
    }
    const custom = unit('<c:dispUnits><c:custUnit val="250"/><c:dispUnitsLbl/></c:dispUnits>', [500, 1000])
    expect(custom.display?.divisor).toBe(250)
    expect(custom.labelValue(1000)).toBe(4)
    expect(custom.display?.label).toBeTruthy()
    // An authored unit label text wins over the default.
    const authored = unit('<c:dispUnits><c:builtInUnit val="billions"/><c:dispUnitsLbl><c:tx><c:rich><a:p><a:r><a:t>USD bn</a:t></a:r></a:p></c:rich></c:tx></c:dispUnitsLbl></c:dispUnits>')
    expect(authored.display).toEqual({ divisor: 1e9, label: 'USD bn' })
  })

  it('the axis records and date categories the model retains are charged to the chart workspace', () => {
    const workspace = createChartWorkspace('desktop')
    const context = { workspace, profile: 'desktop', pkg: undefined, partPath: 'chart1.xml' } as unknown as ChartParseContext
    const xml = space(line(`<c:ser><c:idx val="0"/><c:order val="0"/><c:cat><c:numLit><c:ptCount val="3"/><c:pt idx="0"><c:v>45000</c:v></c:pt><c:pt idx="1"><c:v>45001</c:v></c:pt><c:pt idx="2"><c:v>45010</c:v></c:pt></c:numLit></c:cat><c:val>${lit([1, 2, 3])}</c:val></c:ser>`)
      + '<c:dateAx><c:axId val="1"/><c:title><c:tx><c:rich><a:p><a:r><a:t>When</a:t></a:r></a:p></c:rich></c:tx></c:title><c:crossAx val="2"/></c:dateAx>' + valAx('2', '1', '<c:dispUnits><c:builtInUnit val="thousands"/><c:dispUnitsLbl/></c:dispUnits>'))
    const model = parseXlsxChart(parseXmlOrdered(xml), undefined, context)!
    expect(model.categoryValues?.length).toBe(3)
    const owner = chartModelOwner(model)!
    expect(owner.workspace).toBe(workspace)
    // Values and dates (9 B per point each), the three category labels, and two axis records with their strings.
    const strings = ['45000', '45001', '45010', '1', 'When', '2', 'Thousands', 'Thousands'].reduce((sum, text) => sum + stringBytes(text), 0)
    expect(workspace.used).toBe(pointBytes(3) * 2 + 2 * 256 + strings)
    releaseChartModel(model)
    expect(workspace.used).toBe(0)
    // A workspace without room refuses the chart before the axis records are kept.
    const tight = createChartWorkspace('desktop')
    tight.reserve(tight.limit - 300, 'pressure')
    const refused = parseXlsxChart(parseXmlOrdered(xml), undefined, { ...context, workspace: tight } as ChartParseContext)!
    expect(refused.refused).toBe(true)
    expect(tight.used).toBe(tight.limit - 300)
  })

  it('ticks are bounded to 1000; a tiny majorUnit is enlarged with a diagnostic', () => {
    const graph = resolveAxes(parse(bar(ser(0, [0, 100])) + catAx('1', '2') + valAx('2', '1', '<c:scaling><c:max val="100"/><c:min val="0"/></c:scaling><c:majorUnit val="0.001"/><c:minorUnit val="0.0001"/>')))
    const value = graph.axes.get('2')!
    expect(value.ticks.length).toBeLessThanOrEqual(1000)
    expect(value.ticks.length).toBeGreaterThan(1)
    expect(value.major).toBeGreaterThan(0.001)
    expect(value.minorTicks.length).toBeLessThanOrEqual(1000)
    expect(graph.diagnostics.map(d => d.code)).toContain('chart-tick-limit')
    // The authored unit is a rendering bound, never rewritten on the model.
    expect(value.axis.majorUnit).toBe(0.001)
    // An extreme finite scale is bounded the same way.
    const extreme = resolveAxes(parse(bar(ser(0, [0, 1e300])) + catAx('1', '2') + valAx('2', '1', '<c:majorUnit val="1"/>'))).axes.get('2')!
    expect(extreme.ticks.length).toBeLessThanOrEqual(1000)
  })
})
