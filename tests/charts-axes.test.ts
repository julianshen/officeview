// @vitest-environment node
/**
 * A1a: the axis graph and scales (spec 5.1, 6.5 tick bound). Numeric model tests over the parsed chart and its
 * resolved axes, plus one painter check that the drawing follows the graph. Fixtures are Office-shaped
 * synthetic (the secondary-axis plot mirrors the shape Excel writes), not Office evidence.
 */
import { describe, expect, it } from 'vitest'
import { createCanvas } from 'canvas'
import { parseXmlOrdered } from '../src/core/xml'
import { axisFraction, axisGraphBound, categoryFraction, resolveAxes } from '../src/charts/axes'
import { ChartRefusal } from '../src/charts/own'
import { CHART_STANDALONE_PAINT_PROFILE, receivedChartContext } from '../src/charts/paint'
import { documentChartContext } from '../src/charts/context'
import { ensureChartEngine, ensureChartPainter } from '../src/xlsx/chartLoader'
import { parseXlsx } from '../src/xlsx/parse'
import { chartWorkbook } from './helpers/xlsx-chart-fixture'
import { mainChartParent, prepareSheetAssets } from '../src/components/xlsx/sheetAssets'
import { MainMemoryLedger } from '../src/xlsx/view/mainLedger'
import { MAIN_THREAD_BUDGETS } from '../src/xlsx/view/tileCache'
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

/** A context that records every call (and still draws). */
function recorder(w = 800, h = 500) {
  const canvas = createCanvas(w, h), target = canvas.getContext('2d')
  const calls: Array<[string, unknown[]]> = []
  const ctx = new Proxy(target, {
    get(t, p) { const v = Reflect.get(t, p); return typeof v === 'function' ? (...args: unknown[]) => { calls.push([String(p), args]); return v.apply(t, args) } : v },
    set(t, p, v) { return Reflect.set(t, p, v) },
  }) as unknown as CanvasRenderingContext2D
  return { ctx, calls, texts: () => calls.filter(([name]) => name === 'fillText').map(([, args]) => String(args[0])), of: (name: string) => calls.filter(([n]) => n === name).map(([, args]) => args as number[]) }
}
const dateLine = (dateAx: string, family: 'line' | 'bar' = 'line') => {
  const s = `<c:ser><c:idx val="0"/><c:order val="0"/><c:cat>${lit([45000, 45001, 45002])}</c:cat><c:val>${lit([1, 2, 3])}</c:val></c:ser>`
  const group = family === 'line' ? `<c:lineChart><c:grouping val="standard"/>${s}<c:axId val="1"/><c:axId val="2"/></c:lineChart>` : `<c:barChart><c:barDir val="bar"/><c:grouping val="clustered"/>${s}<c:axId val="1"/><c:axId val="2"/></c:barChart>`
  return parse(`${group}<c:dateAx><c:axId val="1"/><c:baseTimeUnit val="days"/>${dateAx}<c:crossAx val="2"/></c:dateAx>` + valAx('2', '1'))
}

describe('A1a fix round 1: admission, label order, minor elements, extreme scales', () => {
  it('I3: graph and tick construction is admitted to the chart workspace before it is built, refused when full, refunded after paint', () => {
    const plot = bar(ser(0, [3, 7, 2])) + catAx('1', '2') + valAx('2', '1', '<c:majorGridlines/><c:minorUnit val="0.5"/><c:minorGridlines/>')
    const workspace = createChartWorkspace('desktop')
    const context = { workspace, profile: 'desktop', partPath: 'chart1.xml' } as unknown as ChartParseContext
    const model = parseXlsxChart(parseXmlOrdered(space(plot)), undefined, context)!
    const held = workspace.used
    // With room: the paint reserves its construction bound first, and refunds it after the paint.
    const log: string[] = []
    const reserve = workspace.reserve.bind(workspace)
    workspace.reserve = (bytes, tag) => { log.push(`${tag}:${bytes}`); return reserve(bytes, tag) }
    const drawn = recorder()
    paintXlsxChart(model, drawn.ctx, 800, 500)
    expect(log.length).toBe(1)
    expect(log[0]).toMatch(/:paint:\d+$/)
    expect(Number(log[0].split(':').pop())).toBeGreaterThanOrEqual(axisGraphBound(model, 3))
    expect(workspace.used).toBe(held)
    expect(drawn.of('fillRect').length).toBeGreaterThan(3)
    // Full: resolveAxes refuses before building, and the paint is the contained card (no native bars), nothing held.
    workspace.reserve(workspace.limit - workspace.used, 'pressure')
    expect(() => resolveAxes(model, undefined, undefined, { workspace, tag: 'probe' })).toThrow(ChartRefusal)
    const full = recorder()
    paintXlsxChart(model, full.ctx, 800, 500)
    expect(full.texts().join(' ')).toMatch(/memory/)
    expect(full.texts()).not.toContain('C1')
    expect(workspace.used).toBe(workspace.limit)
    workspace.release('pressure')
    expect(workspace.used).toBe(held)
    // An unowned model (a structured clone) with nothing received paints in a bounded per-call workspace (R2).
    const clone = recorder()
    paintXlsxChart(structuredClone(model), clone.ctx, 800, 500)
    expect(clone.of('fillRect').length).toBe(drawn.of('fillRect').length)
  })

  it('I4: date labels are thinned in screen order: a reversed horizontal axis and both vertical orientations keep every label that fits', () => {
    const labels = ['3/15/2023', '3/16/2023', '3/17/2023']
    for (const model of [dateLine(''), dateLine('<c:scaling><c:orientation val="maxMin"/></c:scaling>'), dateLine('', 'bar'), dateLine('<c:scaling><c:orientation val="maxMin"/></c:scaling>', 'bar')]) {
      const { ctx, texts } = recorder()
      paintXlsxChart(model, ctx, 800, 500)
      expect(texts().filter(t => labels.includes(t)).sort()).toEqual(labels)
    }
  })

  it('I5: minor ticks and minor gridlines are drawn on scatter x and date axes', () => {
    const lines = (model: ChartModel) => { const r = recorder(); paintXlsxChart(model, r.ctx, 800, 500); return r.of('moveTo').length }
    const scatter = (extra: string) => parse(`<c:scatterChart><c:scatterStyle val="marker"/>${xy(0, [0, 1, 2], [1, 2, 3])}<c:axId val="10"/><c:axId val="20"/></c:scatterChart>`
      + valAx('10', '20', `<c:axPos val="b"/><c:scaling><c:max val="2"/><c:min val="0"/></c:scaling>${extra}<c:majorUnit val="1"/>`) + valAx('20', '10', '<c:axPos val="l"/>'))
    const plain = scatter(''), minor = scatter('<c:minorGridlines/><c:minorTickMark val="out"/><c:minorUnit val="0.25"/>')
    const minors = resolveAxes(minor).axes.get('10')!.minorTicks
    expect(minors).toEqual([0.25, 0.5, 0.75, 1.25, 1.5, 1.75])
    // One gridline and one tick mark per minor tick.
    expect(lines(minor) - lines(plain)).toBe(2 * minors.length)
    // Their geometry: the gridlines sit at the minor fractions of the plot width.
    const r = recorder(); paintXlsxChart(minor, r.ctx, 800, 500)
    const xs = new Set(r.of('moveTo').map(([x]) => x))
    const gx = r.of('moveTo').map(([x]) => x).filter(x => Number.isInteger(x - 0.5))
    expect(gx.length).toBeGreaterThanOrEqual(minors.length)
    expect(xs.size).toBeGreaterThan(minors.length)
    // Date axes: minor ticks from minorUnit/minorTimeUnit, drawn as gridlines and marks.
    const dPlain = dateLine(''), dMinor = dateLine('<c:minorGridlines/><c:minorTickMark val="cross"/><c:minorUnit val="1"/><c:minorTimeUnit val="days"/>')
    const dTicks = resolveAxes(dMinor).axes.get('1')!.minorTicks
    expect(dTicks).toEqual([45000, 45001, 45002])
    expect(lines(dMinor) - lines(dPlain)).toBe(2 * dTicks.length)
  })

  it('I6: extreme finite linear ranges normalise without overflow and paint finite coordinates', () => {
    const model = parse(line(ser(0, [-1e308, 0, 1e308])) + catAx('1', '2') + valAx('2', '1', '<c:scaling><c:max val="1e308"/><c:min val="-1e308"/></c:scaling>'))
    const value = resolveAxes(model).axes.get('2')!
    expect([-1e308, 0, 1e308].map(v => axisFraction(value, v))).toEqual([0, 0.5, 1])
    expect(value.ticks.every(Number.isFinite)).toBe(true)
    const auto = resolveAxes(parse(line(ser(0, [-1e308, 1e308])) + catAx('1', '2') + valAx('2', '1'))).axes.get('2')!
    expect([auto.min, auto.max].every(Number.isFinite)).toBe(true)
    expect(axisFraction(auto, 0)).toBeCloseTo(0.5)
    for (const m of [model, parse(line(ser(0, [-1e308, 1e308])) + catAx('1', '2') + valAx('2', '1'))]) {
      const r = recorder()
      paintXlsxChart(m, r.ctx, 800, 500)
      const coords = [...r.of('moveTo'), ...r.of('lineTo'), ...r.of('fillRect'), ...r.of('arc')].flat().filter(v => typeof v === 'number')
      expect(coords.length).toBeGreaterThan(0)
      expect(coords.every(Number.isFinite)).toBe(true)
    }
  })
})

describe('A1a fix round 2: measured construction, receiving-document budgets, one-sided extremes', () => {
  /** Paints while recording every per-slot array the construction allocates (length >= `slots`). */
  function measured(model: ChartModel, slots: number, paint: () => void): number {
    let bytes = 0
    const note = (length: number) => { if (length >= slots) bytes += length * 8 }
    const g = globalThis as unknown as Record<string, unknown>
    const F64 = Float64Array, Arr = Array, from = Array.from, slice = Array.prototype.slice
    class SpyF64 extends F64 { constructor(...args: unknown[]) { super(...(args as unknown as [number])); note(this.length) } }
    class SpyArray extends Arr<unknown> { constructor(...args: unknown[]) { super(...(args as unknown as [number])); note(this.length) } }
    g.Float64Array = SpyF64; g.Array = SpyArray
    Array.from = function (this: unknown, ...args: Parameters<typeof Array.from>) { const out = (from as (...a: unknown[]) => unknown[]).apply(Arr, args); note(out.length); return out } as typeof Array.from
    Array.prototype.slice = function (this: unknown[], start?: number, end?: number) { const out = slice.call(this, start, end); note(out.length); return out }
    try { paint() } finally { g.Float64Array = F64; g.Array = Arr; Array.from = from; Array.prototype.slice = slice }
    void model
    return bytes
  }
  const many = (family: 'line' | 'area' | 'scatter', grouping = 'standard') => {
    const values = Array.from({ length: 4096 }, (_, i) => (i * 7) % 13)
    const series = Array.from({ length: 64 }, (_, k) => family === 'scatter' ? xy(k, values.map((_, i) => i), values) : ser(k, values))
    const group = family === 'scatter' ? `<c:scatterChart><c:scatterStyle val="lineMarker"/>${series.join('')}<c:axId val="10"/><c:axId val="20"/></c:scatterChart>` + valAx('10', '20', '<c:axPos val="b"/>') + valAx('20', '10', '<c:axPos val="l"/>')
      : `<c:${family}Chart><c:grouping val="${grouping}"/>${series.join('')}<c:axId val="1"/><c:axId val="2"/></c:${family}Chart>` + catAx('1', '2') + valAx('2', '1')
    return parse(group)
  }

  it('R1: 64 series x 4,096 points: the measured construction stays within the admitted bound, which does not scale with series', () => {
    const models = [many('area', 'stacked'), many('line'), many('scatter')]
    for (const model of models) {
      const bound = axisGraphBound(model, 4096)
      const bytes = measured(model, 4096, () => { const r = recorder(400, 300); paintXlsxChart(model, r.ctx, 400, 300) })
      // Stacked areas keep the stack sums and the three streamed buffers: per-slot arrays exist, and stay bounded.
      if (model === models[0]) expect(bytes).toBeGreaterThanOrEqual(6 * 4096 * 8)
      expect(bytes).toBeLessThanOrEqual(bound)
    }
    {
      const model = models[0], bound = axisGraphBound(model, 4096)
      // A parent with exactly the bound admits the paint; one byte less refuses it with the card.
      for (const [room, native] of [[bound, true], [bound - 1, false]] as const) {
        let held = 0
        const workspace = createChartWorkspace('desktop', { reserve: b => { if (held + b > room) return false; held += b; return true }, release: (_t, b) => { held -= b } })
        const context = documentChartContext('desktop', undefined, undefined)
        const r = recorder(400, 300)
        paintXlsxChart(model, r.ctx, 400, 300, undefined, { ...context, workspace })
        expect(r.texts().join(' ').includes('memory')).toBe(!native)
        expect(held).toBe(0)
      }
    }
  }, 120_000)

  it('R2: a transferred model is charged to the receiving document chart facet and its profile; with no headroom it is a card', async () => {
    const plot = bar(ser(0, [3, 7, 2])) + catAx('1', '2') + valAx('2', '1')
    const owned = parseXlsxChart(parseXmlOrdered(space(plot)), undefined, { workspace: createChartWorkspace('desktop'), profile: 'desktop', partPath: 'c.xml' } as unknown as ChartParseContext)!
    const clone = structuredClone(owned)
    expect(chartModelOwner(clone)).toBeUndefined()
    // The receiving side (an M-profile viewer) associates the transferred models with its document context.
    const receiving = { sheets: [{ drawings: [{ content: { kind: 'xlsxChart', chart: clone } }] }] }
    await ensureChartPainter(receiving.sheets[0].drawings as never, { owner: receiving, profile: 'mobile' })
    const context = receivedChartContext(clone)!
    expect(context.profile).toBe('mobile')
    expect(context.workspace.limit).toBe(32 * 1024 * 1024)
    // The same document reuses its context; a paint charges it and refunds after.
    const log: number[] = []
    const reserve = context.workspace.reserve.bind(context.workspace)
    context.workspace.reserve = (b, t) => { log.push(b); return reserve(b, t) }
    const ok = recorder()
    paintXlsxChart(clone, ok.ctx, 800, 500)
    expect(log.length).toBe(1)
    expect(context.workspace.used).toBe(0)
    expect(ok.texts()).toContain('C1')
    // No headroom in the receiving document: the card.
    context.workspace.reserve = reserve
    context.workspace.reserve(context.workspace.limit, 'pressure')
    const full = recorder()
    paintXlsxChart(clone, full.ctx, 800, 500)
    expect(full.texts().join(' ')).toMatch(/memory/)
    context.workspace.release('pressure')
    // getPaintables-style callers pass the viewer profile; the parent ledger is honoured.
    let parentHeld = 0
    const parentContext = documentChartContext('desktop', undefined, { reserve: b => { if (parentHeld + b > 10) return false; parentHeld += b; return true }, release: (_t, b) => { parentHeld -= b } })
    const other = structuredClone(owned)
    await ensureChartEngine([{ content: { kind: 'xlsxChart', chart: other } as never }], { owner: {}, context: parentContext })
    const refused = recorder()
    paintXlsxChart(other, refused.ctx, 800, 500)
    expect(refused.texts().join(' ')).toMatch(/memory/)
    expect(parentHeld).toBe(0)
    // Synchronous paints with nothing received: an explicit context wins; else a bounded per-call workspace.
    const loose = structuredClone(owned)
    const explicit = recorder()
    paintXlsxChart(loose, explicit.ctx, 800, 500, undefined, parentContext)
    expect(explicit.texts().join(' ')).toMatch(/memory/)
    const standalone = recorder()
    paintXlsxChart(loose, standalone.ctx, 800, 500)
    expect(standalone.texts()).toContain('C1')
    expect(CHART_STANDALONE_PAINT_PROFILE).toBe('mobile')
  })

  it('R3: one-sided extreme auto scales fall back to finite bounds with the rendering-bound diagnostic', () => {
    for (const v of [1.7e308, -1.7e308]) {
      const graph = resolveAxes(parse(line(ser(0, [v])) + catAx('1', '2') + valAx('2', '1')))
      const value = graph.axes.get('2')!
      expect([value.min, value.max, value.major].every(Number.isFinite)).toBe(true)
      expect(value.max).toBeGreaterThan(value.min)
      expect(value.ticks.length).toBeGreaterThan(1)
      expect(value.ticks.every(Number.isFinite)).toBe(true)
      expect(Number.isFinite(axisFraction(value, v))).toBe(true)
      expect(axisFraction(value, v)).toBe(v > 0 ? 1 : 0)
      expect(graph.diagnostics.map(d => d.code)).toContain('chart-tick-limit')
      const r = recorder()
      paintXlsxChart(parse(line(ser(0, [v, v / 2])) + catAx('1', '2') + valAx('2', '1')), r.ctx, 800, 500)
      const coords = [...r.of('moveTo'), ...r.of('lineTo'), ...r.of('arc')].flat().filter(c => typeof c === 'number')
      expect(coords.every(Number.isFinite)).toBe(true)
    }
  })
})

describe('A1a fix round 3: the interactive sheet preparation receives charts into the viewer budget', () => {
  const PLOT = bar(ser(0, [3, 7, 2])) + catAx('1', '2') + valAx('2', '1')
  /** A resident workbook whose chart model is replaced by a structured clone (an unowned, transferred model). */
  async function transferredDoc() {
    const doc = await parseXlsx(await chartWorkbook(space(PLOT)))
    const node = doc.sheets[0].drawings![0]
    if (node.content?.kind !== 'xlsxChart') throw new Error('no chart')
    const clone = structuredClone(node.content.chart)
    node.content = { ...node.content, chart: clone }
    expect(chartModelOwner(clone)).toBeUndefined()
    return { doc, clone }
  }
  const refusingParent = () => {
    const calls: number[] = []
    return { calls, parent: { reserve: (bytes: number) => { calls.push(bytes); return false }, release: () => {} } }
  }

  it('R2: prepareSheetAssets keeps a supplied chart context (its parent is called; no headroom paints the card)', async () => {
    const { doc, clone } = await transferredDoc()
    const { calls, parent } = refusingParent()
    const supplied = documentChartContext('mobile', undefined, parent)
    const lease = await prepareSheetAssets(doc, { chartContext: supplied, profile: 'M' })
    try {
      expect(receivedChartContext(clone)).toBe(supplied)
      const r = recorder()
      paintXlsxChart(clone, r.ctx, 800, 500)
      expect(calls.length).toBe(1)
      expect(r.texts().join(' ')).toMatch(/memory/)
    } finally { lease.dispose() }
  })

  it('R2: without a supplied context, the receiving facet is profile-sized and parented to the viewer ledger (chartParent)', async () => {
    const { doc, clone } = await transferredDoc()
    const ledger = new MainMemoryLedger()
    const parent = mainChartParent(ledger, MAIN_THREAD_BUDGETS.M)
    const lease = await prepareSheetAssets(doc, { chartParent: parent, profile: 'M' })
    try {
      const context = receivedChartContext(clone)!
      expect(context.profile).toBe('mobile')
      expect(context.workspace.limit).toBe(32 * 1024 * 1024)
      // With room: the paint's charge reaches the viewer ledger, and returns after the paint.
      let peak = 0
      const set = ledger.set.bind(ledger)
      ledger.set = (pool, owner, bytes) => { if (pool === 'drawings' && owner === 'charts') peak = Math.max(peak, bytes); set(pool, owner, bytes) }
      const ok = recorder()
      paintXlsxChart(clone, ok.ctx, 800, 500)
      expect(ok.texts()).toContain('C1')
      expect(peak).toBeGreaterThan(0)
      expect(ledger.ownerBytes('drawings', 'charts')).toBe(0)
      // The viewer ledger full: the card.
      ledger.set('layout', 'pressure', MAIN_THREAD_BUDGETS.M)
      const full = recorder()
      paintXlsxChart(clone, full.ctx, 800, 500)
      expect(full.texts().join(' ')).toMatch(/memory/)
      expect(ledger.ownerBytes('drawings', 'charts')).toBe(0)
    } finally { lease.dispose() }
  })
})
