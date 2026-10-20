// @vitest-environment node
import { describe, it, expect } from 'vitest'
import { parseXmlOrdered } from '../src/core/xml'
import { parseXlsxChart } from '../src/charts/parse'
import type { ThemeContext } from '../src/drawing/style'

const tx = (size: number, body = '') => `<txPr><bodyPr ${body}/><p><pPr><defRPr sz="${size}"/></pPr></p></txPr>`
export const chartXml = (extra = '', axisExtra = '', rootExtra = '') => `<chartSpace>${rootExtra}<chart><autoTitleDeleted val="1"/><plotArea>${extra}<barChart><ser><cat><strLit><pt idx="0"><v>Alpha long label</v></pt><pt idx="1"><v>Beta long label</v></pt></strLit></cat><val><numLit><pt idx="0"><v>3</v></pt><pt idx="1"><v>7</v></pt></numLit></val></ser><axId val="1"/><axId val="2"/></barChart><catAx><axId val="1"/>${axisExtra}<crossAx val="2"/></catAx><valAx><axId val="2"/><crossAx val="1"/></valAx></plotArea></chart></chartSpace>`
const parse = (xml: string, theme?: ThemeContext) => parseXlsxChart(parseXmlOrdered(xml), theme)!

describe('chart placement input', () => {
  it('uses the theme minor font and clamps txPr to 6–72 pt', () => {
    const theme = { palette: {}, fonts: { minor: { latin: 'Test Minor' } } } as ThemeContext
    expect(parse(chartXml('', '', tx(9000)), theme).fontFamily).toBe('Test Minor')
    expect(parse(chartXml('', '', tx(9000)), theme).fontSizePt).toBe(72)
    expect(parse(chartXml('', '', tx(200))).fontSizePt).toBe(6)
  })
  it('preserves axis rotation, vertical text and per-axis size', () => {
    const model = parse(chartXml('', tx(1800, 'rot="-2700000" vert="vert270"')))
    expect(model.axes[0]).toMatchObject({ text: { sizePt: 18, rotation: -45, vertical: 'vert270' } })
  })
  it('preserves manual modes and leaves omitted values automatic', () => {
    const model = parse(chartXml('<layout><manualLayout><x val="0.1"/><yMode val="edge"/><y val="0.2"/><wMode val="edge"/><w val="0.8"/></manualLayout></layout>'))
    expect(model).toMatchObject({ plotLayout: { x: .1, y: .2, w: .8, yMode: 'edge', wMode: 'edge' } })
  })
})

import * as placement from '../src/charts/layout'
import { documentChartContext } from '../src/charts/context'
const measure = { width: (s: string) => s.length * 7, lineHeight: () => 14 }
describe('charged chart placement', () => {
  it('caches by size, dpr and font generation, keeps charges until invalidation/dispose', () => {
    const model = parse(chartXml()), context = documentChartContext('mobile')
    let calls = 0
    const metrics = { ...measure, width: (s: string) => { calls++; return measure.width(s) } }
    const first = placement.layoutChart(model, 400, 260, 1, 0, metrics, context)
    const initialCalls = calls, held = context.workspace.used
    expect(held).toBeGreaterThan(0)
    expect(placement.layoutChart(model, 400, 260, 1, 0, metrics, context)).toBe(first)
    expect(calls).toBe(initialCalls)
    let expected = held
    for (const [w, dpr, gen] of [[500, 1, 0], [500, 2, 0], [500, 2, 1]]) {
      const next = placement.layoutChart(model, w, 260, dpr, gen, metrics, context)
      expect(next).not.toBe(first)
      expected += next.retainedBytes
      expect(context.workspace.used).toBe(expected)
    }
    placement.releaseChartLayouts(model)
    expect(context.workspace.used).toBe(0)
  })
  it('refuses before measuring when the live parent has no headroom', () => {
    const context = documentChartContext('mobile', undefined, { reserve: () => false, release: () => {} })
    let calls = 0
    expect(() => placement.layoutChart(parse(chartXml()), 400, 260, 1, 0, { ...measure, width: () => { calls++; return 10 } }, context)).toThrow()
    expect(calls).toBe(0)
    expect(context.workspace.used).toBe(0)
  })
  it('resolves manual edge/factor modes in order and falls back on inverted boxes', () => {
    const auto = { x: 20, y: 30, w: 200, h: 100 }, diagnostics: any[] = []
    expect(placement.resolveManualLayout({ x: .1, y: .2, w: .8, h: .7, wMode: 'edge', hMode: 'edge' }, auto, 400, 300, diagnostics)).toEqual({ x: 60, y: 90, w: 260, h: 120 })
    expect(placement.resolveManualLayout({ x: .1, y: .2, xMode: 'edge', yMode: 'edge' }, auto, 400, 300, diagnostics)).toEqual({ x: 40, y: 60, w: 200, h: 100 })
    expect(placement.resolveManualLayout({ x: .9, w: .2, xMode: 'edge', wMode: 'edge' }, auto, 400, 300, diagnostics)).toEqual(auto)
    expect(diagnostics).toContainEqual(expect.objectContaining({ code: 'chart-layout-invalid' }))
  })
  it('thins rotated labels in screen order without overlapping extents', () => {
    const model = parse(chartXml('', tx(1800, 'rot="-2700000"'))), context = documentChartContext('mobile')
    const layout = placement.layoutChart(model, 130, 200, 1, 0, measure, context)
    const labels = layout.labels.get('1')!
    expect(labels.length).toBeGreaterThan(0)
    expect(labels[0].angle).toBe(-Math.PI / 4)
    for (let i = 1; i < labels.length; i++) expect(labels[i].position - labels[i].extent / 2).toBeGreaterThanOrEqual(labels[i-1].position + labels[i-1].extent / 2 + 4)
    placement.releaseChartLayouts(model)
  })
})

import { paintXlsxChart, acquireChartContext } from '../src/charts/paint'
import { createCanvas } from 'canvas'
describe('layout painter lifetime', () => {
  it('retains layouts but refunds paint transients, avoids warm measurements, and disposes with receiver', () => {
    const model = structuredClone(parse(chartXml())), receiver = acquireChartContext({ owner: {}, profile: 'mobile' })
    const ctx = createCanvas(400, 260).getContext('2d')
    let calls = 0; const original = ctx.measureText.bind(ctx)
    ctx.measureText = s => { calls++; return original(s) }
    paintXlsxChart(model, ctx as any, 400, 260, undefined, receiver.context)
    const held = receiver.context.workspace.used, measured = calls
    expect(held).toBeGreaterThan(0)
    paintXlsxChart(model, ctx as any, 400, 260, undefined, receiver.context)
    expect(calls).toBe(measured)
    expect(receiver.context.workspace.used).toBe(held)
    receiver.release()
    expect(receiver.context.workspace.used).toBe(0)
  })
})

import { fillChartRanges } from '../src/charts/fill'
describe('size-independent preparation', () => {
  it('transfers domains, candidate ticks and formatted labels; placement does not rescan values', () => {
    const model = structuredClone(parse(chartXml()))
    expect(model).toHaveProperty('preparedAxes')
    const numeric = model.groups[0].series[0].values
    Object.defineProperty(numeric, 'values', { get: () => { throw new Error('Stage 2 rescanned values') } })
    const context = documentChartContext('mobile')
    const layout = placement.layoutChart(model, 400, 260, 1, 0, measure, context)
    expect(layout.graph.axes.get('2')!.max).toBeGreaterThanOrEqual(7)
    expect(layout.labels.get('2')!.some(l => l.text === '0')).toBe(true)
    placement.releaseChartLayouts(model)
  })
})

import { releaseChartModel, chartModelOwner } from '../src/charts/own'
it('model disposal releases its placement sidecar as well as its prepared data', () => {
  const context = documentChartContext('mobile')
  const model = parseXlsxChart(parseXmlOrdered(chartXml()), undefined, { ...context, partPath: 'c.xml' } as any)!, owner = chartModelOwner(model)!
  placement.layoutChart(model, 400, 260, 1, 0, measure)
  expect(owner.workspace.used).toBeGreaterThan(0)
  releaseChartModel(model)
  expect(owner.workspace.used).toBe(0)
})
it('uses the resolved legend box size and inner plot box literally', () => {
  const xml = chartXml('<layout><manualLayout><layoutTarget val="inner"/><xMode val="edge"/><yMode val="edge"/><x val="0.2"/><y val="0.2"/><w val="0.5"/><h val="0.5"/></manualLayout></layout>').replace('</chart>', '<legend><legendPos val="r"/><layout><manualLayout><xMode val="edge"/><yMode val="edge"/><x val="0.1"/><y val="0.1"/><w val="0.2"/><h val="0.4"/></manualLayout></layout></legend></chart>')
  const model = parse(xml), layout = placement.layoutChart(model, 400, 300, 1, 0, measure)
  expect(layout.plot).toEqual({ x: 80, y: 60, w: 200, h: 150 })
  expect(layout.legend).toEqual({ x: 40, y: 30, w: 80, h: 120 })
  const ctx = createCanvas(400, 300).getContext('2d'), swatches: number[][] = []
  const fill = ctx.fillRect.bind(ctx)
  ctx.fillRect = (x, y, w, h) => { swatches.push([x,y,w,h]); fill(x,y,w,h) }
  paintXlsxChart(model, ctx as any, 400, 300)
  expect(swatches).toContainEqual([46, 30 + (model.fontSizePt * 96 / 72 + 6 - model.fontSizePt * 96 / 72 * .7) / 2, model.fontSizePt * 96 / 72 * .7, model.fontSizePt * 96 / 72 * .7])
  placement.releaseChartLayouts(model)
})
it('explicit receiving contexts retain layouts until the last preparation releases them', () => {
  const model = structuredClone(parse(chartXml())), context = documentChartContext('mobile')
  const a = acquireChartContext({ owner: {}, context }), b = acquireChartContext({ owner: {}, context })
  placement.layoutChart(model, 400, 260, 1, 0, measure, context)
  a.release(); expect(context.workspace.used).toBeGreaterThan(0)
  b.release(); expect(context.workspace.used).toBe(0)
})
it('honours tickMarkSkip even when category labels are hidden', () => {
  const paint = (skip: number) => {
    const model = parse(chartXml('', `<tickLblPos val="none"/><majorTickMark val="out"/><tickMarkSkip val="${skip}"/>`))
    const ctx = createCanvas(400, 260).getContext('2d'), marks: number[][] = []
    const line = ctx.lineTo.bind(ctx); ctx.lineTo = (x, y) => { marks.push([x,y]); line(x,y) }
    paintXlsxChart(model, ctx as any, 400, 260)
    return marks.length
  }
  expect(paint(1) - paint(2)).toBe(1)
})
import { PointWriter } from '../src/charts/points'
it('drops later area series at the path work bound with a diagnostic', () => {
  const model = parse(chartXml()), points = new PointWriter(4096)
  for (let i = 0; i < 4096; i++) points.value(i, i % 7)
  delete model.preparedAxes
  model.groups[0].family = 'area'
  model.groups[0].series = Array.from({ length: 64 }, () => ({ name: 's', color: '112233', values: points.done() }))
  model.categories = []
  const context = documentChartContext('desktop'); let segments = 0
  const ctx = new Proxy({ font: '', measureText: () => ({ width: 10 }) }, { get: (t: any, k: any) => k in t ? t[k] : (..._args: any[]) => { if (k === 'lineTo' || k === 'moveTo' || k === 'arc') segments++ }, set: (t:any,k:any,v:any) => {t[k]=v;return true} })
  paintXlsxChart(model, ctx as any, 400, 260, undefined, context)
  const layout = placement.layoutChart(model, 400, 260, 1, 0, measure, context)
  expect(segments).toBeGreaterThan(4096)
  expect(segments).toBeLessThanOrEqual(262144)
  expect(layout.diagnostics).toContainEqual(expect.objectContaining({ code: 'chart-path-limit' }))
  placement.releaseChartLayouts(model)
})
it('keeps an invalidated layout charged until its active paint returns', () => {
  const model = structuredClone(parse(chartXml())), receiver = acquireChartContext({ owner: {}, profile: 'mobile' })
  const ctx = createCanvas(400, 260).getContext('2d'), original = ctx.fillRect.bind(ctx)
  let once = false, during = 0, expected = 0
  ctx.fillRect = (x,y,w,h) => { if (!once) { once = true; const layout = placement.layoutChart(model, 400, 260, 1, 0, measure, receiver.context); expected = layout.retainedBytes + placement.paintWorkspaceBytes(model, layout.view); receiver.release(); during = receiver.context.workspace.used }; original(x,y,w,h) }
  paintXlsxChart(model, ctx as any, 400, 260, undefined, receiver.context)
  expect(during).toBe(expected)
  expect(receiver.context.workspace.used).toBe(0)
})
it('invalidates charged widths on font publication and resolves the theme font through the font manager', () => {
  const model = structuredClone(parse(chartXml())), receiver = acquireChartContext({ owner: {}, profile: 'mobile' })
  const ctx = createCanvas(400, 260).getContext('2d'), fonts = new EventTarget()
  Object.defineProperty(ctx.canvas, 'ownerDocument', { value: { fonts } })
  const seen: string[] = []; const original = ctx.measureText.bind(ctx)
  ctx.measureText = text => { seen.push(ctx.font); return original(text) }
  paintXlsxChart(model, ctx as any, 400, 260, () => '"Registered Minor"', receiver.context)
  expect(seen.some(font => font.includes('Registered Minor'))).toBe(true)
  const count = seen.length
  fonts.dispatchEvent(new Event('loadingdone'))
  expect(receiver.context.workspace.used).toBe(0)
  paintXlsxChart(model, ctx as any, 400, 260, () => '"Registered Minor"', receiver.context)
  expect(seen.length).toBeGreaterThan(count)
  receiver.release()
})
it('replaces stage-1 ownership on a refill instead of accumulating obsolete snapshots', () => {
  const context = documentChartContext('mobile')
  const first = parseXlsxChart(parseXmlOrdered(chartXml()), undefined, { ...context, partPath: 'c.xml' } as any)!
  const before = context.workspace.used
  const second = fillChartRanges(first, 'S', () => undefined)
  expect(second.preparedAxes).not.toBe(first.preparedAxes)
  expect(context.workspace.used).toBe(before)
  releaseChartModel(second)
  expect(context.workspace.used).toBe(0)
})
it('keeps authored chart text within 72 pt when painting a title', () => {
  const model = parse(chartXml('', '', tx(7200)).replace('<autoTitleDeleted val="1"/>', '<title><tx><rich><p><r><t>Title</t></r></p></rich></tx></title>'))
  const ctx = createCanvas(400, 260).getContext('2d'), sizes: number[] = [], fill = ctx.fillText.bind(ctx)
  ctx.fillText = ((text: string, x: number, y: number, width?: number) => { if (text === 'Title') sizes.push(parseFloat(ctx.font)); if(width === undefined)fill(text,x,y);else fill(text,x,y,width) }) as typeof ctx.fillText
  paintXlsxChart(model, ctx as any, 400, 260)
  expect(sizes).toEqual([96])
})
it('keeps existing value labels when their boxes fit without overlap', () => {
  const model = parse(chartXml()), context = documentChartContext('mobile')
  const layout = placement.layoutChart(model, 100, 100, 1, 0, measure, context)
  expect(layout.labels.get('2')).toHaveLength(layout.graph.axes.get('2')!.ticks.length)
  placement.releaseChartLayouts(model)
})
it('reuses prepared value-label strings during warm paint', () => {
  const model = parse(chartXml()), context = documentChartContext('mobile')
  const ctx = createCanvas(400, 260).getContext('2d')
  paintXlsxChart(model, ctx as any, 400, 260, undefined, context)
  const layout = placement.layoutChart(model, 400, 260, 1, 0, measure, context)
  for (const axis of layout.graph.axes.values()) axis.labelValue = () => { throw new Error('Reformatted during paint') }
  expect(() => paintXlsxChart(model, ctx as any, 400, 260, undefined, context)).not.toThrow()
  placement.releaseChartLayouts(model)
})
import { stampWorkbookEpoch } from '../src/charts/fill'
it('contains an epoch preparation refusal to its chart', () => {
  const context = documentChartContext('mobile')
  const model = parseXlsxChart(parseXmlOrdered(chartXml()), undefined, { ...context, partPath: 'c.xml' } as any)!
  context.workspace.reserve(context.workspace.limit - context.workspace.used, 'pressure')
  const drawings = [{ content: { kind: 'xlsxChart', chart: model } }]
  expect(() => stampWorkbookEpoch(drawings as any, '1904')).not.toThrow()
  expect(drawings[0].content.chart.refused).toBe(true)
  context.workspace.release('pressure')
  releaseChartModel(drawings[0].content.chart)
  expect(context.workspace.used).toBe(0)
})

describe('A1b review: retained placement cache', () => {
  it('paints 200 small charts natively in one mobile document and retains a tight footprint', () => {
    const context = documentChartContext('mobile'), seed = structuredClone(parse(chartXml()))
    const ctx = createCanvas(400, 260).getContext('2d'), texts: string[] = []
    const draw = ctx.fillText.bind(ctx)
    ctx.fillText = ((text: string, ...args: any[]) => { texts.push(text); (draw as any)(text, ...args) }) as typeof ctx.fillText
    for (let i = 0; i < 200; i++) {
      texts.length = 0
      paintXlsxChart(structuredClone(seed), ctx as any, 400, 260, undefined, context)
      expect(texts).toContain('Alpha long label')
      expect(texts.join(' ')).not.toContain('memory')
    }
    expect(context.workspace.used).toBeLessThan(200 * 32_768)
    placement.releaseWorkspaceLayouts(context.workspace)
    expect(context.workspace.used).toBe(0)
  })
  it('revisits A → B → A without measurement or a new charge', () => {
    const model = parse(chartXml()), context = documentChartContext('mobile')
    let calls = 0; const metrics = { ...measure, width: (s: string) => { calls++; return s.length * 7 } }
    const a = placement.layoutChart(model, 400, 260, 1, 0, metrics, context)
    placement.layoutChart(model, 600, 260, 1, 0, metrics, context)
    const held = context.workspace.used, count = calls
    expect(placement.layoutChart(model, 400, 260, 1, 0, metrics, context)).toBe(a)
    expect([context.workspace.used, calls]).toEqual([held, count])
    placement.releaseChartLayouts(model)
  })
  it('does not call or retain the measuring callback after construction', () => {
    const model = parse(chartXml()), context = documentChartContext('mobile')
    let live = true
    const layout = placement.layoutChart(model, 400, 260, 1, 0, { ...measure, width: s => { if (!live) throw Error('retained paint context'); return s.length * 7 } }, context)
    live = false
    expect(() => layout.width('new unmeasured text', '99px Test')).toThrow(/not measured/)
    placement.releaseChartLayouts(model)
  })
  it('uses a constant automatic label interval for varying widths', () => {
    const model = parse(chartXml()), context = documentChartContext('mobile')
    model.categories = ['L'.repeat(18), 'b', 'c', 'L'.repeat(18), 'e', 'f', 'g', 'h']; delete model.preparedAxes
    const layout = placement.layoutChart(model, 360, 260, 1, 0, measure, context)
    const indices = layout.labels.get('1')!.map(x => x.index).sort((a,b) => a-b)
    expect(indices.length).toBeGreaterThan(2)
    expect(new Set(indices.slice(1).map((v,i) => v-indices[i])).size).toBe(1)
    placement.releaseChartLayouts(model)
  })
  it.each([['vert', 90], ['eaVert', 90], ['vert270', -90]] as const)('places %s labels at the authored quarter turn', (vert, degrees) => {
    const model = parse(chartXml('', tx(1800, `vert="${vert}"`)))
    const layout = placement.layoutChart(model, 500, 350, 1, 0, measure)
    expect(layout.labels.get('1')!.every(x => x.angle === degrees*Math.PI/180)).toBe(true)
    expect(layout.labels.get('1')!.length).toBeGreaterThan(0)
    placement.releaseChartLayouts(model)
  })
  it('respects authored tickLblSkip and guards the remaining label extents', () => {
    const model = parse(chartXml('', '<tickLblSkip val="2"/>')), context = documentChartContext('mobile')
    model.categories = Array.from({length: 10}, (_,i) => `Label ${i}`); delete model.preparedAxes
    const layout = placement.layoutChart(model, 400, 260, 1, 0, measure, context), labels = layout.labels.get('1')!
    expect(labels.length).toBeGreaterThan(1)
    expect(labels.every(x => x.index % 2 === 0)).toBe(true)
    for(let i=1;i<labels.length;i++) expect(labels[i].position-labels[i].extent/2).toBeGreaterThanOrEqual(labels[i-1].position+labels[i-1].extent/2+4)
    placement.releaseChartLayouts(model)
  })
  it('uses the reserved rotated axis band for the display-unit label', () => {
    const model = parse(chartXml().replace('<valAx><axId val="2"/>', `<valAx><axId val="2"/>${tx(3000,'rot="5400000"')}<dispUnits><builtInUnit val="hundreds"/><dispUnitsLbl/></dispUnits>`))
    const context = documentChartContext('mobile'), ctx = createCanvas(600,400).getContext('2d')
    const moves: number[][] = [], translate = ctx.translate.bind(ctx)
    ctx.translate = (x,y) => { moves.push([x,y]); translate(x,y) }
    paintXlsxChart(model, ctx as any, 600,400,undefined,context)
    const layout = placement.layoutChart(model,600,400,1,0,measure,context)
    // 30 pt = 40px; rotated text's cross extent is 40px, plus six px gap.
    expect(moves.some(([x,y]) => Math.abs(x-(layout.plot.x-46))<1e-8 && Math.abs(y-(layout.plot.y+layout.plot.h/2))<1e-8)).toBe(true)
    placement.releaseChartLayouts(model)
  })
  it('does not compress a horizontal category label already admitted by layout', () => {
    const model = parse(chartXml('', '<tickLblSkip val="1"/>')), context = documentChartContext('mobile')
    const ctx = createCanvas(500,300).getContext('2d'), widths: unknown[] = [], fill = ctx.fillText.bind(ctx)
    ctx.fillText = ((text:string,x:number,y:number,max?:number) => { if(model.categories.includes(text)) widths.push(max); if(max===undefined)fill(text,x,y);else fill(text,x,y,max) }) as typeof ctx.fillText
    paintXlsxChart(model,ctx as any,500,300,undefined,context)
    expect(widths.length).toBeGreaterThan(0); expect(widths.every(x=>x===undefined)).toBe(true)
    placement.releaseChartLayouts(model)
  })
})
it('evicts the least recently used unpinned layout before refusing new work', () => {
  let held = 0, cap = Infinity
  const context = documentChartContext('mobile',undefined,{reserve: b => { if(held+b>cap)return false;held+=b;return true },release:(_t,b)=>{held-=b}})
  const a=parse(chartXml()),b=parse(chartXml()),c=parse(chartXml())
  const first=placement.layoutChart(a,400,260,1,0,measure,context)
  const second=placement.layoutChart(b,400,260,1,0,measure,context)
  placement.layoutChart(a,400,260,1,0,measure,context) // A is newer than B.
  cap=held+placement.layoutBound(c)-1
  expect(()=>placement.layoutChart(c,400,260,1,0,measure,context)).not.toThrow()
  expect(placement.layoutChart(a,400,260,1,0,measure,context)).toBe(first)
  cap=Infinity
  expect(placement.layoutChart(b,400,260,1,0,measure,context)).not.toBe(second)
  placement.releaseWorkspaceLayouts(context.workspace);expect(held).toBe(0)
})
it('refuses a single genuinely oversized chart before measurement',()=>{
  const model=parse(chartXml()); model.categories=['X'.repeat(20_000_000)]
  let measured=0;const context=documentChartContext('mobile')
  expect(()=>placement.layoutChart(model,400,260,1,0,{...measure,width:()=>{measured++;return 1}},context)).toThrow()
  expect(measured).toBe(0);expect(context.workspace.used).toBe(0)
})
import v8 from 'node:v8'
import vm from 'node:vm'
v8.setFlagsFromString('--expose_gc')
const collect = vm.runInNewContext('gc') as () => void
it('allows the first painting canvas to collect while the cached layout stays live', async()=>{
  const model=parse(chartXml()), context=documentChartContext('mobile')
  function paintOnce(){const canvas=createCanvas(400,260);paintXlsxChart(model,canvas.getContext('2d') as any,400,260,undefined,context);return new WeakRef(canvas)}
  const canvas=paintOnce()
  for(let i=0;i<20;i++){await new Promise(resolve=>setTimeout(resolve,0));collect();if(!canvas.deref())break}
  expect(canvas.deref()).toBeUndefined();expect(context.workspace.used).toBeGreaterThan(0)
  placement.releaseChartLayouts(model)
})
it('drops the stale stage-1 snapshot when an epoch preparation is refused',()=>{
  const context=documentChartContext('mobile'),model=parseXlsxChart(parseXmlOrdered(chartXml()),undefined,{...context,partPath:'x'} as any)!
  context.workspace.reserve(context.workspace.limit-context.workspace.used,'pressure')
  const before=context.workspace.used,drawings=[{content:{kind:'xlsxChart',chart:model}}]
  stampWorkbookEpoch(drawings as any,'1904')
  expect(drawings[0].content.chart.preparedAxes).toBeUndefined()
  expect(context.workspace.used).toBeLessThan(before)
  context.workspace.release('pressure');releaseChartModel(drawings[0].content.chart);expect(context.workspace.used).toBe(0)
})
it('does not let page zoom invalidate layout; physical DPR remains the explicit key',()=>{
  const model=parse(chartXml()),context=documentChartContext('mobile'),ctx=createCanvas(400,260).getContext('2d')
  let calls=0;const width=ctx.measureText.bind(ctx);ctx.measureText=t=>{calls++;return width(t)}
  paintXlsxChart(model,ctx as any,400,260,undefined,context);const count=calls
  ctx.scale(2,2);paintXlsxChart(model,ctx as any,400,260,undefined,context)
  expect(calls).toBe(count);placement.releaseChartLayouts(model)
})
it('invalidates worker layouts when the realm font set publishes',()=>{
  const fonts=new EventTarget(),previous=Object.getOwnPropertyDescriptor(globalThis,'fonts')
  Object.defineProperty(globalThis,'fonts',{value:fonts,configurable:true})
  const model=parse(chartXml()),context=documentChartContext('mobile'),ctx=createCanvas(400,260).getContext('2d')
  try{paintXlsxChart(model,ctx as any,400,260,undefined,context);expect(context.workspace.used).toBeGreaterThan(0);fonts.dispatchEvent(new Event('loadingdone'));expect(context.workspace.used).toBe(0)}
  finally{if(previous)Object.defineProperty(globalThis,'fonts',previous);else delete (globalThis as any).fonts;placement.releaseChartLayouts(model)}
})
import { build } from 'esbuild'
import { createRequire } from 'node:module'
it('paints a structuredClone-transferred model synchronously in a fresh realm and workspace',async()=>{
  const result=await build({stdin:{contents:`import {paintXlsxChart,acquireChartContext} from './src/charts/paint'; export function paint(model,ctx){const receiver=acquireChartContext({owner:{},profile:'mobile'});paintXlsxChart(model,ctx,400,260,undefined,receiver.context);const held=receiver.context.workspace.used;receiver.release();return [held,receiver.context.workspace.used]}`,resolveDir:process.cwd()},bundle:true,write:false,platform:'node',format:'cjs',packages:'external'})
  const realm=vm.createContext({module:{exports:{}},require:createRequire(import.meta.url),console,TextDecoder,TextEncoder,URL,setTimeout,clearTimeout})
  vm.runInContext(result.outputFiles[0].text,realm)
  const ctx=createCanvas(400,260).getContext('2d'),text:string[]=[],fill=ctx.fillText.bind(ctx)
  ctx.fillText=((s:string,...args:any[])=>{text.push(s);(fill as any)(s,...args)}) as typeof ctx.fillText
  const [held,after]=realm.module.exports.paint(structuredClone(parse(chartXml())),ctx)
  expect(text).toContain('Alpha long label');expect(text.join(' ')).not.toContain('memory')
  expect(held).toBeGreaterThan(0);expect(after).toBe(0)
},30_000)
