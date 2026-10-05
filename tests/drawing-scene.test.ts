import { afterEach, expect, test, vi } from 'vitest'
import { createCanvas } from 'canvas'
import type { SceneNode } from '../src/drawing/scene'
import { paintScene } from '../src/drawing/scene-paint'
import { renderSlide } from '../src/pptx/render'
import type { PptxShape, PptxSlide, PptxTextBody } from '../src/pptx/types'
import { buildTextIndex, createRecordingContext, findMatches } from '../src/core/search'
import * as textMetrics from '../src/core/text-metrics'

const emu = (px: number) => px * 9525
const box = <Extra extends Partial<SceneNode> = Record<never, never>>(extra: Extra = {} as Extra): Extra & { xEmu: number; yEmu: number; widthEmu: number; heightEmu: number; geometry: 'rect' } => ({ xEmu: emu(10), yEmu: emu(10), widthEmu: emu(40), heightEmu: emu(30), geometry: 'rect', ...extra })
const surface = () => createCanvas(240, 180).getContext('2d') as unknown as CanvasRenderingContext2D
const pixel = (ctx: CanvasRenderingContext2D, x: number, y: number) => Array.from(ctx.getImageData(x, y, 1, 1).data)
const image = () => {
  const canvas = createCanvas(20, 10), ctx = canvas.getContext('2d')
  ctx.fillStyle = '#ff0000'; ctx.fillRect(0, 0, 10, 10)
  ctx.fillStyle = '#00ff00'; ctx.fillRect(10, 0, 10, 10)
  return canvas as unknown as CanvasImageSource
}

afterEach(() => vi.restoreAllMocks())

test('paints an empty scene without clearing the caller surface or painting a watermark', () => {
  const ctx = surface()
  ctx.fillStyle = '#ff0000'; ctx.fillRect(0, 0, 4, 4)
  paintScene([], ctx)
  expect(pixel(ctx, 1, 1)).toEqual([255, 0, 0, 255])
  expect(pixel(ctx, 100, 100)).toEqual([0, 0, 0, 0])
})

test('paints mixed nodes and nested children in source order, with geometry, adapter content, picture then text', () => {
  const ctx = surface(), order: string[] = []
  type Node = SceneNode<string> & { id: string; table?: boolean; children?: Node[] }
  const nodes: Node[] = [
    { ...box({ fill: '#ff0000' }), id: 'shape' },
    { ...box(), id: 'group', group: { off: { x: 0, y: 0 }, ext: { width: emu(40), height: emu(30) }, chOff: { x: emu(10), y: emu(10) }, chExt: { width: emu(40), height: emu(30) } }, children: [
      { ...box({ fill: '#0000ff', imageIndex: 0 }), id: 'mixed', table: true, textBody: 'text' },
      { ...box({ fill: '#00ffff' }), id: 'child' },
    ] },
    { ...box({ fill: '#ff00ff' }), id: 'last' },
  ]
  const draw = ctx.drawImage.bind(ctx)
  ctx.drawImage = ((...args: Parameters<typeof ctx.drawImage>) => { order.push('picture'); draw(...args) }) as typeof ctx.drawImage
  paintScene(nodes, ctx, {
    images: [image()],
    paintContent(node, context, w, h) {
      order.push(node.id)
      if (node.table) { expect(pixel(context, 20, 20)).toEqual([0, 0, 255, 255]); context.fillStyle = '#ffff00'; context.fillRect(0, 0, w, h) }
    },
    paintText(node, context, w, h) {
      expect(node.textBody).toBe('text'); order.push('text')
      expect(pixel(context, 20, 20)).toEqual([255, 0, 0, 255])
      context.fillStyle = '#ffffff'; context.fillRect(0, 0, w, h)
    },
  })
  expect(order).toEqual(['shape', 'mixed', 'picture', 'text', 'child', 'last'])
  expect(pixel(ctx, 20, 20)).toEqual([255, 0, 255, 255])
})

test('recursively composes nonuniform group scales, rotation, flips and caller affine placement', () => {
  const ctx = surface()
  ctx.setTransform(1.2, .2, .3, .8, 10, 5)
  const group = (x: number, y: number, w: number, h: number, cx: number, cy: number, cw: number, ch: number) => ({ off: { x: emu(x), y: emu(y) }, ext: { width: emu(w), height: emu(h) }, chOff: { x: emu(cx), y: emu(cy) }, chExt: { width: emu(cw), height: emu(ch) } })
  const leaf: SceneNode = box({ xEmu: emu(8), yEmu: emu(10), widthEmu: emu(12), heightEmu: emu(8), rotationDeg: 90, flipH: true, fill: '#00ff00' })
  const inner: SceneNode = box({ xEmu: emu(15), yEmu: emu(12), widthEmu: emu(40), heightEmu: emu(20), rotationDeg: 90, flipV: true, group: group(15, 12, 40, 20, 2, 4, 20, 40), children: [leaf] })
  const outer: SceneNode = box({ xEmu: emu(80), yEmu: emu(70), widthEmu: emu(100), heightEmu: emu(60), rotationDeg: 90, flipH: true, group: group(80, 70, 100, 60, 5, 7, 50, 20), children: [inner] })
  const before = ctx.getTransform(), seen: number[][] = []
  paintScene([outer], ctx, { paintContent(_node, context, w, h) {
    const m = context.getTransform(); seen.push([m.a, m.b, m.c, m.d, m.e, m.f]); expect([w, h]).toEqual([12, 8])
  } })
  // Independently expanded maps: outer=(181-3v,160-2u), inner=(23+.5v,-2+2u),
  // leaf=(18-v,20-u), hence scene=(79+6v,94+u), followed by caller affine map.
  expect(seen).toHaveLength(1)
  seen[0].forEach((value, i) => expect(value).toBeCloseTo([.3, .8, 7.2, 1.2, 133, 96][i], 10))
  expect(pixel(ctx, 163, 105)).toEqual([0, 255, 0, 255])
  expect(pixel(ctx, 115, 100)).toEqual([0, 0, 0, 0])
  expect(ctx.getTransform()).toEqual(before)
})

test('skips invalid transforms and groups, restores caller paint state and does not leak clipping', () => {
  const ctx = surface()
  ctx.translate(3, 4); ctx.globalAlpha = .5; ctx.fillStyle = '#123456'; ctx.strokeStyle = '#654321'; ctx.lineWidth = 7; ctx.setLineDash([3, 5]); ctx.font = '9px serif'
  const before = ctx.getTransform(), visited: SceneNode[] = []
  const invalid = [box({ transformValid: false }), box({ xEmu: NaN }), box({ widthEmu: -1 }), box({ rotationDeg: Infinity }), box({ group: { off: { x: 0, y: 0 }, ext: { width: 40, height: 30 }, chOff: { x: 0, y: 0 }, chExt: { width: 0, height: 30 } }, children: [box({ fill: '#ff0000' })] })]
  const valid = box({ fill: '#00ff00' })
  paintScene([...invalid, valid], ctx, { paintContent(node, context) {
    visited.push(node); context.beginPath(); context.rect(0, 0, 1, 1); context.clip(); context.fillStyle = '#ffffff'; context.font = '22px sans-serif'; context.setLineDash([])
  } })
  expect(visited).toEqual([valid]); expect(pixel(ctx, 20, 20)).toEqual([0, 255, 0, 128])
  expect(ctx.getTransform()).toEqual(before); expect(ctx.globalAlpha).toBe(.5)
  expect(ctx.fillStyle).toBe('#123456'); expect(ctx.strokeStyle).toBe('#654321'); expect(ctx.lineWidth).toBe(7); expect(ctx.getLineDash()).toEqual([3, 5]); expect(ctx.font).toBe('9px serif')
  ctx.fillRect(100, 100, 2, 2); expect(pixel(ctx, 103, 104)[3]).toBe(128)
})

test('restores nested caller state if an adapter callback throws', () => {
  const ctx = surface()
  ctx.setTransform(2, .1, .4, 3, 5, 6); ctx.fillStyle = '#123456'; ctx.globalAlpha = .75
  const before = ctx.getTransform()
  expect(() => paintScene([box()], ctx, { paintContent(_node, context) { context.scale(8, 9); context.globalAlpha = .1; context.fillStyle = '#ffffff'; throw new Error('adapter failure') } })).toThrow('adapter failure')
  expect(ctx.getTransform()).toEqual(before); expect(ctx.globalAlpha).toBe(.75); expect(ctx.fillStyle).toBe('#123456')
})

const invalidGroupExtents = [
  ...(['width', 'height'] as const).flatMap(axis => [-1, NaN, Infinity, -Infinity].map(value => ({ field: 'ext' as const, axis, value }))),
  ...(['width', 'height'] as const).flatMap(axis => [-1, 0, NaN, Infinity, -Infinity].map(value => ({ field: 'chExt' as const, axis, value }))),
]
test.each(invalidGroupExtents)('skips group $field extent ($axis): $value and restores caller state around valid neighbors', ({ field, axis, value }) => {
  const ctx = surface(), visited: SceneNode[] = []
  ctx.beginPath(); ctx.rect(0, 0, 210, 150); ctx.clip()
  ctx.translate(3, 4); ctx.globalAlpha = .5; ctx.fillStyle = '#123456'; ctx.strokeStyle = '#654321'; ctx.lineWidth = 7; ctx.setLineDash([3, 5]); ctx.font = '9px serif'
  const before = ctx.getTransform()
  const group = { off: { x: 0, y: 0 }, ext: { width: emu(40), height: emu(30) }, chOff: { x: 0, y: 0 }, chExt: { width: emu(40), height: emu(30) } }
  group[field][axis] = value
  const descendant = box({ xEmu: 0, yEmu: 0, fill: '#ff0000' })
  const invalid: SceneNode = box({ xEmu: emu(50), rotationDeg: 90, flipH: true, group, children: [descendant] })
  const first = box({ yEmu: emu(50), fill: '#0000ff' }), last = box({ xEmu: emu(100), yEmu: emu(50), fill: '#00ff00' })
  paintScene([first, invalid, last], ctx, { paintContent(node, context) {
    visited.push(node)
    context.scale(4, 5); context.globalAlpha = .1; context.fillStyle = '#ffffff'; context.strokeStyle = '#ffffff'; context.lineWidth = 1; context.font = '22px sans-serif'; context.setLineDash([])
    context.beginPath(); context.rect(0, 0, 1, 1); context.clip()
  } })
  expect(visited).toEqual([first, last])
  expect(pixel(ctx, 20, 60)).toEqual([0, 0, 255, 128]); expect(pixel(ctx, 110, 60)).toEqual([0, 255, 0, 128])
  expect(pixel(ctx, 60, 30)).toEqual([0, 0, 0, 0])
  expect(ctx.getTransform()).toEqual(before); expect(ctx.globalAlpha).toBe(.5)
  expect(ctx.fillStyle).toBe('#123456'); expect(ctx.strokeStyle).toBe('#654321'); expect(ctx.lineWidth).toBe(7); expect(ctx.getLineDash()).toEqual([3, 5]); expect(ctx.font).toBe('9px serif')
  // Adapter clipping was restored, and the caller's original clipping survived.
  ctx.fillRect(170, 120, 70, 60)
  expect(pixel(ctx, 180, 130)[3]).toBe(128); expect(pixel(ctx, 220, 160)[3]).toBe(0)
})

test('honors source crop, destination fillRect and same-asset per-use opacity under caller alpha', () => {
  const ctx = surface(), shared = image()
  ctx.globalAlpha = .5
  const first = box({ imageIndex: 0, image: { srcRect: { l: .5, t: 0, r: 0, b: 0 }, fillRect: { l: .25, t: .25, r: .25, b: .25 }, opacity: .5 } })
  const second = box({ xEmu: emu(70), imageIndex: 0, image: { srcRect: { l: 0, t: 0, r: .5, b: 0 }, opacity: 1 } })
  const draws = vi.spyOn(ctx, 'drawImage')
  paintScene([first, second], ctx, { images: [shared] })
  expect(draws.mock.calls[0]).toEqual([shared, 10, 0, 10, 10, 10, 7.5, 20, 15])
  expect(draws.mock.calls[1]).toEqual([shared, 0, 0, 10, 10, 0, 0, 40, 30])
  expect(pixel(ctx, 15, 15)).toEqual([0, 0, 0, 0])
  expect(pixel(ctx, 25, 20)).toEqual([0, 255, 0, 64])
  expect(pixel(ctx, 80, 20)).toEqual([255, 0, 0, 128])
  expect(ctx.globalAlpha).toBe(.5)
})

test('uses caller-provided stable image indices through nested nodes and reports the exact missing node', () => {
  const ctx = surface(), first = image(), second = image()
  const missing = box({ imageIndex: 9 }), child = box({ imageIndex: 1 })
  const parent = box({ group: { off: { x: 0, y: 0 }, ext: { width: 1, height: 1 }, chOff: { x: 0, y: 0 }, chExt: { width: 1, height: 1 } }, children: [child, missing] })
  const draws = vi.spyOn(ctx, 'drawImage'), missingNodes: SceneNode[] = []
  paintScene([box({ imageIndex: 0 }), parent, box({ imageIndex: 0 })], ctx, { images: [first, second], missingImage: node => missingNodes.push(node) })
  expect(draws.mock.calls.map(call => call[0])).toEqual([first, second, first]); expect(missingNodes).toEqual([missing])
  paintScene([missing], ctx, { missingImage: node => missingNodes.push(node) })
  expect(missingNodes).toHaveLength(1)
})

test('invalid picture crops, destination extents and opacity leave neighboring pictures intact', () => {
  const ctx = surface(), draws = vi.spyOn(ctx, 'drawImage')
  paintScene([
    box({ imageIndex: 0, image: { opacity: NaN } }),
    box({ imageIndex: 0, image: { opacity: -1 } }),
    box({ imageIndex: 0, image: { opacity: 2 } }),
    box({ imageIndex: 0, image: { srcRect: { l: NaN, t: 0, r: 0, b: 0 } } }),
    box({ imageIndex: 0, image: { srcRect: { l: .7, t: 0, r: .7, b: 0 } } }),
    box({ imageIndex: 0, image: { fillRect: { l: .7, t: 0, r: .7, b: 0 } } }),
    box({ imageIndex: 0, image: { fillRect: { l: Infinity, t: 0, r: 0, b: 0 } } }),
    box({ imageIndex: 0 }),
  ], ctx, { images: [image()] })
  expect(draws).toHaveBeenCalledTimes(1); expect(pixel(ctx, 20, 20)).toEqual([255, 0, 0, 255])
})

test('PPTX preserves resolver aliases, styled wrapped paragraph source offsets, scoped search and caller state', async () => {
  // Fixed advances isolate wrapping from fonts available on the test machine.
  vi.spyOn(textMetrics, 'createLocalTextMeasurer').mockImplementation(() => text => ({ width: text.length * 10, actualBoundingBoxAscent: 12, actualBoundingBoxDescent: 4 } as TextMetrics))
  const textBody: PptxTextBody = { anchor: 't', wrap: true, insetLeftEmu: 0, insetRightEmu: 0, insetTopEmu: 0, insetBottomEmu: 0, paragraphs: [{ align: 'left', level: 0, runs: [{ text: 'AB ', fontFamily: 'Embedded Face', bold: true }, { text: 'CD EF', fontFamily: 'Embedded Face', italic: true, characterSpacingPt: 3 }] }] }
  const shape: PptxShape = { ...box({ widthEmu: emu(55), heightEmu: emu(120) }), textBody, geometry: 'rect' }
  const slide: PptxSlide = { index: 0, widthEmu: emu(240), heightEmu: emu(180), shapes: [shape, { ...shape, xEmu: emu(110) }] }
  const ctx = surface(), spans: Parameters<typeof createRecordingContext>[1] = [], families: string[] = []
  ctx.setTransform(-2, .3, .7, 3, 200, 10); ctx.fillStyle = '#123456'; ctx.font = '9px serif'; ctx.globalAlpha = .75
  const before = ctx.getTransform(), resolver = (family: string) => { families.push(family); return '__officeview_embedded_face' }
  renderSlide(slide, createRecordingContext(ctx, spans), undefined, undefined, undefined, resolver)
  expect(ctx.getTransform()).toEqual(before); expect(ctx.fillStyle).toBe('#123456'); expect(ctx.font).toBe('9px serif'); expect(ctx.globalAlpha).toBe(.75)
  expect(new Set(families)).toEqual(new Set(['Embedded Face']))
  const paintedFonts: string[] = []
  vi.spyOn(ctx, 'fillText').mockImplementation(() => { paintedFonts.push(ctx.font) })
  renderSlide(slide, ctx, undefined, undefined, undefined, resolver)
  expect(paintedFonts.length).toBeGreaterThan(0)
  expect(paintedFonts.every(font => font.includes('__officeview_embedded_face'))).toBe(true)
  const sources = [...new Set(spans.map(span => span.logical!.source))]
  expect(sources).toHaveLength(2)
  for (const source of sources) {
    const recorded = spans.filter(span => span.logical!.source === source)
    expect(recorded.map(span => span.text).join('')).toBe('AB CD EF')
    expect(new Set(recorded.map(span => span.placement!.y)).size).toBeGreaterThan(1)
    let start = 0
    for (const span of recorded) { expect(span.logical).toMatchObject({ start, end: start + span.text.length }); start += span.text.length }
  }
  const index = await buildTextIndex([{ spec: { widthPx: 240, heightPx: 180 }, paint: context => renderSlide(slide, context, undefined, undefined, undefined, resolver) }])
  expect(findMatches(index, 'AB CD EF')).toHaveLength(2); expect(findMatches(index, 'EFAB')).toHaveLength(0)
})
