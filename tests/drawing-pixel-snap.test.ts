import { expect, test } from 'vitest'
import { createCanvas } from 'canvas'
import type { SceneNode } from '../src/drawing/scene'
import type { GeometryCommand } from '../src/drawing/geometry'
import type { DrawingLine } from '../src/drawing/style'
import { paintScene } from '../src/drawing/scene-paint'

// PowerPoint's device-pixel rule, measured from its movie export (1 px = 6350 EMU probe deck):
// fill edges land on round-half-up pixels with no anti-aliasing; a line of W = max(1, round(width)) px
// covers pixels [s - floor(W/2), s - floor(W/2) + W) around each snapped edge s.
const emu = (px: number) => px * 9525
const rect: GeometryCommand[] = [['moveTo', '0', '0'], ['lnTo', 'w', '0'], ['lnTo', 'w', 'h'], ['lnTo', '0', 'h'], ['close']]
const node = (style: { fill?: boolean; line?: DrawingLine }): SceneNode => ({
  xEmu: emu(10), yEmu: emu(10), widthEmu: emu(40), heightEmu: emu(30), geometry: 'rect',
  drawingGeometry: { preset: 'rect', adjustments: [], guides: [], paths: [{ commands: rect }] },
  drawingStyle: {
    ...(style.fill ? { fill: { kind: 'solid' as const, color: { r: 128, g: 128, b: 128, a: 1 } } } : {}),
    ...(style.line ? { line: { fill: { kind: 'solid' as const, color: { r: 0, g: 0, b: 0, a: 1 } }, ...style.line } } : {}),
    issues: [],
  },
})
// The caller places the slide half a pixel right and a quarter down: device box x 10.5..50.5, y 10.25..40.25.
function paint(n: SceneNode, pixelSnap: boolean, sx = 1) {
  const ctx = createCanvas(160, 60).getContext('2d') as unknown as CanvasRenderingContext2D
  ctx.translate(0.5, 0.25)
  ctx.scale(sx, 1)
  paintScene([n], ctx, { pixelSnap })
  return {
    alphaAt: (x: number, y: number) => ctx.getImageData(x, y, 1, 1).data[3],
    red: (x: number, y: number) => { const d = ctx.getImageData(x, y, 1, 1).data; return d[3] === 255 ? d[0] : d[3] === 0 ? undefined : -1 },
    row: (y: number, xs: number[]) => xs.map(x => ctx.getImageData(x, y, 1, 1).data[3]),
    col: (x: number, ys: number[]) => ys.map(y => ctx.getImageData(x, y, 1, 1).data[3]),
  }
}

test('fills snap their edges to round-half-up device pixels, without anti-aliasing', () => {
  const p = paint(node({ fill: true }), true)
  expect(p.row(25, [10, 11, 50, 51])).toEqual([0, 255, 255, 0])
  expect(p.col(25, [9, 10, 39, 40])).toEqual([0, 255, 255, 0])
})

test('without pixelSnap the fill keeps its anti-aliased edge', () => {
  const p = paint(node({ fill: true }), false)
  expect(p.alphaAt(10, 25)).toBeGreaterThan(64)
  expect(p.alphaAt(10, 25)).toBeLessThan(192)
})

test('a 1 px line covers the snapped edge pixel: inside the fill at left/top, outside it at right/bottom', () => {
  const p = paint(node({ line: { width: 1 } }), true)
  expect(p.row(25, [10, 11, 12, 49, 50, 51, 52])).toEqual([0, 255, 0, 0, 0, 255, 0])
  expect(p.col(25, [9, 10, 11, 39, 40, 41])).toEqual([0, 255, 0, 0, 255, 0])
})

test('line widths round to whole pixels (at least one), even widths straddle the snapped edge', () => {
  for (const width of [0.5, 1]) expect(paint(node({ line: { width } }), true).row(25, [10, 11, 12])).toEqual([0, 255, 0])
  for (const width of [1.5, 2]) expect(paint(node({ line: { width } }), true).row(25, [9, 10, 11, 12, 50, 51, 52])).toEqual([0, 255, 255, 0, 255, 255, 0])
  expect(paint(node({ line: { width: 3 } }), true).row(25, [9, 10, 11, 12, 13])).toEqual([0, 255, 255, 255, 0])
})

test('fill plus line: the fill stops where the right/bottom line begins', () => {
  const p = paint(node({ fill: true, line: { width: 1 } }), true)
  expect([10, 11, 12, 50, 51, 52].map(x => p.red(x, 25))).toEqual([undefined, 0, 128, 128, 0, undefined])
})

test('under unequal axis scales each axis keeps its own whole-pixel width', () => {
  // x scale 2: device box x 20.5..100.5 (snapped 21..101), y 10.25..40.25 (10..40); a 1 px line is 2 px wide
  // on the vertical edges and 1 px on the horizontal ones, as it was without snapping, but on whole pixels.
  const p = paint(node({ line: { width: 1 } }), true, 2)
  expect(p.row(25, [19, 20, 21, 22, 99, 100, 101, 102])).toEqual([0, 255, 255, 0, 0, 255, 255, 0])
  expect(p.col(60, [9, 10, 11, 39, 40, 41])).toEqual([0, 255, 0, 0, 255, 0])
})

test('rotated shapes are not snapped', () => {
  const p = paint({ ...node({ fill: true }), rotationDeg: 30 }, true)
  const edge = p.row(12, Array.from({ length: 60 }, (_, i) => i)).filter(a => a > 0 && a < 255)
  expect(edge.length).toBeGreaterThan(0)
})
