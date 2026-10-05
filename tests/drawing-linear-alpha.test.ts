import { describe, expect, test } from 'vitest'
import { createCanvas } from 'canvas'
import { paintGeometry } from '../src/drawing/paint'
import type { ResolvedPath } from '../src/drawing/geometry'
import type { DrawingColor, DrawingFill } from '../src/drawing/style'

const color = (r: number, g: number, b: number, a: number): DrawingColor => ({ r, g, b, a })
const rectangle: ResolvedPath = { fill: 'norm', stroke: false, commands: [['moveTo', 0, 0], ['lnTo', 256, 0], ['lnTo', 256, 64], ['lnTo', 0, 64], ['close']] }
type Gradient = Extract<DrawingFill, { kind: 'gradient' }>
const fill = (stops: Gradient['stops']): Gradient => ({ kind: 'gradient', gradient: 'linear', angle: 0, scaled: false, stops })
function surface() {
  const ctx = createCanvas(300, 100).getContext('2d') as unknown as CanvasRenderingContext2D
  ctx.fillStyle = 'black'; ctx.fillRect(0, 0, 300, 100)
  return { ctx, pixel: (x: number, y = 32) => Array.from(ctx.getImageData(x, y, 1, 1).data) }
}
function closeChannels(actual: number[], expected: number[], tolerance = 2) {
  expected.forEach((channel, i) => expect(Math.abs(actual[i] - channel)).toBeLessThanOrEqual(tolerance))
}

describe('premultiplied linear DrawingML gradients', () => {
  test('opaque red fades to transparent blue over black without invisible blue bleeding', () => {
    const s = surface()
    paintGeometry(s.ctx, [rectangle], { issues: [], fill: fill([{ position: 0, color: color(255, 0, 0, 1) }, { position: 1, color: color(0, 0, 255, 0) }]) }, 256, 64)
    for (const x of [31, 63, 127, 191, 239]) {
      closeChannels(s.pixel(x), [255 * (1 - (x + .5) / 256), 0, 0, 255])
    }
  })
  test('three varying alpha stops interpolate premultiplied channels independently', () => {
    const s = surface()
    paintGeometry(s.ctx, [rectangle], { issues: [], fill: fill([{ position: 0, color: color(255, 0, 0, 1) }, { position: .5, color: color(0, 255, 0, .2) }, { position: 1, color: color(0, 0, 255, .8) }]) }, 256, 64)
    // Pixel centers; source contributions in each half are linearly interpolated.
    closeChannels(s.pixel(63), [255 * (1 - 63.5 / 128), 51 * (63.5 / 128), 0, 255])
    closeChannels(s.pixel(191), [0, 51 * (1 - 63.5 / 128), 204 * (63.5 / 128), 255])
  })
  test('duplicate hard stops retain source ordering while adjacent ramps use premultiplied colors', () => {
    const s = surface()
    const stops = [{ position: 0, color: color(255, 0, 0, 1) }, { position: .5, color: color(0, 0, 255, 0) }, { position: .5, color: color(0, 255, 0, 1) }, { position: 1, color: color(255, 0, 255, 0) }]
    paintGeometry(s.ctx, [rectangle], { issues: [], fill: fill(stops) }, 256, 64)
    closeChannels(s.pixel(120), [255 * (1 - 120.5 / 128), 0, 0, 255])
    closeChannels(s.pixel(128), [0, 255 * (1 - .5 / 128), 0, 255])
    closeChannels(s.pixel(191), [0, 255 * (1 - 63.5 / 128), 0, 255])
    expect(stops.map(s => s.position)).toEqual([0, .5, .5, 1])
  })
  test.each([1, .4])('opaque or constant alpha %s retains native Canvas bytes and original stop order', alpha => {
    const actual = surface(), expected = surface()
    // Deliberately unsorted; duplicate offsets keep insertion order.
    const stops = [{ position: 1, color: color(0, 0, 255, alpha) }, { position: .5, color: color(0, 255, 0, alpha) }, { position: .5, color: color(255, 0, 0, alpha) }, { position: 0, color: color(255, 255, 0, alpha) }]
    const native = expected.ctx.createLinearGradient(0, 32, 256, 32)
    for (const stop of stops) native.addColorStop(stop.position, `rgba(${stop.color.r},${stop.color.g},${stop.color.b},${alpha})`)
    expected.ctx.fillStyle = native; expected.ctx.fillRect(0, 0, 256, 64)
    paintGeometry(actual.ctx, [rectangle], { issues: [], fill: fill(stops) }, 256, 64)
    expect(actual.ctx.getImageData(0, 0, 256, 64).data).toEqual(expected.ctx.getImageData(0, 0, 256, 64).data)
  })
  test.each([1000, 10000])('bounds additional samples for %s source stops without nonfinite derived coordinates', count => {
    const s = surface(), added: Array<{ offset: number; css: string }> = []
    const original = s.ctx.createLinearGradient.bind(s.ctx)
    s.ctx.createLinearGradient = (...args) => {
      const gradient = original(...args), add = gradient.addColorStop.bind(gradient)
      gradient.addColorStop = (offset, css) => { added.push({ offset, css }); add(offset, css) }
      return gradient
    }
    const stops = Array.from({ length: count }, (_, i) => ({ position: i / (count - 1), color: color(i % 2 ? 255 : 0, 0, i % 2 ? 0 : 255, i % 2 ? 1 : .1) }))
    expect(paintGeometry(s.ctx, [rectangle], { issues: [], fill: fill(stops) }, 256, 64)).toEqual([])
    if (count < 4096) expect(added.length).toBeGreaterThan(stops.length)
    else expect(added.length).toBeGreaterThanOrEqual(stops.length)
    expect(added.length).toBeLessThanOrEqual(stops.length + 4096)
    expect(added.every(s => Number.isFinite(s.offset) && s.offset >= 0 && s.offset <= 1 && !/NaN|Infinity/.test(s.css))).toBe(true)
  })
  test('shaded translucent fill preserves caller transform, alpha, paint and clip', () => {
    const s = surface()
    s.ctx.beginPath(); s.ctx.rect(10, 10, 50, 50); s.ctx.clip()
    s.ctx.translate(10, 10); s.ctx.globalAlpha = .5; s.ctx.fillStyle = '#abcdef'; s.ctx.strokeStyle = '#123456'
    const transform = s.ctx.getTransform()
    paintGeometry(s.ctx, [{ ...rectangle, fill: 'darken' }], { issues: [], fill: fill([{ position: 0, color: color(255, 0, 0, 1) }, { position: 1, color: color(0, 0, 255, 0) }]) }, 256, 64)
    expect(s.ctx.getTransform()).toEqual(transform); expect(s.ctx.globalAlpha).toBe(.5); expect(s.ctx.strokeStyle).toBe('#123456')
    expect(s.pixel(70)).toEqual([0, 0, 0, 255]); expect(s.pixel(30)[2]).toBeLessThanOrEqual(1); expect(s.pixel(30)[0]).toBeGreaterThan(80)
    // node-canvas caches a stale gradient getter after restore; check actual paint state.
    const before = s.pixel(50, 55)
    s.ctx.fillRect(40, 45, 5, 5)
    closeChannels(s.pixel(50, 55), [(before[0] + 171) / 2, (before[1] + 205) / 2, (before[2] + 239) / 2, 255])
  })
})
