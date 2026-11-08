/**
 * Every canvas-construction site (spec §7.3/§9.2, the nine sites T1 found)
 * defers to an installed CanvasBackend; with none installed, behaviour is
 * unchanged.
 */
import { afterEach, beforeEach, expect, it, vi } from 'vitest'
import { createCanvas, loadImage } from 'canvas'
import { installCanvasBackend, type CanvasBackend, type CanvasLike, type ImageLike } from '../../src/core/canvas-backend'
import { decodeImage, decodeImageAsset } from '../../src/core/images'
import { createLocalTextMeasurer } from '../../src/core/text-metrics'
import { buildTextIndex } from '../../src/core/search'
import { scratchSurface } from '../../src/drawing/paint'
import { getPatternTile, clearPatternTileCache } from '../../src/drawing/pattern'
import { rasteriseSvg } from '../../src/slideshow/render/svg-raster'

const calls: string[] = []
const counting: CanvasBackend = {
  name: 'counting-node-canvas',
  createCanvas(w, h) { calls.push('createCanvas'); return createCanvas(w, h) as unknown as CanvasLike },
  async decodeImage(bytes) { calls.push('decodeImage'); return await loadImage(Buffer.from(bytes)) as unknown as ImageLike },
  async rasteriseSvg(bytes) { calls.push('rasteriseSvg'); return await loadImage(Buffer.from(bytes)) as unknown as ImageLike },
}

function png(w: number, h: number): Uint8Array {
  const c = createCanvas(w, h)
  const ctx = c.getContext('2d')
  ctx.fillStyle = '#336699'
  ctx.fillRect(0, 0, w, h)
  return new Uint8Array(c.toBuffer('image/png'))
}
const svg = new TextEncoder().encode('<svg xmlns="http://www.w3.org/2000/svg" width="4" height="4"><rect width="4" height="4" fill="red"/></svg>')
const baseCtx = () => createCanvas(8, 8).getContext('2d') as unknown as CanvasRenderingContext2D

let domCanvas: ReturnType<typeof vi.spyOn>
beforeEach(() => {
  calls.length = 0
  clearPatternTileCache()
  const create = document.createElement.bind(document)
  domCanvas = vi.spyOn(document, 'createElement').mockImplementation(((tag: string, o?: ElementCreationOptions) => create(tag, o)) as typeof document.createElement)
})
afterEach(() => { installCanvasBackend(undefined); vi.restoreAllMocks() })

const domCanvases = () => domCanvas.mock.calls.filter(([tag]: unknown[]) => String(tag).toLowerCase() === 'canvas').length

it('scratch surfaces, pattern tiles and the local text measurer use the installed backend', () => {
  installCanvasBackend(counting)
  expect(scratchSurface(baseCtx(), 10, 10)).toBeDefined()
  expect(getPatternTile(baseCtx(), 'pct50', '#000000', '#ffffff')).toBeDefined()
  expect(createLocalTextMeasurer(baseCtx())('Hello').width).toBeGreaterThan(0)
  expect(calls.filter((c) => c === 'createCanvas').length).toBe(3)
  expect(domCanvases()).toBe(0)
})

it('image decode, duotone/2×2 surfaces, SVG raster and the search context use the installed backend', async () => {
  installCanvasBackend(counting)
  expect((await decodeImage(png(3, 3), 'image/png')).width).toBe(3)
  await decodeImageAsset({ data: png(2, 2), mime: 'image/png', flatTwoByTwo: true })
  const factory = (w: number, h: number) => Object.assign(createCanvas(w, h), {}) as unknown as HTMLCanvasElement
  expect(await rasteriseSvg(svg, 4, 4, factory as never)).toBeDefined()
  await buildTextIndex([])
  expect(calls).toContain('decodeImage')
  expect(calls).toContain('rasteriseSvg')
  // flatten surface + search tiny context
  expect(calls.filter((c) => c === 'createCanvas').length).toBeGreaterThanOrEqual(2)
  expect(domCanvases()).toBe(0)
})

it('with no backend installed the historical DOM path is unchanged', () => {
  expect(scratchSurface(baseCtx(), 10, 10)).toBeDefined()
  expect(calls).toEqual([])
})
