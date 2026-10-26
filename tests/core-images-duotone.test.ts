/**
 * `a:duotone` (ECMA-376 20.1.8.23): each pixel's luminance picks a colour on the line from the first
 * colour (luminance 0) to the second (luminance 1); the pixel's alpha is kept. Applied once, at decode.
 */
import { createCanvas } from 'canvas'
import { describe, expect, it } from 'vitest'
import { decodeImageAsset, duotonePixels } from '../src/core/images'
import { decodeImageAssets } from '../src/render/paint'

const DARK = { r: 30, g: 81, b: 85, a: 1 }
const LIGHT = { r: 200, g: 240, b: 100, a: 1 }

function png(pixels: Array<[number, number, number, number]>): Uint8Array {
  const c = createCanvas(pixels.length, 1)
  const ctx = c.getContext('2d')
  const img = ctx.createImageData(pixels.length, 1)
  pixels.forEach((p, i) => img.data.set(p, i * 4))
  ctx.putImageData(img, 0, 0)
  return new Uint8Array(c.toBuffer('image/png'))
}
function pixelsOf(image: CanvasImageSource | undefined): number[][] {
  const src = image as unknown as { width: number; height: number }
  const c = createCanvas(src.width, src.height)
  const ctx = c.getContext('2d')
  ctx.drawImage(image as never, 0, 0)
  const data = ctx.getImageData(0, 0, src.width, src.height).data
  return Array.from({ length: src.width * src.height }, (_, i) => Array.from(data.subarray(i * 4, i * 4 + 4)))
}

describe('duotonePixels', () => {
  it('maps black to the first colour, white to the second and grey between them', () => {
    const data = new Uint8ClampedArray([0, 0, 0, 255, 255, 255, 255, 255, 51, 51, 51, 255])
    duotonePixels(data, DARK, LIGHT)
    expect(Array.from(data)).toEqual([30, 81, 85, 255, 200, 240, 100, 255, 64, 113, 88, 255])
  })

  it('reads luminance with the Rec. 601 weights and keeps the pixel alpha', () => {
    // Pure green has luminance 0.587; a half-transparent pixel stays half transparent.
    const data = new Uint8ClampedArray([0, 255, 0, 128])
    duotonePixels(data, { r: 0, g: 0, b: 0, a: 1 }, { r: 255, g: 255, b: 255, a: 1 })
    expect(Array.from(data)).toEqual([150, 150, 150, 128])
  })

  it('a translucent duotone colour scales the pixel alpha', () => {
    const data = new Uint8ClampedArray([255, 255, 255, 200])
    duotonePixels(data, DARK, { ...LIGHT, a: 0.5 })
    expect(data[3]).toBe(100)
  })
})

describe('decodeImageAsset with a duotone', () => {
  it('returns the recoloured image', async () => {
    const data = png([[0, 0, 0, 255], [255, 255, 255, 255]])
    expect(pixelsOf(await decodeImageAsset({ data, mime: 'image/png' }))).toEqual([[0, 0, 0, 255], [255, 255, 255, 255]])
    expect(pixelsOf(await decodeImageAsset({ data, mime: 'image/png', duotone: { dark: DARK, light: LIGHT } }))).toEqual([[30, 81, 85, 255], [200, 240, 100, 255]])
  })

  it('reserves its canvas and pixel buffer through scratch first, and refuses when scratch does', async () => {
    const data = png([[0, 0, 0, 255], [255, 255, 255, 255]])
    const asked: number[] = []
    let released = 0
    const granted = await decodeImageAsset({ data, mime: 'image/png', duotone: { dark: DARK, light: LIGHT } }, undefined, bytes => { asked.push(bytes); return { release: () => { released++ } } })
    // Two 2 × 1 RGBA buffers, released once the recolour is done.
    expect(asked).toEqual([16])
    expect(released).toBe(1)
    expect(pixelsOf(granted)).toEqual([[30, 81, 85, 255], [200, 240, 100, 255]])
    expect(await decodeImageAsset({ data, mime: 'image/png', duotone: { dark: DARK, light: LIGHT } }, undefined, () => undefined)).toBeUndefined()
  })

  it('decodeImageAssets keeps a duotoned use of a picture apart from a plain use of the same bytes', async () => {
    const data = png([[255, 255, 255, 255]])
    const [plain, toned, again] = await decodeImageAssets([
      { data, mime: 'image/png' },
      { data, mime: 'image/png', duotone: { dark: DARK, light: LIGHT } },
      { data, mime: 'image/png', duotone: { dark: DARK, light: LIGHT } },
    ])
    expect(pixelsOf(plain)).toEqual([[255, 255, 255, 255]])
    expect(pixelsOf(toned)).toEqual([[200, 240, 100, 255]])
    // Identical uses still decode once.
    expect(again).toBe(toned)
  })
})
