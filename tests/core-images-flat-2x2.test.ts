/**
 * tdf146223: PowerPoint paints a stretched picture that is exactly 2 × 2 pixels as one colour. Probe decks
 * exported as movies from PowerPoint 16.114 measured it: a PNG gets the plain mean of its four pixels
 * (one white pixel on black → 64, the deck's greys 0/57/32/92 → 45); a JPEG gets a top-left-weighted
 * sample (one white pixel top-left → 90, bottom-right → 42, the deck's greys → 37). 1 × 1, 1 × 2, 2 × 1,
 * 2 × 3, 3 × 2 and larger pictures are stretched smoothly as before.
 */
import { createCanvas } from 'canvas'
import { describe, expect, it } from 'vitest'
import { decodeImageAsset, flatTwoByTwoPixels } from '../src/core/images'

const grey = (...values: number[]) => new Uint8ClampedArray(values.flatMap(v => [v, v, v, 255]))
const first = (data: Uint8ClampedArray) => Array.from(data.subarray(0, 4))

function png(width: number, height: number, values: number[]): Uint8Array {
  const c = createCanvas(width, height)
  const ctx = c.getContext('2d')
  const img = ctx.createImageData(width, height)
  img.data.set(grey(...values))
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

describe('flatTwoByTwoPixels', () => {
  it('a PNG becomes the mean of its four pixels', () => {
    for (const [values, want] of [[[255, 0, 0, 0], 64], [[0, 0, 0, 255], 64], [[0, 0, 255, 255], 128], [[0, 57, 32, 92], 45]] as const) {
      const data = grey(...values)
      flatTwoByTwoPixels(data, false)
      expect(Array.from(data), values.join()).toEqual(Array.from(grey(want, want, want, want)))
    }
  })

  it('a JPEG weights the top-left pixel most, as PowerPoint does', () => {
    for (const [values, want] of [[[255, 0, 0, 0], 90], [[0, 0, 0, 255], 42], [[0, 255, 0, 255], 104], [[128, 128, 128, 0], 106], [[0, 57, 32, 92], 37]] as const) {
      const data = grey(...values)
      flatTwoByTwoPixels(data, true)
      expect(first(data)[0], values.join()).toBeCloseTo(want, -0.5)
      expect(Array.from(data)).toEqual([...first(data), ...first(data), ...first(data), ...first(data)])
    }
  })

  it('transparent pixels add no colour', () => {
    const data = new Uint8ClampedArray([255, 0, 0, 255, 0, 0, 255, 0, 0, 0, 255, 0, 255, 0, 0, 255])
    flatTwoByTwoPixels(data, false)
    expect(first(data)).toEqual([255, 0, 0, 128])
  })
})

describe('decodeImageAsset with flatTwoByTwo', () => {
  it('flattens a 2 × 2 picture to one colour and keeps its size', async () => {
    const image = await decodeImageAsset({ data: png(2, 2, [0, 57, 32, 92]), mime: 'image/png', flatTwoByTwo: true })
    expect(pixelsOf(image)).toEqual(Array(4).fill([45, 45, 45, 255]))
  })

  it('leaves other sizes, and pictures without the flag, as they are', async () => {
    for (const [w, h, values] of [[1, 1, [100]], [1, 2, [0, 255]], [2, 1, [0, 255]], [2, 3, [0, 255, 0, 255, 0, 255]], [3, 2, [0, 128, 255, 0, 128, 255]]] as const) {
      const image = await decodeImageAsset({ data: png(w, h, [...values]), mime: 'image/png', flatTwoByTwo: true })
      expect(pixelsOf(image).map(p => p[0]), `${w}x${h}`).toEqual([...values])
    }
    const plain = await decodeImageAsset({ data: png(2, 2, [0, 57, 32, 92]), mime: 'image/png' })
    expect(pixelsOf(plain).map(p => p[0])).toEqual([0, 57, 32, 92])
  })
})
