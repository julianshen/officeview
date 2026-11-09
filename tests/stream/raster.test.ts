import { describe, expect, it } from 'vitest'
import { createCanvas } from 'canvas'
import { decodeRasterBody, encodeRasterBody, pngChunkTypes, RasterBodyError } from '../../src/stream/raster'

const png = (w: number, h: number) => new Uint8Array(createCanvas(w, h).toBuffer('image/png'))

describe('UNIT_RASTER body', () => {
  it('round-trips', () => {
    const p = png(5, 3)
    expect(decodeRasterBody(encodeRasterBody(5, 3, p))).toEqual({ w: 5, h: 3, png: p })
    expect(pngChunkTypes(p)[0]).toBe('IHDR')
    expect(pngChunkTypes(p).at(-1)).toBe('IEND')
  })
  it('rejects an IHDR that disagrees with the body header', () => {
    expect(() => decodeRasterBody(encodeRasterBody(5, 4, png(5, 3)))).toThrow(/ihdr-size/)
  })
  it('admits w×h×4 against the cap before decoding', () => {
    expect(() => decodeRasterBody(encodeRasterBody(5, 3, png(5, 3)), { maxRgbaBytes: 59 })).toThrow(/decoded-cap/)
  })
  it('rejects truncation, a wrong length and a non-PNG', () => {
    const b = encodeRasterBody(5, 3, png(5, 3))
    expect(() => decodeRasterBody(b.subarray(0, 11))).toThrow(RasterBodyError)
    expect(() => decodeRasterBody(b.subarray(0, b.length - 1))).toThrow(/length/)
    const notPng = b.slice(); notPng[12] = 0
    expect(() => decodeRasterBody(notPng)).toThrow(/signature/)
  })
})
