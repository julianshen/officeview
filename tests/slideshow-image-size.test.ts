/**
 * Slideshow PR #18 review 2 follow-up: header sizing for BMP (and TIFF), so a raster is reserved before it
 * is decoded instead of refused. A BMP picture still paints in a show. jsdom with node-canvas; every clock
 * is driven. Decks are "Office-shaped synthetic".
 */
import { afterEach, expect, it } from 'vitest'
import { imageSize } from '../src/slideshow/player/stage'
import { buildPptxDeck } from '../src/testdata/ooxml-builders'
import { installShowSeams, loadDeck, startShow, type ShowSeams } from './helpers/slideshow-show-rig'

/** A bottom-up 24-bit BITMAPINFOHEADER BMP of one colour. */
function bmp(w: number, h: number, [r, g, b]: [number, number, number], topDown = false): Uint8Array {
  const row = (w * 3 + 3) & ~3
  const size = 54 + row * h
  const out = new Uint8Array(size)
  const v = new DataView(out.buffer)
  out[0] = 0x42; out[1] = 0x4d
  v.setUint32(2, size, true); v.setUint32(10, 54, true)
  v.setUint32(14, 40, true); v.setInt32(18, w, true); v.setInt32(22, topDown ? -h : h, true)
  v.setUint16(26, 1, true); v.setUint16(28, 24, true); v.setUint32(34, row * h, true)
  for (let y = 0; y < h; y++) for (let x = 0; x < w; x++) { const o = 54 + y * row + x * 3; out[o] = b; out[o + 1] = g; out[o + 2] = r }
  return out
}

it('sizes BMP headers: BITMAPINFOHEADER (top-down height is negative) and BITMAPCOREHEADER', () => {
  expect(imageSize(bmp(7, 5, [0, 0, 0]))).toEqual({ w: 7, h: 5 })
  expect(imageSize(bmp(7, 5, [0, 0, 0], true))).toEqual({ w: 7, h: 5 })
  const core = new Uint8Array(26 + 12)
  const v = new DataView(core.buffer)
  core[0] = 0x42; core[1] = 0x4d
  v.setUint32(14, 12, true); v.setUint16(18, 640, true); v.setUint16(20, 480, true)
  expect(imageSize(core)).toEqual({ w: 640, h: 480 })
})

it('sizes TIFF headers from IFD0 ImageWidth / ImageLength, little- and big-endian, SHORT and LONG', () => {
  const tiff = (le: boolean, wType: 3 | 4) => {
    const out = new Uint8Array(8 + 2 + 2 * 12 + 4)
    const v = new DataView(out.buffer)
    out.set(le ? [0x49, 0x49] : [0x4d, 0x4d]); v.setUint16(2, 42, le); v.setUint32(4, 8, le)
    v.setUint16(8, 2, le)
    const entry = (o: number, tag: number, type: number, value: number) => {
      v.setUint16(o, tag, le); v.setUint16(o + 2, type, le); v.setUint32(o + 4, 1, le)
      if (type === 3) v.setUint16(o + 8, value, le); else v.setUint32(o + 8, value, le)
    }
    entry(10, 256, wType, 1200); entry(22, 257, 3, 900)
    return out
  }
  expect(imageSize(tiff(true, 3))).toEqual({ w: 1200, h: 900 })
  expect(imageSize(tiff(false, 4))).toEqual({ w: 1200, h: 900 })
  expect(imageSize(new Uint8Array([0x49, 0x49, 42, 0, 8, 0, 0, 0]))).toBeUndefined()
})

let seams: ShowSeams | undefined
afterEach(() => { seams?.restore(); seams = undefined; document.body.replaceChildren() })

it('a BMP picture is reserved from its header and paints', async () => {
  seams = installShowSeams({ width: 960, height: 720 })
  const doc = await loadDeck(await buildPptxDeck({
    slideSize: [9144000, 6858000],
    slides: [{ shapes: [{ id: 2, off: ['0', '0'], ext: ['9144000', '6858000'], image: { data: bmp(4, 3, [0xcc, 0x33, 0x11]) } }] }],
  }))
  const r = await startShow(doc, { seams, size: { width: 960, height: 720 } })
  try {
    expect(r.diagnostics.filter(d => d.code === 'image-workspace')).toEqual([])
    const c = r.frameCanvas()!
    const px = c.getContext('2d')!.getImageData(c.width >> 1, c.height >> 1, 1, 1).data
    expect([...px.slice(0, 3)]).toEqual([0xcc, 0x33, 0x11])
  } finally { r.exit() }
})
