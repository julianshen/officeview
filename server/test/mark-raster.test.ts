/** Stage 2a Task 2: lossless PNG codec + markRaster (the per-session encode job body). */
import { describe, expect, it } from 'bun:test'
import { readFileSync } from 'node:fs'
import { resolve } from 'node:path'
import { createCanvas } from '@napi-rs/canvas'
import { decodePng, encodePng, PngUnsupported } from '../src/mark/png'
import { markRaster } from '../src/mark/mark-raster'
import { basis, deriveBasisKey } from '../src/mark/basis'
import { extendedCodeword } from '../src/mark/bch'
import { messageFor } from '../src/mark/message'
import { DEFAULT_PARAMS } from '../src/mark/params'
import { DocumentHost } from '../src/render'
import { decideScores, scoreCapture } from '../../scripts/stream-s0/mark/detect'

const root = resolve(import.meta.dir, '../..')
const TAU = { abs: 7.99, margin: 3.85 }
const key = Uint8Array.from({ length: 32 }, (_, i) => i + 1)
const codeFor = (token: number) => extendedCodeword(messageFor(key, 1, token, 17))

function lcg(seed: number) { let s = seed >>> 0; return () => (s = (Math.imul(s, 1664525) + 1013904223) >>> 0) >>> 24 }
function skiaSurface(w: number, h: number, paint: (ctx: CanvasRenderingContext2D) => void) {
  const c = createCanvas(w, h), ctx = c.getContext('2d') as unknown as CanvasRenderingContext2D
  paint(ctx)
  return { png: new Uint8Array(c.toBuffer('image/png')), rgba: new Uint8ClampedArray(ctx.getImageData(0, 0, w, h).data) }
}

describe('png codec', () => {
  it('decodes an opaque Skia PNG to the surface pixels', () => {
    const s = skiaSurface(203, 97, (ctx) => { ctx.fillStyle = '#fff'; ctx.fillRect(0, 0, 203, 97); ctx.fillStyle = '#c33'; ctx.font = '30px sans-serif'; ctx.fillText('Mark', 10, 60) })
    const d = decodePng(s.png)
    expect([d.w, d.h]).toEqual([203, 97])
    expect(Buffer.from(d.rgba).equals(Buffer.from(s.rgba))).toBe(true)
  })
  it('decodes alpha 0/128/255 regions as straight alpha', () => {
    const s = skiaSurface(64, 32, (ctx) => { ctx.fillStyle = 'rgba(40,90,200,1)'; ctx.fillRect(0, 0, 20, 32); ctx.fillStyle = 'rgba(40,90,200,0.5)'; ctx.fillRect(20, 0, 20, 32) })
    const d = decodePng(s.png)
    expect(Buffer.from(d.rgba).equals(Buffer.from(s.rgba))).toBe(true)
    expect(d.rgba[(5 * 64 + 25) * 4 + 3]).toBe(128)
    expect(d.rgba[(5 * 64 + 50) * 4 + 3]).toBe(0)
  })
  for (const [w, h] of [[1, 1], [1, 37], [41, 1], [129, 77]] as const) it(`round-trips random RGBA ${w}x${h} bit-exactly`, () => {
    const r = lcg(w * 1000 + h), rgba = Uint8ClampedArray.from({ length: w * h * 4 }, r)
    const d = decodePng(encodePng(w, h, rgba))
    expect([d.w, d.h]).toEqual([w, h])
    expect(Buffer.from(d.rgba).equals(Buffer.from(rgba))).toBe(true)
  })
  it('emits only IHDR/IDAT/IEND and refuses unsupported PNGs', () => {
    const png = encodePng(3, 2, new Uint8ClampedArray(24).fill(9))
    const types: string[] = []
    for (let at = 8; at < png.length;) { const len = new DataView(png.buffer, png.byteOffset).getUint32(at); types.push(String.fromCharCode(...png.subarray(at + 4, at + 8))); at += 12 + len }
    expect(types).toEqual(['IHDR', 'IDAT', 'IEND'])
    const bad = png.slice(); bad[8 + 8 + 8] = 16 // bit depth 16
    expect(() => decodePng(bad)).toThrow(PngUnsupported)
    expect(() => decodePng(new Uint8Array(10))).toThrow(PngUnsupported)
  })
})

describe('markRaster', async () => {
  const host = new DocumentHost()
  const doc = await host.open('pydocx', new Uint8Array(readFileSync(resolve(root, 'corpus/pydocx-test.docx'))), 'visible-only')
  const unit = host.render(doc.docKey, { unitIndex: 0, scale: 1, projection: 'final-raster', mark: { kind: 'fixed' } }) as { png: Uint8Array; w: number; h: number }
  const kBasis = deriveBasisKey(key, 'docref-poi', '0@1')
  const canonical = decodePng(unit.png)
  const out = markRaster({ png: unit.png, w: unit.w, h: unit.h, kBasis, code: codeFor(3) })

  it('touches only chip support; every other pixel (alpha included) is identical', () => {
    if ('unmarked' in out) throw new Error(`unexpected unmarked: ${out.unmarked}`)
    expect(out.carriers + out.masked).toBe(512)
    expect(out.carriers).toBeGreaterThan(400)
    console.log(`[mark-raster] ${unit.w}x${unit.h} skia=${unit.png.length}B reencoded=${out.png.length}B carriers=${out.carriers} pedestal=${out.pedestalCells}`)
    const marked = decodePng(out.png)
    const inChip = new Uint8Array(unit.w * unit.h)
    for (const c of basis(kBasis, unit.w, unit.h, DEFAULT_PARAMS)) for (let y = c.y; y < c.y + c.cell; y++) for (let x = c.x; x < c.x + c.cell; x++) inChip[y * unit.w + x] = 1
    let diffOutside = 0, diffInside = 0
    for (let p = 0; p < unit.w * unit.h; p++) for (let ch = 0; ch < 4; ch++) if (marked.rgba[p * 4 + ch] !== canonical.rgba[p * 4 + ch]) { if (inChip[p] && ch < 3) diffInside++; else diffOutside++ }
    expect(diffOutside).toBe(0)
    expect(diffInside).toBeGreaterThan(0)
  })
  it('detects to the embedded code at Stage-0 tau over 64 candidates', () => {
    if ('unmarked' in out) throw new Error('unmarked')
    const chips = basis(kBasis, unit.w, unit.h, DEFAULT_PARAMS)
    const codes = Array.from({ length: 64 }, (_, i) => codeFor(i))
    const bmp = (d: { w: number; h: number; rgba: Uint8ClampedArray }) => ({ width: d.w, height: d.h, data: d.rgba })
    const decision = decideScores(scoreCapture(bmp(decodePng(out.png)), bmp(canonical), chips, codes, DEFAULT_PARAMS), TAU)
    expect(decision).toMatchObject({ kind: 'attributed', index: 3 })
  })
  it('a unit below carrier capacity ships unmarked with the input bytes', () => {
    const s = skiaSurface(100, 50, (ctx) => { ctx.fillStyle = '#fff'; ctx.fillRect(0, 0, 100, 50) })
    const r = markRaster({ png: s.png, w: 100, h: 50, kBasis, code: codeFor(1) })
    expect(r).toEqual({ unmarked: 'capacity', png: s.png })
  })
  it('an all-transparent surface ships unmarked (every chip masked)', () => {
    const s = skiaSurface(512, 512, () => {})
    const r = markRaster({ png: s.png, w: 512, h: 512, kBasis, code: codeFor(1) })
    expect('unmarked' in r && r.unmarked).toBe('all-masked')
    expect(r.png).toBe(s.png)
  })
})
