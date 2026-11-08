import { describe, expect, it } from 'bun:test'
import { cpSync, mkdtempSync, writeFileSync, readFileSync } from 'node:fs'
import { tmpdir } from 'node:os'
import { join } from 'node:path'
import { createCanvas, GlobalFonts } from '@napi-rs/canvas'
import { createRecorder } from '../../src/stream/record/recorder'
import { verifyPack } from '../../scripts/stream-s1/fontpack-lock'
import { faceFor, facesUsed, loadPack, PACK_DIR, parseFont, pinFonts, pinFontString, registerPack, requiredFacesOf, unitNeedsRaster } from '../src/fonts'

const pack = loadPack()

function record(paint: (ctx: CanvasRenderingContext2D) => void): Uint8Array {
  const c = createCanvas(200, 60)
  const rec = createRecorder(c.getContext('2d') as unknown as CanvasRenderingContext2D, { scale: 1, readImage: () => { throw new Error('no images') }, partialAlphaSnapshots: 'allow', patternSources: 'snapshot' })
  paint(rec.ctx)
  const out = rec.finish()
  if (out.discriminator !== 1) throw new Error(`tile: ${out.reason}`)
  return out.ops
}

describe('pinned font pack', () => {
  it('pack.json pins 20 faces + 6 aliases and matches the bytes', () => {
    expect(pack.faces.length).toBe(20)
    expect(Object.keys(pack.aliases).sort()).toEqual(['Arial', 'Calibri', 'Cambria', 'Courier New', 'Helvetica', 'Times New Roman'])
    expect(verifyPack(PACK_DIR)).toEqual([])
    const dir = mkdtempSync(join(tmpdir(), 'pack-'))
    cpSync(PACK_DIR, dir, { recursive: true })
    const f = join(dir, 'Caladea-Bold.ttf'), b = readFileSync(f); b[100] ^= 1; writeFileSync(f, b)
    expect(verifyPack(dir)).toEqual(['Caladea-Bold.ttf hash drift'])
  })

  it('aliases paint with the pack face, per weight', () => {
    registerPack(GlobalFonts)
    const draw = (font: string) => { const c = createCanvas(160, 40), x = pinFonts(c.getContext('2d'), pack); x.font = font; x.fillText('Hamburgefonts', 2, 30); return Buffer.from(x.getImageData(0, 0, 160, 40).data) }
    expect(draw('bold 20px Calibri').equals(draw('bold 20px Carlito'))).toBe(true)
    expect(draw('italic 20px "Times New Roman"').equals(draw('italic 20px "Liberation Serif"'))).toBe(true)
    expect(draw('bold 20px Calibri').equals(draw('20px Calibri'))).toBe(false)
  })

  it('pins alias names even when a system face of that name exists', () => {
    expect(pinFontString('italic 20px "Times New Roman"', pack)).toBe('italic 20px "Liberation Serif"')
    expect(pinFontString('bold 12pt Calibri, sans-serif', pack)).toBe('bold 12pt "Carlito", sans-serif')
    expect(pinFontString('12pt Aptos', pack)).toBe('12pt Aptos')
  })

  it('parses font shorthands', () => {
    expect(parseFont('italic bold 10.5pt Arial')).toEqual({ families: ['Arial'], weight: 700, italic: true })
    expect(parseFont('15.94923772935894pt "Calibri"')).toEqual({ families: ['Calibri'], weight: 400, italic: false })
    expect(parseFont('600 12px "Courier New", monospace')?.families).toEqual(['Courier New', 'monospace'])
    expect(faceFor('12px Aptos, Calibri', pack)).toBeUndefined()   // first family decides; no fallback counted
  })

  it('facesUsed follows font state through save/restore', () => {
    const ops = record((ctx) => {
      ctx.save(); ctx.font = 'bold 12pt Calibri'; ctx.fillText('a', 1, 20); ctx.restore()
      ctx.font = '10pt "Courier New"'; ctx.fillText('b', 1, 40)
    })
    expect(facesUsed(ops, pack)).toEqual({ faces: ['carlito-bold', 'liberation-mono-regular'], unpinned: [] })
    const aptos = record((ctx) => { ctx.font = '12pt Aptos'; ctx.fillText('c', 1, 20) })
    const use = facesUsed(aptos, pack)
    expect(use.faces).toEqual([]); expect(use.unpinned.length).toBe(1)
    expect(unitNeedsRaster(use, [])).toBe(true)
    expect(unitNeedsRaster({ faces: ['carlito-bold'], unpinned: [] }, ['carlito-bold'])).toBe(true)
    expect(unitNeedsRaster({ faces: ['carlito-bold'], unpinned: [] }, [])).toBe(false)
  })

  it('requiredFaces carries identity only', () => {
    const r = requiredFacesOf(['carlito-bold'], pack)
    expect(r).toEqual([{ id: 'carlito-bold', hash: pack.faces.find((f) => f.id === 'carlito-bold')!.hash96, alias: 'Calibri' }])
  })
})
