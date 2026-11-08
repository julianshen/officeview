import { describe, expect, it } from 'bun:test'
import { readFileSync } from 'node:fs'
import { resolve } from 'node:path'
import { loadImage } from '@napi-rs/canvas'
import { paintWatermark } from '../../src/core/watermark'
import { replayUnit } from '../../src/stream/replay'
import { pngChunkTypes } from '../../src/stream/raster'
import { DocumentHost, markOptions, RenderRefusal } from '../src/render'

const root = resolve(import.meta.dir, '../..')
const host = new DocumentHost()
const bytes = (name: string) => new Uint8Array(readFileSync(resolve(root, 'corpus', name)))
const DECK = 'pypptx-test.pptx', SHEET = 'poi-56295.xlsx'

async function rgbaOfPng(png: Uint8Array, w: number, h: number) {
  const img = await loadImage(Buffer.from(png))
  const c = host.backend.createCanvas(w, h), ctx = c.getContext('2d') as unknown as CanvasRenderingContext2D
  ctx.drawImage(img as never, 0, 0)
  return Buffer.from(ctx.getImageData(0, 0, w, h).data)
}

describe('render core', () => {
  it('opens once per (docRef, policy) and reports integer unit sizes', async () => {
    const before = host.parses
    const a = await host.open('d1', bytes(DECK), 'visible-only')
    const b = await host.open('d1', bytes(DECK), 'visible-only')
    expect(b.docKey).toBe(a.docKey)
    expect(host.parses - before).toBe(1)
    expect(a.units.length).toBeGreaterThan(0)
    for (const u of a.units) { expect(Number.isInteger(u.w)).toBe(true); expect(u.w).toBeGreaterThan(0) }
    expect(a.layoutFingerprint).toMatch(/^[0-9a-f]{64}$/)
    const full = await host.open('d1', bytes(DECK), 'full')
    expect(full.docKey).not.toBe(a.docKey)
    expect(full.layoutFingerprint).not.toBe(a.layoutFingerprint)
  })

  it('final-raster is metadata-free and equals a direct paint + the visible mark', async () => {
    const { docKey } = await host.open('d1', bytes(DECK), 'visible-only')
    const r = host.render(docKey, { unitIndex: 0, scale: 1, projection: 'final-raster', mark: { kind: 'fixed' } })
    if (r.kind !== 'raster') throw new Error('expected raster')
    expect(new Set(pngChunkTypes(r.png))).toEqual(new Set(['IHDR', 'IDAT', 'IEND']))
    const off = host.render(docKey, { unitIndex: 0, scale: 1, projection: 'final-raster', mark: { kind: 'off' } })
    if (off.kind !== 'raster') throw new Error('expected raster')
    // Direct: unmarked raster + the mark painted the same way.
    const c = host.backend.createCanvas(r.w, r.h), ctx = c.getContext('2d') as unknown as CanvasRenderingContext2D
    ctx.drawImage(await loadImage(Buffer.from(off.png)) as never, 0, 0)
    paintWatermark(ctx, { widthPx: r.w, heightPx: r.h }, markOptions({ kind: 'fixed' }))
    expect((await rgbaOfPng(r.png, r.w, r.h)).equals(Buffer.from(ctx.getImageData(0, 0, r.w, r.h).data))).toBe(true)
    expect((await rgbaOfPng(r.png, r.w, r.h)).equals(await rgbaOfPng(off.png, r.w, r.h))).toBe(false)
  })

  it('a commands unit replays to the identical pixels of the marked raster', async () => {
    const outcomes: string[] = []
    for (const name of [DECK, SHEET]) {
      const { docKey } = await host.open(`d-${name}`, bytes(name), 'visible-only')
      for (const scale of [1, 2]) {
        const cmd = host.render(docKey, { unitIndex: 0, scale, projection: 'commands', mark: { kind: 'viewer', text: 'viewer-7' } })
        const ras = host.render(docKey, { unitIndex: 0, scale, projection: 'final-raster', mark: { kind: 'viewer', text: 'viewer-7' } })
        if (ras.kind !== 'raster') throw new Error('raster expected')
        outcomes.push(`${name}@${scale}:${cmd.kind === 'commands' ? 'commands' : cmd.fallback}`)
        if (cmd.kind !== 'commands') { expect(cmd.fallback).toBeDefined(); continue }
        const c = host.backend.createCanvas(cmd.w, cmd.h), ctx = c.getContext('2d') as unknown as CanvasRenderingContext2D
        await replayUnit(ctx, { ops: cmd.ops, assets: new Map(cmd.assets.map((a) => [a.id, a])) }, host.backend)
        expect(Buffer.from(ctx.getImageData(0, 0, cmd.w, cmd.h).data).equals(await rgbaOfPng(ras.png, ras.w, ras.h)), `${name}@${scale}`).toBe(true)
      }
    }
    console.log(`[render] ${outcomes.join(' ')}`)
    expect(outcomes.filter((o) => o.endsWith(':commands')).length).toBeGreaterThanOrEqual(2)
  })

  it('a missing face turns a command unit into a raster', async () => {
    const { docKey } = await host.open('d1', bytes(DECK), 'visible-only')
    const cmd = host.render(docKey, { unitIndex: 0, scale: 1, projection: 'commands', mark: { kind: 'off' } })
    if (cmd.kind !== 'commands' || cmd.faces.length === 0) return expect(cmd.kind === 'raster' || cmd.faces.length === 0).toBe(true)
    const r = host.render(docKey, { unitIndex: 0, scale: 1, projection: 'commands', mark: { kind: 'off' }, missingFaces: [cmd.faces[0]!] })
    expect(r).toMatchObject({ kind: 'raster', fallback: 'face-missing' })
  })

  it('refuses a bad unit index, scale, and unknown documents', async () => {
    const { docKey } = await host.open('d1', bytes(DECK), 'visible-only')
    for (const job of [{ unitIndex: 999 }, { unitIndex: -1 }, { unitIndex: 0.5 }, { unitIndex: 0, scale: 3 }, { unitIndex: 0, scale: Number.NaN }]) {
      expect(() => host.render(docKey, { scale: 1, projection: 'final-raster', mark: { kind: 'off' }, ...job })).toThrow(RenderRefusal)
    }
    expect(() => host.render('nope', { unitIndex: 0, scale: 1, projection: 'final-raster', mark: { kind: 'off' } })).toThrow(/unknown-doc/)
  })

  it('refuses ODT/garbage before parsing', async () => {
    await expect(host.open('x', new Uint8Array([1, 2, 3]), 'visible-only')).rejects.toThrow(/unsupported-format/)
  })
})
