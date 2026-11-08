// @vitest-environment node
import { createCanvas, loadImage } from 'canvas'
import { describe, expect, it } from 'vitest'
import { RECORD_TEXT } from '../../src/core/text-recording'
import { createRecorder, AssetTable } from '../../src/stream/record/recorder'
import { Opcode, OpReader } from '../../src/stream/ops'
import { replayUnit } from '../../src/stream/replay'
import { listCorpus, loadUnits, sizeCeil } from '../../scripts/stream-s0/corpus'

const context = (w = 180, h = 100) => createCanvas(w, h).getContext('2d') as unknown as CanvasRenderingContext2D
const pixels = (c: CanvasRenderingContext2D) => Buffer.from(c.getImageData(0, 0, c.canvas.width, c.canvas.height).data)
const backend = {
  createCanvas: (w: number, h: number) => createCanvas(w, h) as unknown as HTMLCanvasElement,
  async imageFromPixels(w: number, h: number, rgba: Uint8Array) {
    const c = createCanvas(w, h), ctx = c.getContext('2d'), image = ctx.createImageData(w, h)
    image.data.set(rgba); ctx.putImageData(image, 0, 0)
    return await loadImage(c.toBuffer('image/png')) as unknown as HTMLImageElement
  },
}
async function compare(paint: (ctx: CanvasRenderingContext2D) => void, w = 180, h = 100) {
  const direct = context(w, h), backing = context(w, h), replay = context(w, h)
  paint(direct)
  const recorder = createRecorder(backing, { scale: 1, patternSources: process.env.S0_BACKEND === 'napi' ? 'snapshot' : 'live' }); paint(recorder.ctx)
  const unit = recorder.finish(); expect(unit.discriminator, unit.discriminator === 3 ? unit.reason : '').toBe(1)
  if (unit.discriminator !== 1) throw new Error(unit.reason)
  expect(Buffer.compare(pixels(direct), pixels(backing))).toBe(0)
  await replayUnit(replay, { ops: unit.ops, assets: new Map(unit.assets.map(a => [a.id, a])) }, backend)
  expect(Buffer.compare(pixels(direct), pixels(replay))).toBe(0)
  return unit
}

describe('recorder facade', () => {
  it('records path, transform, alpha and nested clips/save/restore in order', async () => {
    const unit = await compare(ctx => {
      ctx.fillStyle = 'white'; ctx.fillRect(0, 0, 180, 100)
      ctx.save(); ctx.translate(6, 7); ctx.beginPath(); ctx.rect(0, 0, 100, 75); ctx.clip()
      ctx.save(); ctx.scale(1.2, .8); ctx.globalAlpha = .6; ctx.fillStyle = '#f12'; ctx.fillRect(0, 0, 70, 60); ctx.restore()
      ctx.beginPath(); ctx.moveTo(0, 0); ctx.lineTo(90, 40); ctx.quadraticCurveTo(30, 60, 10, 70); ctx.closePath(); ctx.stroke(); ctx.restore()
    })
    const ops = [...new OpReader(unit.ops)]
    expect(ops.filter(o => o.code === Opcode.STATE_PUSH)).toHaveLength(2)
    expect(ops.filter(o => o.code === Opcode.STATE_POP)).toHaveLength(2)
    expect(ops.some(o => o.code === Opcode.TRANSFORM)).toBe(true)
  })
  it('snapshots a scratch canvas at each draw, including redraws', async () => {
    const unit = await compare(ctx => {
      const scratch = createCanvas(20, 20), s = scratch.getContext('2d')
      s.fillStyle = 'blue'; s.fillRect(0, 0, 20, 20); ctx.drawImage(scratch as unknown as CanvasImageSource, 0, 0)
      s.fillStyle = 'red'; s.fillRect(0, 0, 20, 20); ctx.drawImage(scratch as unknown as CanvasImageSource, 30, 0, 40, 40)
      ctx.drawImage(scratch as unknown as CanvasImageSource, 2, 3, 10, 11, 80, 10, 20, 22)
    })
    expect(unit.assets.filter(a => a.type === 'bitmap')).toHaveLength(2)
    expect(new Set(unit.assets.map(a => a.hash)).size).toBe(2)
  })
  it('preserves native measurement branches, maxWidth and strokeText', async () => {
    await compare(ctx => {
      ctx.font = '20px sans-serif'; ctx.fillStyle = '#034'; ctx.strokeStyle = '#e30'
      ctx.fillRect(0, 0, ctx.measureText('Hello').width > 30 ? 40 : 10, 7)
      ctx.fillText('Hello 🦊', 6, 35, 65); ctx.strokeText('World', 10, 72)
    })
  })
  it('preserves shadows and dash state', async () => {
    await compare(ctx => {
      ctx.shadowColor = 'rgba(0,0,255,0.8)'; ctx.shadowBlur = 4; ctx.shadowOffsetX = 3; ctx.shadowOffsetY = 5
      ctx.strokeStyle = 'red'; ctx.lineWidth = 3; ctx.setLineDash([4, 2]); ctx.lineDashOffset = 1
      expect(ctx.getLineDash()).toEqual([4, 2]); ctx.strokeRect(20, 15, 60, 40)
    })
  })
  it('preserves linear/radial gradients and mutation after assignment', async () => {
    await compare(ctx => {
      const g = ctx.createLinearGradient(0, 0, 70, 0); g.addColorStop(0, 'red'); g.addColorStop(1, 'blue')
      ctx.fillStyle = g; ctx.fillRect(0, 0, 70, 30); g.addColorStop(.5, 'yellow'); ctx.fillRect(0, 35, 70, 30)
      const r = ctx.createRadialGradient(110, 40, 1, 110, 40, 30); r.addColorStop(0, 'white'); r.addColorStop(1, 'green')
      ctx.fillStyle = r; ctx.fillRect(75, 0, 70, 75)
    })
  })
  it('preserves a gradient created and assigned under different transforms', async () => {
    await compare(ctx => {
      ctx.translate(10, 5); const g = ctx.createLinearGradient(0, 0, 50, 0); g.addColorStop(0, 'red'); g.addColorStop(1, 'blue')
      ctx.translate(20, 0); ctx.fillStyle = g; ctx.fillRect(0, 0, 60, 50)
    })
  })
  it('preserves patterns and wrapper setTransform', async () => {
    await compare(ctx => {
      const source = createCanvas(8, 8), s = source.getContext('2d'); s.fillStyle = 'red'; s.fillRect(0, 0, 4, 8)
      const p = ctx.createPattern(source as unknown as CanvasImageSource, 'repeat')!
      const m = ctx.getTransform(); m.a = 2; m.d = 2; m.e = 3; m.f = 4; p.setTransform(m)
      ctx.fillStyle = p; ctx.fillRect(0, 0, 75, 60)
      m.e = 7; p.setTransform(m); ctx.fillRect(85, 0, 75, 60)
    })
  })
  it('forwards reads without exposing RECORD_TEXT or reading devicePixelRatio', () => {
    const ctx = context(); Object.defineProperty(ctx, RECORD_TEXT, { value: () => {}, configurable: true })
    const previous = Object.getOwnPropertyDescriptor(globalThis, 'devicePixelRatio')
    Object.defineProperty(globalThis, 'devicePixelRatio', { configurable: true, get() { throw new Error('DPR read') } })
    try {
      const rec = createRecorder(ctx, { scale: 2 }); expect(RECORD_TEXT in rec.ctx).toBe(false); expect((rec.ctx as any)[RECORD_TEXT]).toBeUndefined()
      expect('letterSpacing' in rec.ctx).toBe('letterSpacing' in ctx)
      rec.ctx.translate(4, 5); expect(rec.ctx.getTransform().e).toBe(4); expect(rec.finish().discriminator).toBe(1)
    } finally { if (previous) Object.defineProperty(globalThis, 'devicePixelRatio', previous); else Reflect.deleteProperty(globalThis, 'devicePixelRatio') }
  })
  it('latches the first unsupported operation and still paints the backing canvas', () => {
    const backing = context(); const rec = createRecorder(backing, { scale: 1 }); const img = backing.createImageData(1, 1); img.data.set([255, 0, 0, 255])
    rec.ctx.putImageData(img, 0, 0); (rec.ctx as any).unsupportedState = 1
    expect(rec.finish()).toEqual({ discriminator: 3, reason: 'putImageData' })
    expect([...backing.getImageData(0, 0, 1, 1).data]).toEqual([255, 0, 0, 255])
  })
  it('hashes canonical owned asset bytes and deduplicates snapshots', () => {
    const assets = new AssetTable(), rgba = Uint8Array.of(255, 0, 0, 255)
    const id = assets.add({ type: 'bitmap', source: 'canvas', w: 1, h: 1, rgba })
    expect(assets.add({ type: 'bitmap', source: 'canvas', w: 1, h: 1, rgba: rgba.slice() })).toBe(id)
    rgba.fill(0); const a = assets.all()[0]; expect(a.hash).toMatch(/^[a-f0-9]{64}$/)
    if (a.type === 'bitmap') expect([...a.rgba]).toEqual([255, 0, 0, 255])
  })
})

it.each(['poi-55733.docx', 'poi-text-highlight.pptx'])('corpus unit 0 zero-diff: %s', async name => {
  const file = listCorpus().find(f => f.name === name); expect(file).toBeDefined()
  const units = await loadUnits(file!)
  try { const u = units[0]; await compare(u.paint, sizeCeil(u.widthPx), sizeCeil(u.heightPx)) } finally { units.dispose() }
}, 60_000)

it('captures decoded images via a synchronous backend reader, or explicitly falls back', async () => {
  const { loadImage } = await import('canvas')
  const source = createCanvas(12, 12); source.getContext('2d').fillRect(0, 0, 12, 12)
  const image = await loadImage(source.toBuffer('image/png'))
  const backing = context(), result = context(), direct = context()
  direct.drawImage(image as unknown as CanvasImageSource, 4, 5)
  let reads = 0
  const recorder = createRecorder(backing, { scale: 1, readImage(input) {
    reads++
    const c = createCanvas(12, 12), ctx = c.getContext('2d'); ctx.drawImage(input as never, 0, 0)
    return { w: 12, h: 12, rgba: new Uint8Array(ctx.getImageData(0, 0, 12, 12).data) }
  } })
  recorder.ctx.drawImage(image as unknown as CanvasImageSource, 4, 5)
  const unit = recorder.finish(); expect(unit.discriminator, unit.discriminator === 3 ? unit.reason : '').toBe(1); expect(reads).toBe(1)
  if (unit.discriminator !== 1) throw new Error(unit.reason)
  await replayUnit(result, { ops: unit.ops, assets: new Map(unit.assets.map(a => [a.id, a])) }, backend)
  expect(Buffer.compare(pixels(direct), pixels(result))).toBe(0)
  const withoutReader = createRecorder(context(), { scale: 1 }); withoutReader.ctx.drawImage(image as unknown as CanvasImageSource, 0, 0)
  expect(withoutReader.finish()).toEqual({ discriminator: 3, reason: 'drawImage' })
})
it('captures preconfigured scale and styles without a DPR lookup', async () => {
  const backing = context(), output = context(), direct = context(); direct.scale(2, 2); direct.fillStyle = 'green'; direct.globalAlpha = .5; direct.fillRect(3, 4, 10, 11); backing.scale(2, 2); backing.fillStyle = 'green'; backing.globalAlpha = .5
  const recorder = createRecorder(backing, { scale: 2 }); recorder.ctx.fillRect(3, 4, 10, 11)
  const unit = recorder.finish(); if (unit.discriminator !== 1) throw new Error(unit.reason)
  await replayUnit(output, { ops: unit.ops, assets: new Map(unit.assets.map(a => [a.id, a])) }, backend)
  expect(Buffer.compare(pixels(direct), pixels(output))).toBe(0)
})
it('refuses partially transparent snapshots instead of claiming a lossy command replay', () => {
  const paint = (ctx: CanvasRenderingContext2D) => {
    ctx.fillStyle = '#41aa72'; ctx.fillRect(0, 0, 180, 100)
    const c = createCanvas(30, 30), s = c.getContext('2d')
    s.fillStyle = 'rgba(153,77,211,.31)'; s.fillRect(.3, .4, 23.2, 18.1)
    ctx.drawImage(c as unknown as CanvasImageSource, 12, 13)
  }
  const backing = context(), direct = context(), recorder = createRecorder(backing, { scale: 1 })
  paint(direct); paint(recorder.ctx)
  expect(Buffer.compare(pixels(backing), pixels(direct))).toBe(0)
  expect(recorder.finish()).toEqual({ discriminator: 3, reason: 'drawImage' })
})


it.each(['rgba(0,0,0,0.875)', 'rgba(0,0,0,0.395)', '#00000040'])('keeps assigned colour precision: %s', async color => {
  await compare(ctx => {
    ctx.fillStyle = color; ctx.fillRect(0, 0, 30, 30)
    ctx.strokeStyle = color; ctx.lineWidth = 5; ctx.strokeRect(40, 10, 30, 30)
    ctx.shadowColor = color; ctx.shadowOffsetX = 15; ctx.fillStyle = 'red'; ctx.fillRect(90, 10, 10, 10)
  })
})
it('refuses ambiguous initial partial-alpha colours', () => {
  const ctx = context(); ctx.fillStyle = 'rgba(0,0,0,.875)'
  expect(createRecorder(ctx, { scale: 1 }).finish()).toEqual({ discriminator: 3, reason: 'initial-state' })
})
it('snapshots live pattern source mutations at paint time', async () => {
  await compare(ctx => {
    const c = createCanvas(8, 8), s = c.getContext('2d')
    s.fillStyle = 'red'; s.fillRect(0, 0, 8, 8)
    ctx.fillStyle = ctx.createPattern(c as unknown as CanvasImageSource, 'repeat')!
    ctx.fillRect(0, 0, 30, 30)
    s.fillStyle = 'blue'; s.fillRect(0, 0, 8, 8)
    ctx.fillRect(40, 0, 30, 30)
  })
})
it.each(['getContext', 'width', 'height'])('latches backing canvas escape: %s', key => {
  const rec = createRecorder(context(), { scale: 1 })
  expect(rec.ctx.canvas.width).toBe(180); expect(rec.ctx.canvas.height).toBe(100)
  if (key === 'getContext') rec.ctx.canvas.getContext('2d')
  else (rec.ctx.canvas as any)[key] = 20
  expect(rec.finish()).toEqual({ discriminator: 3, reason: `canvas.${key}` })
})
it('refuses a readImage whose dimensions differ from its decoded source', async () => {
  const c = createCanvas(12, 12), image = await loadImage(c.toBuffer('image/png'))
  const rec = createRecorder(context(), { scale: 1, readImage: () => ({ w: 1, h: 1, rgba: new Uint8Array(4) }) })
  rec.ctx.drawImage(image as unknown as CanvasImageSource, 0, 0)
  expect(rec.finish()).toEqual({ discriminator: 3, reason: 'drawImage' })
})
it('includes source kind in bitmap deduplication', () => {
  const table = new AssetTable(), rgba = Uint8Array.of(255, 0, 0, 255)
  const canvas = table.add({ type: 'bitmap', source: 'canvas', w: 1, h: 1, rgba })
  const image = table.add({ type: 'bitmap', source: 'image', w: 1, h: 1, rgba })
  expect(image).not.toBe(canvas)
})
it('preserves evenodd fill and saved gradient mutations', async () => {
  await compare(ctx => {
    ctx.beginPath(); ctx.rect(0, 0, 40, 40); ctx.rect(10, 10, 20, 20); ctx.fill('evenodd')
    const g = ctx.createLinearGradient(50, 0, 100, 0); g.addColorStop(0, 'red'); g.addColorStop(1, 'blue')
    ctx.fillStyle = g; ctx.save(); ctx.fillStyle = 'black'; g.addColorStop(.5, 'yellow'); ctx.restore()
    ctx.fillRect(50, 0, 70, 50)
  })
})

it.skipIf(process.env.S0_BACKEND !== 'napi')('replays JPEG image sampling at both scales and all qualities exactly', async () => {
  const source = createCanvas(320, 240), s = source.getContext('2d'), data = s.createImageData(320, 240)
  for (let y = 0; y < 240; y++) for (let x = 0; x < 320; x++) {
    const at = (y * 320 + x) * 4
    data.data.set([(x * 13 + y * 3) % 256, (x * 7 ^ y * 11) % 256, (x * y) % 256, 255], at)
  }
  s.putImageData(data, 0, 0)
  const image = await loadImage(source.toBuffer('image/jpeg'))
  for (const [w, h] of [[137, 103], [800, 600]]) for (const quality of ['low', 'medium', 'high'] as const) {
    const direct = context(w, h), backing = context(w, h), result = context(w, h)
    const recorder = createRecorder(backing, { scale: 1, readImage(input) {
      const c = createCanvas(320, 240), cctx = c.getContext('2d'); cctx.drawImage(input as never, 0, 0)
      return { w: 320, h: 240, rgba: new Uint8Array(cctx.getImageData(0, 0, 320, 240).data) }
    } })
    for (const ctx of [direct, recorder.ctx]) { ctx.imageSmoothingQuality = quality; ctx.drawImage(image as unknown as CanvasImageSource, 0, 0, w, h) }
    const unit = recorder.finish(); if (unit.discriminator !== 1) throw new Error(unit.reason)
    await replayUnit(result, { ops: unit.ops, assets: new Map(unit.assets.map(a => [a.id, a])) }, backend)
    expect(Buffer.compare(pixels(direct), pixels(result)), `${w}x${h} ${quality}`).toBe(0)
  }
})
it.skipIf(process.env.S0_BACKEND !== 'napi')('allows exact partial-alpha snapshots on napi explicitly', async () => {
  const scratch = createCanvas(30, 30), s = scratch.getContext('2d')
  s.fillStyle = 'rgba(153,77,211,.31)'; s.fillRect(.3, .4, 23.2, 18.1)
  const direct = context(), backing = context(), result = context()
  const rec = createRecorder(backing, { scale: 1, partialAlphaSnapshots: 'allow' })
  for (const ctx of [direct, rec.ctx]) { ctx.fillStyle = '#41aa72'; ctx.fillRect(0, 0, 180, 100); ctx.drawImage(scratch as unknown as CanvasImageSource, 12, 13) }
  const unit = rec.finish(); expect(unit.discriminator, unit.discriminator === 3 ? unit.reason : '').toBe(1); if (unit.discriminator !== 1) throw new Error(unit.reason)
  await replayUnit(result, { ops: unit.ops, assets: new Map(unit.assets.map(a => [a.id, a])) }, backend)
  expect(Buffer.compare(pixels(direct), pixels(result))).toBe(0)
})
it('does not rebind an unchanged gradient around an already-built path', async () => {
  const unit = await compare(ctx => {
    const gradient = ctx.createLinearGradient(0, 0, 40, 0); gradient.addColorStop(0, 'red'); gradient.addColorStop(1, 'blue')
    ctx.fillStyle = gradient; ctx.beginPath(); ctx.rect(2.3, 3.7, 40.2, 30.1); ctx.fill(); ctx.stroke()
  })
  expect([...new OpReader(unit.ops)].filter(op => op.code === Opcode.STYLE_OBJ)).toHaveLength(1)
})
