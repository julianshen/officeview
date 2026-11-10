import { afterAll, describe, expect, it } from 'bun:test'
import { OpReader, Opcode } from '../../src/stream/ops'
import { loadImage } from '@napi-rs/canvas'
import { decodeAssets } from '../../src/stream/asset-wire'
import { decodeRasterBody } from '../../src/stream/raster'
import { replayUnit } from '../../src/stream/replay'
import type { Asset } from '../../src/stream/ops'
import { DocumentHost } from '../src/render'
import { augmentedFixtures } from './fixtures'
import { Client, ORIGIN, start, disposePool } from './harness'

afterAll(() => disposePool())
const refusal = async (c: Client, reqId: number) => (await c.next((m) => 'text' in m && m.text.reqId === reqId && m.text.type === 'refused') as { text: Record<string, unknown> }).text.code

describe('Review Focus', () => {
  it('ops before READY and a second READY are refused; the session stays open', async () => {
    const h = await start(), c = new Client(h.url); await c.hello()
    c.unit(5)
    expect(await refusal(c, 5)).toBe('policy-refused')
    await c.ready()
    c.send({ reqId: 6, type: 'ready', tokenJti: c.jti() })
    expect(await refusal(c, 6)).toBe('policy-refused')
    c.unit(7)
    expect((await c.unitMessages(7)).unit.discriminator).toBe(2)
    expect(c.msgs.filter((m) => 'binary' in m && m.binary.reqId === 5)).toEqual([])
    c.close(); h.stop()
  }, 30_000)

  it('bad unit requests are refused before any pool job', async () => {
    const h = await start(), c = new Client(h.url); const m = await c.hello(); await c.ready()
    const pagesBefore = h.stream.stats().pool.workers.reduce((n, w) => n + w.pages, 0)
    const bad: Array<[number, unknown, unknown]> = [[51, m.units.count, 1], [52, -1, 1], [53, 0.5, 1], [54, 0, 3], [55, 0, Number.NaN], [56, '0', 1], [57, 0, 1.5]]
    for (const [reqId, unitIndex, scale] of bad) c.send({ reqId, type: 'unit', token: m.session.token, unitIndex, scale })
    for (const [reqId] of bad) expect(await refusal(c, reqId)).toBe('policy-refused')
    c.send({ reqId: 0, type: 'unit', token: m.session.token, unitIndex: 0, scale: 1 })
    expect(await refusal(c, 0)).toBe('policy-refused')
    expect(h.stream.stats().pool.workers.reduce((n, w) => n + w.pages, 0)).toBe(pagesBefore)
    c.close(); h.stop()
  }, 30_000)

  it('client→server binary frames are dropped (before HELLO they close 4000)', async () => {
    const h = await start()
    const pre = new Client(h.url); await pre.opened(); pre.ws.send(new Uint8Array([0x4f, 0x56, 0x53, 0x44]))
    expect((await pre.closed).code).toBe(4000)
    const c = new Client(h.url); await c.hello(); await c.ready()
    c.ws.send(new Uint8Array(1024))
    c.unit(8)
    expect((await c.unitMessages(8)).unit.discriminator).toBe(2)
    c.close(); h.stop()
  }, 30_000)

  it('a token from another session or document is refused', async () => {
    const h = await start(), a = new Client(h.url), b = new Client(h.url)
    await a.hello(); await a.ready(); await b.hello(); await b.ready()
    a.unit(9, 0, 1, b.manifest!.session.token)
    expect(await refusal(a, 9)).toBe('policy-refused')
    a.unit(10, 0, 1, 'garbage')
    expect(await refusal(a, 10)).toBe('policy-refused')
    a.close(); b.close(); h.stop()
  }, 30_000)
})

describe('HTTP surface (§6.1: no document route)', () => {
  it('non-upgrade requests are 404 and unlisted origins 403', async () => {
    const h = await start()
    const http = h.url.replace('ws://', 'http://').replace('/stream', '')
    for (const p of ['/', '/stream', '/doc', '/corpus/pypptx-test.pptx']) expect((await fetch(http + p)).status).toBe(404)
    const evil = await fetch(`${http}/stream`, { headers: { upgrade: 'websocket', connection: 'Upgrade', origin: 'http://evil.example', 'sec-websocket-version': '13', 'sec-websocket-key': 'dGhlIHNhbXBsZSBub25jZQ==' } })
    expect(evil.status).toBe(403)
    const ok = await fetch(`${http}/stream`, { headers: { origin: ORIGIN } })
    expect(ok.status).toBe(404)
    h.stop()
  })
})

describe('T2 over the wire: 13 hidden constructs never transmitted', async () => {
  const fixtures = await augmentedFixtures()
  const enc = new TextEncoder()
  const contains = (hay: Uint8Array, needle: Uint8Array) => { outer: for (let i = 0; i + needle.length <= hay.length; i++) { for (let j = 0; j < needle.length; j++) if (hay[i + j] !== needle[j]) continue outer; return true } return false }
  const host = new DocumentHost()
  for (const f of fixtures) it(`${f.id}`, async () => {
    // The hidden source really is in the input (non-vacuous).
    expect(contains(f.withHidden, enc.encode(f.marker)) || (await (await import('jszip')).default.loadAsync(f.withHidden).then(async (z) => (await Promise.all(Object.values(z.files).map((e) => e.async('string')))).join(''))).includes(f.marker)).toBe(true)
    const pixels: Record<string, Buffer[]> = { relaxed: [], protected: [], marked: [] }
    const texts: string[] = []
    // `protected` (forensic off) is pixel-compared with the relaxed replay; `marked`
    // is the production default: leak-checked, and within the Layer-1 shift of it.
    const runs = { relaxed: { projection: 'relaxed' }, protected: { projection: 'protected', forensicMark: false }, marked: { projection: 'protected' } } as const
    for (const projection of ['relaxed', 'protected', 'marked'] as const) {
      const h = await start({ doc: f.withHidden, policy: runs[projection] }), c = new Client(h.url)
      const m = await c.hello(); await c.ready()
      expect(contains(enc.encode(JSON.stringify(m)), enc.encode(f.marker))).toBe(false)
      for (let u = 0; u < m.units.count; u++) {
        c.unit(100 + u, u, 1)
        const { messages, unit } = await c.unitMessages(100 + u)
        for (const x of messages) expect(contains(x.payload, enc.encode(f.marker)), `${f.id} ${projection} unit ${u} kind ${x.kind}`).toBe(false)
        const size = m.units.list[u]!
        const cv = host.backend.createCanvas(size.w, size.h), ctx = cv.getContext('2d') as unknown as CanvasRenderingContext2D
        if (unit.discriminator === 1) {
          for (const op of new OpReader(unit.body)) if (op.code === Opcode.GLYPH_RUN) texts.push(op.text)
          const known = new Map<number, Asset>()
          for (const x of messages.slice(0, -1)) for (const a of await decodeAssets(x.payload, known)) known.set(a.id, a)
          await replayUnit(ctx, { ops: unit.body, assets: known }, host.backend)
        } else {
          ctx.drawImage(await loadImage(Buffer.from(decodeRasterBody(unit.body).png)) as never, 0, 0)
        }
        pixels[projection]!.push(Buffer.from(ctx.getImageData(0, 0, size.w, size.h).data))
      }
      c.close(); h.stop()
    }
    // Visible evidence: the relaxed stream draws the control text, and the
    // protected rasters are exactly those pixels (the control is visible there too).
    const all = texts.join('|') + texts.join('')
    expect(all.includes(f.control), `${f.id}: control text drawn`).toBe(true)
    expect(all.includes(f.marker)).toBe(false)
    expect(pixels.protected!.length).toBe(pixels.relaxed!.length)
    pixels.protected!.forEach((p, i) => expect(p.equals(pixels.relaxed![i]!), `${f.id} unit ${i} protected == relaxed`).toBe(true))
    // amplitude 2 + pedestal 2 bound the per-channel Layer-1 shift.
    pixels.marked!.forEach((p, i) => { let d = 0; for (let k = 0; k < p.length; k++) d = Math.max(d, Math.abs(p[k]! - pixels.protected![i]![k]!)); expect(d, `${f.id} unit ${i} marked within 4 of canonical`).toBeLessThanOrEqual(4) })
  }, 60_000)
})
