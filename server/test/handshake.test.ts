import { afterAll, describe, expect, it } from 'bun:test'
import { loadImage } from '@napi-rs/canvas'
import { decodeAssets } from '../../src/stream/asset-wire'
import { decodeRasterBody, pngChunkTypes } from '../../src/stream/raster'
import { replayUnit } from '../../src/stream/replay'
import type { Asset } from '../../src/stream/ops'
import { DocumentHost } from '../src/render'
import { Client, start, corpus, disposePool } from './harness'

const host = new DocumentHost()
const deck = corpus('pypptx-test.pptx')
afterAll(() => disposePool())

async function rgba(png: Uint8Array, w: number, h: number) {
  const c = host.backend.createCanvas(w, h), ctx = c.getContext('2d') as unknown as CanvasRenderingContext2D
  ctx.drawImage(await loadImage(Buffer.from(png)) as never, 0, 0)
  return Buffer.from(ctx.getImageData(0, 0, w, h).data)
}

describe('§6.2 handshake', () => {
  it('protected default: MANIFEST → READY → UNIT_RASTER is the pool raster plus the session mark', async () => {
    const h = await start()
    const c = new Client(h.url)
    const m = await c.hello()
    expect(m.policy).toMatchObject({ protection: 'final-raster', contentPolicy: 'visible-only', contentPolicyId: 'visible-only@1', viewerMarkId: 'fixed' })
    expect(m.recorder.projection).toBe('final-raster')
    expect(m.recorder.requiredFaces).toEqual([])
    expect(m.notices.filter((n) => n.startsWith('forensic'))).toEqual([])
    expect(m.can).toMatchObject({ print: true, search: true, copy: false, textA11y: false, notes: false })
    expect(h.stream.audit.get(m.session.sid)).toMatchObject({ viewerId: 'alice', markSettings: { visible: 'fixed', forensic: false }, forensicLayers: [] })
    await c.ready()
    c.unit(10)
    const { unit, messages } = await c.unitMessages(10)
    expect(messages.map((x) => x.kind)).toEqual([3])
    expect(unit.discriminator).toBe(2)
    const body = decodeRasterBody(unit.body)
    expect([body.w, body.h]).toEqual([m.units.list[0]!.w, m.units.list[0]!.h])
    expect(new Set(pngChunkTypes(body.png))).toEqual(new Set(['IHDR', 'IDAT', 'IEND']))
    const { docKey } = await host.open('x', deck, 'visible-only')
    const direct = host.render(docKey, { unitIndex: 0, scale: 1, projection: 'final-raster', mark: { kind: 'fixed' } })
    if (direct.kind !== 'raster') throw new Error('raster')
    // The canonical cache entry is the pool raster; the stream carries it plus this session's Layer-1 mark.
    expect(Buffer.from(h.stream.canonicalOf(m.session.sid, 0, 1)!).equals(Buffer.from(direct.png))).toBe(true)
    const [got, want] = [await rgba(body.png, body.w, body.h), await rgba(direct.png, body.w, body.h)]
    let maxd = 0, changed = 0
    for (let k = 0; k < got.length; k++) { const d = Math.abs(got[k]! - want[k]!); if (d) changed++; maxd = Math.max(maxd, d) }
    expect(changed).toBeGreaterThan(0)
    expect(maxd).toBeLessThanOrEqual(4)
    c.close(); h.stop()
  }, 30_000)

  it('relaxed session: ASSET_DEFINE then UNIT_COMMANDS replaying to the marked raster', async () => {
    const h = await start({ policy: { projection: 'relaxed', visibleMark: 'viewer' }, viewerId: 'bob' })
    const c = new Client(h.url)
    const m = await c.hello()
    expect(m.recorder.projection).toBe('commands')
    expect(m.recorder.requiredFaces.length).toBe(20)
    expect(m.notices).toContain('relaxed-session')
    expect(h.stream.audit.get(m.session.sid)).toMatchObject({ relaxed: true })
    await c.ready()
    for (const scale of [1, 2]) {
      const reqId = 20 + scale
      c.unit(reqId, 0, scale)
      const { unit, messages } = await c.unitMessages(reqId)
      expect(messages.at(-1)!.kind).toBe(4)
      const known = new Map<number, Asset>()
      for (const x of messages.slice(0, -1)) { expect(x.kind).toBe(2); for (const a of await decodeAssets(x.payload, known)) known.set(a.id, a) }
      const { docKey } = await host.open('x', deck, 'visible-only')
      const direct = host.render(docKey, { unitIndex: 0, scale, projection: 'final-raster', mark: { kind: 'viewer', text: 'bob' } })
      if (direct.kind !== 'raster') throw new Error('raster')
      const cv = host.backend.createCanvas(direct.w, direct.h), ctx = cv.getContext('2d') as unknown as CanvasRenderingContext2D
      await replayUnit(ctx, { ops: unit.body, assets: known }, host.backend)
      expect(Buffer.from(ctx.getImageData(0, 0, direct.w, direct.h).data).equals(await rgba(direct.png, direct.w, direct.h)), `scale ${scale}`).toBe(true)
    }
    c.close(); h.stop()
  }, 30_000)

  it('READY reporting a missing face streams that unit as a raster with face-missing', async () => {
    const h = await start({ policy: { projection: 'relaxed' } })
    const c = new Client(h.url)
    const m = await c.hello()
    await c.ready({ missingFaces: m.recorder.requiredFaces.map((f) => f.id) })
    c.unit(30)
    const { unit } = await c.unitMessages(30)
    expect(unit.discriminator).toBe(2)
    const diag = await c.next((x) => 'text' in x && x.text.type === 'diag')
    expect((diag as { text: Record<string, unknown> }).text).toMatchObject({ reqId: 30, code: 'face-missing', unitIndex: 0 })
    c.close(); h.stop()
  }, 30_000)
})
