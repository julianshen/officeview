/** 1d Task 2: the browser StreamSession against the real 1c server. */
import { afterAll, describe, expect, it } from 'bun:test'
import { pngChunkTypes } from '../../src/stream/raster'
import { StreamSession, type StreamNotice } from '../../src/stream/remote/session'
import { StreamSessionError } from '../../src/stream/remote/errors'
import { ORIGIN, corpus, disposePool, start } from './harness'

afterAll(() => disposePool())
const settle = (ms: number) => Bun.sleep(ms)
type Sock = WebSocket & { terminate?: () => void }

function factory() {
  const sockets: Sock[] = []
  const client = (s: { url: string }) => { const ws = new WebSocket(s.url, { headers: { Origin: ORIGIN } } as never) as Sock; sockets.push(ws); return ws }
  return { client, sockets }
}

describe('StreamSession against the 1c server', () => {
  it('opens a protected session and receives a critical-chunk raster', async () => {
    const h = await start(), f = factory(), notices: StreamNotice[] = []
    const s = await StreamSession.open({ kind: 'stream', url: h.url }, { client: f.client, onNotice: (n) => notices.push(n) })
    expect(s.manifest.recorder.projection).toBe('final-raster')
    expect(notices.map((n) => n.code)).toContain('forensic-mark-stage-2')
    const u = await s.unit(0, 1)
    if (u.kind !== 'raster') throw new Error('expected raster')
    expect(u.w).toBe(s.manifest.units.list[0]!.w)
    expect(pngChunkTypes(u.png).every((t) => ['IHDR', 'PLTE', 'IDAT', 'IEND', 'tRNS'].includes(t))).toBe(true)
    s.close(); h.stop()
  })
  it('a relaxed session yields command units with resolved assets', async () => {
    const h = await start({ doc: corpus('pypptx-test.pptx'), policy: { projection: 'relaxed' } }), f = factory()
    const s = await StreamSession.open({ kind: 'stream', url: h.url }, { client: f.client, faces: async () => [] })
    const u = await s.unit(0, 1)
    if (u.kind !== 'commands') throw new Error(`expected commands, got ${u.kind}`)
    expect(u.ops.length).toBeGreaterThan(0)
    expect(u.assets).toBeInstanceOf(Map)
    s.close(); h.stop()
  })
  it('RENEW mid-flight: renew-ack, in-flight unit resolves, later units use the new token', async () => {
    const h = await start({ tokenTtlSec: 6, renewLeadSec: 5 }), f = factory()
    const s = await StreamSession.open({ kind: 'stream', url: h.url }, { client: f.client })
    const first = s.unit(0, 1)
    const t0 = s.token
    await settle(1600)
    expect(s.token).not.toBe(t0)
    await first
    expect((await s.unit(0, 2)).kind).toBe('raster')
    s.close(); h.stop()
  })
  it('a terminal close (revoke 4005) rejects pending units with the typed code', async () => {
    const h = await start(), f = factory()
    let closedWith: StreamSessionError | undefined
    const s = await StreamSession.open({ kind: 'stream', url: h.url }, { client: f.client, onState: (st, e) => { if (st === 'closed') closedWith = e } })
    const p = s.unit(0, 2)
    h.stream.revoke(s.manifest.session.sid)
    const e = await p.then(() => undefined, (x) => x)
    expect(e).toBeInstanceOf(StreamSessionError)
    expect((e as StreamSessionError).code).toBe('revoked')
    await settle(20)
    expect(closedWith?.code).toBe('revoked')
    h.stop()
  })
  it('an abnormal drop resumes the same sid and re-issues the pending unit', async () => {
    const h = await start(), f = factory(), notices: StreamNotice[] = []
    const s = await StreamSession.open({ kind: 'stream', url: h.url }, { client: f.client, onNotice: (n) => notices.push(n) })
    const sid = s.manifest.session.sid
    await s.unit(0, 1)
    f.sockets[0]!.terminate!()
    const u = await s.unit(0, 2)
    expect(u.kind).toBe('raster')
    expect(f.sockets.length).toBe(2)
    expect(s.manifest.session.sid).toBe(sid)
    expect(notices.some((n) => n.code === 'resumed')).toBe(true)
    s.close(); h.stop()
  })
  it('a drop after the grace window ends the session auth-failed', async () => {
    const h = await start({ detachGraceMs: 100 }), f = factory()
    let closedWith: StreamSessionError | undefined
    const s = await StreamSession.open({ kind: 'stream', url: h.url }, { client: f.client, onState: (st, e) => { if (st === 'closed') closedWith = e } })
    f.sockets[0]!.terminate!()
    await settle(1500)
    expect(closedWith?.code).toBe('auth-failed')
    await expect(s.unit(0, 1)).rejects.toBeInstanceOf(StreamSessionError)
    h.stop()
  })
})
