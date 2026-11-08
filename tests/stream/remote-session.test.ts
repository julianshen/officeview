/** 1d Task 2: StreamSession ordering and guard cases over a scripted socket. */
import { describe, expect, it, vi } from 'vitest'
import { gzipSync } from 'node:zlib'
import { encodeMessage, FRAGMENT_CAP, Reassembler } from '../../src/stream/frame'
import { encodeManifest, type Manifest } from '../../src/stream/manifest'
import { encodeRasterBody } from '../../src/stream/raster'
import { encodeUnit } from '../../src/stream/unit'
import { StreamSession } from '../../src/stream/remote/session'
import { StreamSessionError } from '../../src/stream/remote/errors'

const exampleManifest = (): Manifest => ({
  v: 1, docRef: 'a'.repeat(24), kind: 'pptx', name: 'deck.pptx',
  units: { count: 2, list: [{ w: 960, h: 540 }, { w: 960, h: 540 }] },
  layoutFingerprint: 'f'.repeat(64),
  session: { sid: '0123456789ab', token: 'x.y.z', expiresSec: 300, resume: 'session' },
  policy: { contentPolicy: 'visible-only', protection: 'final-raster', semantic: 'visible-only', contentPolicyId: 'visible-only@1', viewerMarkId: 'fixed' },
  can: { search: false, copy: false, print: true, textA11y: false, present: false, notes: false, links: false },
  recorder: { commands: 'ovsd-1', textMode: 'text', scalesBuckets: [1, 2], requiredFaces: [{ id: 'carlito-regular', hash: 'f6418f708baede9789daef5d', alias: 'Calibri' }], projection: 'final-raster' },
  limits: { maxUnitPx: 8192, maxAssetBytes: 8 * 1024 * 1024, wsFrameBytes: 1 << 20, tilePx: 512 },
  notices: ['forensic-mark-stage-2'],
})

const b64url = (o: object) => btoa(JSON.stringify(o)).replace(/\+/g, '-').replace(/\//g, '_').replace(/=+$/, '')
const tokenFor = (jti: string) => `h.${b64url({ jti })}.s`
const settle = (ms = 20) => new Promise((r) => setTimeout(r, ms))

class FakeSocket extends EventTarget {
  binaryType: BinaryType = 'blob'
  readyState = 0
  sent: Array<Record<string, any>> = []
  closes: Array<{ code?: number; reason?: string }> = []
  send(d: string) { this.sent.push(JSON.parse(d)) }
  close(code?: number, reason?: string) {
    this.closes.push({ code, reason })
    if (this.readyState === 3) return
    this.readyState = 3
    queueMicrotask(() => this.dispatchEvent(Object.assign(new Event('close'), { code: code ?? 1005, reason: reason ?? '' })))
  }
  open() { this.readyState = 1; this.dispatchEvent(new Event('open')) }
  deliver(data: ArrayBuffer | string) { this.dispatchEvent(Object.assign(new Event('message'), { data })) }
  frame(kind: number, reqId: number, payload: Uint8Array) {
    for (const f of encodeMessage({ ver: 1, kind: kind as never, reqId, payload, compress: (b) => new Uint8Array(gzipSync(b)) })) this.deliver(f.slice().buffer)
  }
  units() { return this.sent.filter((m) => m.type === 'unit') }
}

function pngHeader(w: number, h: number) {
  const b = new Uint8Array(33), v = new DataView(b.buffer)
  b.set([0x89, 0x50, 0x4e, 0x47, 0x0d, 0x0a, 0x1a, 0x0a]); v.setUint32(8, 13); b.set([73, 72, 68, 82], 12)
  v.setUint32(16, w); v.setUint32(20, h); b.set([8, 6, 0, 0, 0], 24)
  return b
}
const raster = (unitId: number, w = 2, h = 2) => encodeUnit({ unitId, discriminator: 2, body: encodeRasterBody(w, h, pngHeader(w, h)) })

async function opened(m: Partial<Manifest> = {}, extra: { faces?: () => Promise<string[]> } = {}) {
  const sock = new FakeSocket()
  const manifest: Manifest = { ...exampleManifest(), recorder: { ...exampleManifest().recorder, requiredFaces: [] }, session: { ...exampleManifest().session, token: tokenFor('j1') }, ...m }
  const p = StreamSession.open({ kind: 'stream', url: 'wss://x/stream' }, { client: () => sock as never, faces: extra.faces ?? (async () => []) })
  sock.open(); await settle()
  const hello = sock.sent.find((x) => x.type === 'hello')!
  sock.frame(1, hello.reqId, encodeManifest(manifest))
  const s = await p
  return { s, sock, manifest }
}

describe('StreamSession (scripted socket)', () => {
  it('sends HELLO, then READY with the manifest token jti', async () => {
    const { sock } = await opened()
    expect(sock.sent[0]).toMatchObject({ type: 'hello', client: 'officeview/1.0', recVersions: [1] })
    expect(sock.sent.find((x) => x.type === 'ready')).toMatchObject({ tokenJti: 'j1' })
  })
  it('a message over 1 MiB closes 4007 before reassembly and fails pending units', async () => {
    const { s, sock } = await opened()
    const push = vi.spyOn(Reassembler.prototype, 'push')
    const p = s.unit(0, 1)
    await settle()
    sock.deliver(new ArrayBuffer(FRAGMENT_CAP + 1))
    const e = await p.then(() => undefined, (x) => x)
    expect(push).not.toHaveBeenCalled()
    expect(sock.closes[0]?.code).toBe(4007)
    expect((e as StreamSessionError).code).toBe('oversized-frame')
    push.mockRestore()
  })
  it('a frame for a reqId the client never issued is dropped; the real one resolves', async () => {
    const { s, sock } = await opened()
    const p = s.unit(0, 1)
    await settle()
    sock.frame(3, 999, raster(0))
    const req = sock.units()[0]!
    sock.frame(3, req.reqId, raster(0))
    expect((await p).kind).toBe('raster')
  })
  it('an aborted unit releases its slot and its late frame is dropped', async () => {
    const { s, sock } = await opened()
    const ac = new AbortController()
    const p = s.unit(0, 1, ac.signal)
    await settle()
    ac.abort()
    await expect(p).rejects.toBeDefined()
    sock.frame(3, sock.units()[0]!.reqId, raster(0))
    await settle()
    expect(s.stats()).toMatchObject({ inFlight: 0, inFlightBytes: 0 })
  })
  it('close() during the handshake closes the socket once, rejects open, sends no READY', async () => {
    const sock = new FakeSocket(), ac = new AbortController()
    const p = StreamSession.open({ kind: 'stream', url: 'wss://x' }, { client: () => sock as never, signal: ac.signal })
    sock.open(); await settle()
    ac.abort(); ac.abort()
    const e = await p.then(() => undefined, (x) => x)
    expect((e as Error).name).toBe('AbortError')
    await settle()
    expect(sock.closes.length).toBe(1)
    expect(sock.sent.some((x) => x.type === 'ready')).toBe(false)
  })
  it('READY is sent at ~1.5 s with every unresolved face missing when hashing stalls', async () => {
    const faces = [{ id: 'carlito-regular', hash: 'f6418f708baede9789daef5d' }, { id: 'carlito-bold', hash: 'bb5d20f79b82599ec7298359' }]
    const sock = new FakeSocket()
    const manifest = { ...exampleManifest(), recorder: { ...exampleManifest().recorder, requiredFaces: faces }, session: { ...exampleManifest().session, token: tokenFor('j1') } }
    const t0 = Date.now()
    const p = StreamSession.open({ kind: 'stream', url: 'wss://x' }, { client: () => sock as never, faces: () => new Promise(() => {}) })
    sock.open(); await settle()
    sock.frame(1, sock.sent[0]!.reqId, encodeManifest(manifest))
    await p
    const ready = sock.sent.find((x) => x.type === 'ready')!
    expect(Date.now() - t0).toBeLessThan(1900)
    expect(ready.missingFaces).toEqual(['carlito-regular', 'carlito-bold'])
  })
  it('at most 8 unit requests are outstanding; the ninth goes out when one settles', async () => {
    const { s, sock } = await opened({ units: { count: 9, list: Array.from({ length: 9 }, () => ({ w: 10, h: 10 })) } })
    const ps = Array.from({ length: 9 }, (_, i) => s.unit(i, 1))
    await settle()
    expect(sock.units().length).toBe(8)
    sock.frame(3, sock.units()[0]!.reqId, raster(0))
    await ps[0]
    await settle()
    expect(sock.units().length).toBe(9)
  })
  it('the 96 MiB aggregate reservation admits one 64 MiB-raster unit at a time', async () => {
    const { s, sock } = await opened({ units: { count: 2, list: [{ w: 4096, h: 4096 }, { w: 4096, h: 4096 }] } })
    void s.unit(0, 1); void s.unit(1, 1)
    await settle()
    expect(sock.units().length).toBe(1)
    expect(s.stats().inFlightBytes).toBeLessThanOrEqual(96 * 1024 * 1024)
  })
  it('session-rotated is re-issued once under the current token; a second one rejects', async () => {
    const { s, sock } = await opened()
    const p = s.unit(0, 1)
    await settle()
    sock.deliver(JSON.stringify({ reqId: sock.units()[0]!.reqId, type: 'refused', code: 'session-rotated' }))
    await settle()
    expect(sock.units().length).toBe(2)
    sock.deliver(JSON.stringify({ reqId: sock.units()[1]!.reqId, type: 'refused', code: 'session-rotated' }))
    expect(((await p.then(() => undefined, (x) => x)) as StreamSessionError).code).toBe('session-rotated')
  })
  it('RENEW: renew-ack carries the new jti and later units use the new token', async () => {
    const { s, sock } = await opened()
    sock.deliver(JSON.stringify({ reqId: 0, type: 'renew', token: tokenFor('j2'), jti: 'j2', exp: 9e9 }))
    await settle()
    expect(sock.sent.find((x) => x.type === 'renew-ack')).toMatchObject({ jti: 'j2' })
    void s.unit(0, 1)
    await settle()
    expect(sock.units()[0]!.token).toBe(tokenFor('j2'))
  })
})
