/** Stage 3c Task 3: StreamSession.notes() / RemoteOfficeDoc.paintNotes() — kind-7 admission only for a live notes request, fail closed. */
import { describe, expect, it, vi } from 'vitest'
import { gzipSync } from 'node:zlib'
import { encodeMessage, Reassembler } from '../../src/stream/frame'
import { encodeManifest, type Manifest } from '../../src/stream/manifest'
import { encodeRasterBody } from '../../src/stream/raster'
import { encodeUnit } from '../../src/stream/unit'
import { notesBodyCap, notesGrant, StreamSession, type SessionOptions } from '../../src/stream/remote/session'
import { StreamSessionError } from '../../src/stream/remote/errors'
import { createRemoteDoc } from '../../src/stream/remote/remote-doc'
import type { CanvasBackend, CanvasLike } from '../../src/core/canvas-backend'

const MiB = 1024 * 1024
const b64url = (o: object) => btoa(JSON.stringify(o)).replace(/\+/g, '-').replace(/\//g, '_').replace(/=+$/, '')
const tokenFor = (jti: string) => `h.${b64url({ jti })}.s`
const settle = (ms = 20) => new Promise((r) => setTimeout(r, ms))

const exampleManifest = (): Manifest => ({
  v: 1, docRef: 'a'.repeat(24), kind: 'pptx', name: 'deck.pptx',
  units: { count: 3, list: [{ w: 960, h: 540 }, { w: 960, h: 540 }, { w: 960, h: 540, hidden: true }] },
  layoutFingerprint: 'f'.repeat(64),
  session: { sid: '0123456789ab', token: tokenFor('j1'), expiresSec: 300, resume: 'session' },
  policy: { contentPolicy: 'visible-only', protection: 'final-raster', semantic: 'visible-only', contentPolicyId: 'visible-only@1', viewerMarkId: 'fixed' },
  can: { search: true, copy: false, print: true, textA11y: false, present: true, notes: true, links: false },
  caps: { copy: null, a11y: 'off', print: { maxPages: null, maxSessionPages: null } },
  recorder: { commands: 'ovsd-1', textMode: 'text', scalesBuckets: [1, 2], requiredFaces: [], projection: 'final-raster' },
  limits: { maxUnitPx: 8192, maxAssetBytes: 8 * 1024 * 1024, wsFrameBytes: 1 << 20, tilePx: 512 },
  notices: [],
})

class FakeSocket extends EventTarget {
  binaryType: BinaryType = 'blob'
  readyState = 0
  sent: Array<Record<string, any>> = []
  send(d: string) { this.sent.push(JSON.parse(d)) }
  close(code?: number, reason?: string) {
    if (this.readyState === 3) return
    this.readyState = 3
    queueMicrotask(() => this.dispatchEvent(Object.assign(new Event('close'), { code: code ?? 1005, reason: reason ?? '' })))
  }
  open() { this.readyState = 1; this.dispatchEvent(new Event('open')) }
  deliver(data: ArrayBuffer | string) { this.dispatchEvent(Object.assign(new Event('message'), { data })) }
  text(m: object) { this.deliver(JSON.stringify(m)) }
  frame(kind: number, reqId: number, payload: Uint8Array) {
    for (const f of encodeMessage({ ver: 1, kind: kind as never, reqId, payload, compress: (b) => new Uint8Array(gzipSync(b)) })) this.deliver(f.slice().buffer)
  }
  of(type: string) { return this.sent.filter((m) => m.type === type) }
  drop(code: number) { this.readyState = 3; this.dispatchEvent(Object.assign(new Event('close'), { code, reason: '' })) }
}

function pngHeader(w: number, h: number) {
  const b = new Uint8Array(33), v = new DataView(b.buffer)
  b.set([0x89, 0x50, 0x4e, 0x47, 0x0d, 0x0a, 0x1a, 0x0a]); v.setUint32(8, 13); b.set([73, 72, 68, 82], 12)
  v.setUint32(16, w); v.setUint32(20, h); b.set([8, 6, 0, 0, 0], 24)
  return b
}
/** A NOTES_RASTER body: `u32 LE unitIndex` + the raster body. */
function notesBody(unitIndex: number, w: number, h: number) {
  const r = encodeRasterBody(w, h, pngHeader(w, h)), out = new Uint8Array(4 + r.length)
  new DataView(out.buffer).setUint32(0, unitIndex, true); out.set(r, 4)
  return out
}

async function opened(o: Partial<SessionOptions> = {}, m: Partial<Manifest> = {}) {
  const socks: FakeSocket[] = []
  const manifest: Manifest = { ...exampleManifest(), ...m }
  const client = () => { const s = new FakeSocket(); socks.push(s); return s as never }
  const p = StreamSession.open({ kind: 'stream', url: 'wss://x/stream' }, { client, faces: async () => [], ...o })
  const sock = socks[0]!
  sock.open(); await settle()
  sock.frame(1, sock.of('hello')[0]!.reqId, encodeManifest(manifest))
  const s = await p
  return { s, sock, socks, manifest }
}
const failure = (p: Promise<unknown>) => p.then(() => { throw new Error('expected a rejection') }, (e) => e)

/** A canvas + backend pair recording what was painted. */
function surface() {
  const draws: unknown[] = []
  const ctx = { setTransform() {}, clearRect() {}, drawImage: (img: unknown) => { draws.push(img) } }
  const canvas = { width: 7, height: 7, getContext: () => ctx } as unknown as CanvasLike
  const decoded: Array<{ w: number; h: number }> = []
  const backend = {
    name: 'fake',
    createCanvas: () => canvas,
    decodeImage: async (png: Uint8Array) => {
      const v = new DataView(png.buffer, png.byteOffset, png.byteLength), img = { width: v.getUint32(16), height: v.getUint32(20) }
      decoded.push({ w: img.width, h: img.height })
      return img
    },
  } as unknown as CanvasBackend
  return { canvas, backend, draws, decoded }
}

/** Starts a paintNotes and returns its promise (observed) and the sent request. */
async function painting(s: StreamSession, sock: FakeSocket, unit = 0, dpr = 2, signal?: AbortSignal) {
  const sf = surface()
  const doc = createRemoteDoc(s, { backend: sf.backend })
  const p = doc.paintNotes(unit, dpr, sf.canvas, signal)
  void p.catch(() => {})
  await settle()
  const req = sock.of('notes').at(-1)
  return { p, req, reqId: req?.reqId as number, doc, ...sf }
}

/** The session still serves units. */
async function usable(s: StreamSession, sock: FakeSocket) {
  const u = s.unit(0, 1)
  await settle()
  const req = sock.of('unit').at(-1)!
  sock.frame(3, req.reqId, encodeUnit({ unitId: 0, discriminator: 2, body: encodeRasterBody(2, 2, pngHeader(2, 2)) }))
  const lease = await u
  lease.release()
  expect(lease.kind).toBe('raster')
}

describe('StreamSession.notes / RemoteOfficeDoc.paintNotes (scripted socket)', () => {
  it('happy path: one notes request, notes-start, one kind-7 frame → canvas painted at the decoded size, grant released', async () => {
    const { s, sock } = await opened()
    const { p, req, reqId, canvas, draws, decoded } = await painting(s, sock, 1, 2)
    expect(req).toEqual({ reqId, type: 'notes', token: tokenFor('j1'), unit: 1, scale: 2 })
    expect(s.stats()).toMatchObject({ inFlight: 1, inFlightBytes: notesGrant(2) })
    sock.text({ reqId, type: 'notes-start', w: 1440, h: 96, scale: 2 })
    sock.frame(7, reqId, notesBody(1, 1440, 96))
    await expect(p).resolves.toBeUndefined()
    expect([canvas.width, canvas.height]).toEqual([1440, 96])
    expect(decoded).toEqual([{ w: 1440, h: 96 }])
    expect(draws.length).toBe(1)
    await settle()
    expect(s.stats()).toMatchObject({ inFlight: 0, inFlightBytes: 0, reassemblyBytes: 0 })
  })

  it('the grant is encoded cap + the scale\'s decoded body, well under the session aggregate', () => {
    expect(notesBodyCap(1)).toBe(16 + 720 * 8192 * 4)
    expect(notesBodyCap(2)).toBe(16 + 1440 * 8192 * 4)
    expect(notesGrant(2)).toBe(8 * MiB + notesBodyCap(2))
    expect(notesGrant(2)).toBeLessThan(96 * MiB)
  })

  it('the scale follows dpr: the smallest offered bucket ≥ dpr, else the largest', async () => {
    const { s, sock } = await opened()
    for (const [dpr, scale] of [[1, 1], [0.5, 1], [1.5, 2], [2, 2], [3, 2]] as const) {
      const { req, reqId } = await painting(s, sock, 0, dpr)
      expect(req!.scale).toBe(scale)
      sock.text({ reqId, type: 'error', code: 'rate-limited' })
    }
  })

  it('can.notes false → policy-refused before any send', async () => {
    const { s, sock } = await opened({}, { can: { ...exampleManifest().can, notes: false } })
    const { p } = await painting(s, sock)
    expect((await failure(p) as StreamSessionError).code).toBe('policy-refused')
    expect((await failure(s.notes(0, 1)) as StreamSessionError).code).toBe('policy-refused')
    expect(sock.of('notes')).toEqual([])
    expect(s.stats()).toMatchObject({ inFlight: 0, inFlightBytes: 0 })
  })

  it('a unit out of range or hidden, or a scale the server does not take, refuses without sending', async () => {
    const { s, sock } = await opened()
    expect((await failure(s.notes(3, 1)) as StreamSessionError).code).toBe('invalid-unit')
    expect((await failure(s.notes(-1, 1)) as StreamSessionError).code).toBe('invalid-unit')
    expect((await failure(s.notes(0.5, 1)) as StreamSessionError).code).toBe('invalid-unit')
    expect((await failure(s.notes(2, 1)) as StreamSessionError).code).toBe('invalid-unit')
    expect((await failure(s.notes(0, 3 as never)) as StreamSessionError).code).toBe('policy-refused')
    expect(sock.of('notes')).toEqual([])
  })

  it('an unsolicited kind-7 frame (no live notes request) is rejected: nothing paints, the session stays usable', async () => {
    const { s, sock } = await opened()
    const push = vi.spyOn(Reassembler.prototype, 'push')
    sock.frame(7, 99, notesBody(0, 720, 48))
    await settle()
    const results = (await Promise.all(push.mock.results.map((r) => r.value))).filter(Boolean)
    expect(results).toEqual([{ ok: false, reqId: 99, reason: 'unexpected-reqId' }])
    push.mockRestore()
    // A notes request whose reqId is not yet sent does not admit kind 7 either; a live one is unaffected.
    const { p, reqId, draws } = await painting(s, sock, 0, 1)
    sock.frame(7, reqId + 1, notesBody(0, 720, 48))
    sock.text({ reqId, type: 'notes-start', w: 720, h: 48, scale: 1 })
    sock.frame(7, reqId, notesBody(0, 720, 48))
    await p
    expect(draws.length).toBe(1)
    await usable(s, sock)
  })

  it('a kind-7 frame for another unit than asked → policy-refused, nothing paints', async () => {
    const { s, sock } = await opened()
    const { p, reqId, draws, canvas } = await painting(s, sock, 0, 1)
    sock.text({ reqId, type: 'notes-start', w: 720, h: 48, scale: 1 })
    sock.frame(7, reqId, notesBody(1, 720, 48))
    expect((await failure(p) as StreamSessionError).code).toBe('policy-refused')
    expect(draws).toEqual([])
    expect([canvas.width, canvas.height]).toEqual([7, 7])
    await settle()
    expect(s.stats()).toMatchObject({ inFlight: 0, inFlightBytes: 0 })
    await usable(s, sock)
  })

  const hostile: Array<[string, (sock: FakeSocket, reqId: number) => void]> = [
    ['a frame before notes-start', (k, r) => k.frame(7, r, notesBody(0, 720, 48))],
    ['notes-start at another scale', (k, r) => k.text({ reqId: r, type: 'notes-start', w: 1440, h: 48, scale: 2 })],
    ['notes-start wider than 720 × scale', (k, r) => k.text({ reqId: r, type: 'notes-start', w: 721, h: 48, scale: 1 })],
    ['notes-start narrower than 720 × scale', (k, r) => k.text({ reqId: r, type: 'notes-start', w: 719, h: 48, scale: 1 })],
    ['notes-start taller than 8,192 px', (k, r) => k.text({ reqId: r, type: 'notes-start', w: 720, h: 8193, scale: 1 })],
    ['notes-start with h 0', (k, r) => k.text({ reqId: r, type: 'notes-start', w: 720, h: 0, scale: 1 })],
    ['notes-start with a fractional h', (k, r) => k.text({ reqId: r, type: 'notes-start', w: 720, h: 48.5, scale: 1 })],
    ['a second notes-start', (k, r) => { k.text({ reqId: r, type: 'notes-start', w: 720, h: 48, scale: 1 }); k.text({ reqId: r, type: 'notes-start', w: 720, h: 48, scale: 1 }) }],
    ['a raster whose size is not the notes-start one', (k, r) => { k.text({ reqId: r, type: 'notes-start', w: 720, h: 48, scale: 1 }); k.frame(7, r, notesBody(0, 720, 49)) }],
    ['a malformed body', (k, r) => { k.text({ reqId: r, type: 'notes-start', w: 720, h: 48, scale: 1 }); k.frame(7, r, notesBody(0, 720, 48).subarray(0, 20)) }],
    ['a body shorter than its unit index', (k, r) => { k.text({ reqId: r, type: 'notes-start', w: 720, h: 48, scale: 1 }); k.frame(7, r, new Uint8Array([0, 0, 0, 0])) }],
    ['a unit-kind frame on the request', (k, r) => { k.text({ reqId: r, type: 'notes-start', w: 720, h: 48, scale: 1 }); k.frame(3, r, encodeUnit({ unitId: 0, discriminator: 2, body: encodeRasterBody(2, 2, pngHeader(2, 2)) })) }],
    ['an unknown message type', (k, r) => k.text({ reqId: r, type: 'notes-done' })],
    ['notes-start over 48 KiB', (k, r) => k.text({ reqId: r, type: 'notes-start', w: 720, h: 48, scale: 1, pad: 'x'.repeat(50_000) })],
  ]
  for (const [name, act] of hostile) {
    it(`hostile: ${name} → policy-refused, nothing paints, the session stays usable`, async () => {
      const { s, sock } = await opened()
      const { p, reqId, draws } = await painting(s, sock, 0, 1)
      act(sock, reqId)
      const e = await failure(p) as StreamSessionError
      expect(e).toBeInstanceOf(StreamSessionError)
      expect(e.code).toBe('policy-refused')
      expect(draws).toEqual([])
      await settle()
      expect(s.stats()).toMatchObject({ inFlight: 0, inFlightBytes: 0, reassemblyBytes: 0 })
      await usable(s, sock)
    })
  }

  it('a kind-7 body over the scale\'s decoded cap is a decoded-cap rejection; the grant is released', async () => {
    const { s, sock } = await opened()
    const { p, reqId, draws } = await painting(s, sock, 0, 1)
    sock.text({ reqId, type: 'notes-start', w: 720, h: 48, scale: 1 })
    const push = vi.spyOn(Reassembler.prototype, 'push')
    sock.frame(7, reqId, new Uint8Array(notesBodyCap(1) + 1))
    expect((await failure(p) as StreamSessionError).code).toBe('policy-refused')
    const results = (await Promise.all(push.mock.results.map((r) => r.value))).filter(Boolean)
    expect(results).toEqual([{ ok: false, reqId, reason: 'decoded-cap' }])
    push.mockRestore()
    expect(draws).toEqual([])
    await settle()
    expect(s.stats()).toMatchObject({ inFlight: 0, inFlightBytes: 0, reassemblyBytes: 0 })
  })

  it('refusals reject typed; an unknown code is policy-refused; the session stays usable', async () => {
    const { s, sock } = await opened()
    for (const [type, code, want] of [
      ['error', 'policy-refused', 'policy-refused'], ['error', 'invalid-unit', 'invalid-unit'], ['error', 'rate-limited', 'rate-limited'],
      ['refused', 'rate-limited', 'rate-limited'], ['error', 'no-such-code', 'policy-refused'],
    ] as const) {
      const { p, reqId, draws } = await painting(s, sock, 0, 1)
      sock.text({ reqId, type, code })
      const e = await failure(p) as StreamSessionError
      expect(e).toBeInstanceOf(StreamSessionError)
      expect(e.code).toBe(want)
      expect(draws).toEqual([])
    }
    await settle()
    expect(s.stats()).toMatchObject({ inFlight: 0, inFlightBytes: 0 })
    await usable(s, sock)
  })

  it('session-rotated (stale jti) is re-issued once under a new reqId with the current token, then rejects', async () => {
    const { s, sock } = await opened()
    const { p, reqId } = await painting(s, sock, 0, 1)
    sock.text({ type: 'renew', token: tokenFor('j2'), jti: 'j2' })
    sock.text({ reqId, type: 'error', code: 'session-rotated' })
    await settle()
    const again = sock.of('notes')
    expect(again.length).toBe(2)
    expect(again[1]).toEqual({ reqId: again[1]!.reqId, type: 'notes', token: tokenFor('j2'), unit: 0, scale: 1 })
    expect(again[1]!.reqId).not.toBe(reqId)
    // The retired reqId no longer admits kind 7.
    sock.text({ reqId, type: 'notes-start', w: 720, h: 48, scale: 1 })
    sock.text({ reqId: again[1]!.reqId, type: 'error', code: 'session-rotated' })
    expect((await failure(p) as StreamSessionError).code).toBe('session-rotated')
    await settle()
    expect(sock.of('notes').length).toBe(2)
    expect(s.stats()).toMatchObject({ inFlight: 0, inFlightBytes: 0 })
  })

  it('a retried request paints from the new reqId', async () => {
    const { s, sock } = await opened()
    const { p, reqId, draws } = await painting(s, sock, 0, 1)
    sock.text({ reqId, type: 'error', code: 'session-rotated' })
    await settle()
    const r2 = sock.of('notes')[1]!.reqId
    sock.text({ reqId: r2, type: 'notes-start', w: 720, h: 48, scale: 1 })
    sock.frame(7, r2, notesBody(0, 720, 48))
    await p
    expect(draws.length).toBe(1)
  })

  it('abort mid-flight → rejects with the abort reason, nothing paints, the reservation is released; late answers are dropped', async () => {
    const { s, sock } = await opened()
    const ac = new AbortController()
    const { p, reqId, draws } = await painting(s, sock, 0, 1, ac.signal)
    sock.text({ reqId, type: 'notes-start', w: 720, h: 48, scale: 1 })
    await settle()
    ac.abort()
    expect(((await failure(p)) as Error).name).toBe('AbortError')
    await settle()
    expect(s.stats()).toMatchObject({ inFlight: 0, inFlightBytes: 0, reassemblyBytes: 0 })
    // There is no notes-cancel: the late frame arrives and is dropped silently.
    sock.frame(7, reqId, notesBody(0, 720, 48))
    await settle()
    expect(draws).toEqual([])
    await usable(s, sock)
  })

  it('abort after the frame but before the paint → nothing paints', async () => {
    const { s, sock } = await opened()
    const ac = new AbortController()
    const { p, reqId, draws, backend } = await painting(s, sock, 0, 1, ac.signal)
    const decode = backend.decodeImage
    backend.decodeImage = async (png) => { ac.abort(); return decode(png) }
    sock.text({ reqId, type: 'notes-start', w: 720, h: 48, scale: 1 })
    sock.frame(7, reqId, notesBody(0, 720, 48))
    await p.catch(() => {})
    expect(draws).toEqual([])
    await settle()
    expect(s.stats()).toMatchObject({ inFlight: 0, inFlightBytes: 0 })
  })

  it('an already-aborted signal rejects without sending', async () => {
    const { s, sock } = await opened()
    const ac = new AbortController(); ac.abort()
    expect(((await failure(s.notes(0, 1, ac.signal))) as Error).name).toBe('AbortError')
    expect(sock.of('notes')).toEqual([])
  })

  it('a terminal end (revoked) rejects a pending request with the session failure; nothing paints', async () => {
    const { s, sock } = await opened()
    const { p, reqId, draws } = await painting(s, sock, 0, 1)
    sock.text({ reqId, type: 'notes-start', w: 720, h: 48, scale: 1 })
    await settle()
    sock.drop(4005)
    expect((await failure(p) as StreamSessionError).code).toBe('revoked')
    expect(draws).toEqual([])
    expect((await failure(s.notes(0, 1)) as StreamSessionError).code).toBe('revoked')
  })

  it('doc.ended: after a close(), pending notes reject and later calls refuse', async () => {
    const { s, sock } = await opened()
    const { p, doc, draws } = await painting(s, sock, 0, 1)
    doc.close()
    expect(((await failure(p)) as Error).name).toBe('AbortError')
    expect(draws).toEqual([])
    const sf = surface()
    expect((await failure(doc.paintNotes(0, 1, sf.canvas)) as StreamSessionError).code).toBe('closed')
  })

  it('a socket drop with a sent request rejects session-rotated; it is never re-issued after the resume', async () => {
    const { s, sock, socks, manifest } = await opened()
    const { p, reqId } = await painting(s, sock, 0, 1)
    sock.text({ reqId, type: 'notes-start', w: 720, h: 48, scale: 1 })
    await settle()
    sock.drop(1006)
    expect((await failure(p) as StreamSessionError).code).toBe('session-rotated')
    await settle(300)
    const next = socks[1]!
    next.open(); await settle()
    next.frame(1, next.of('hello')[0]!.reqId, encodeManifest(manifest))
    await settle()
    expect(next.of('notes')).toEqual([])
    expect(s.stats()).toMatchObject({ inFlight: 0, inFlightBytes: 0 })
  })

  it('a request made while resuming goes out after the resume; a resume that turns can.notes off refuses it and later calls client-side', async () => {
    const { s, sock, socks, manifest } = await opened()
    sock.drop(1006)
    await settle()
    const waiting = s.notes(0, 1)
    void waiting.catch(() => {})
    await settle(300)
    const next = socks[1]!
    next.open(); await settle()
    next.frame(1, next.of('hello')[0]!.reqId, encodeManifest({ ...manifest, can: { ...manifest.can, notes: false } }))
    await settle()
    expect((await failure(waiting) as StreamSessionError).code).toBe('policy-refused')
    expect(next.of('notes')).toEqual([])
    const sf = surface()
    const doc = createRemoteDoc(s, { backend: sf.backend })
    expect((await failure(doc.paintNotes(0, 1, sf.canvas)) as StreamSessionError).code).toBe('policy-refused')
    expect(next.of('notes')).toEqual([])
    await settle()
    expect(s.stats()).toMatchObject({ inFlight: 0, inFlightBytes: 0 })
  })

  it('a request made while resuming is sent on the new socket when notes stay on', async () => {
    const { s, sock, socks, manifest } = await opened()
    sock.drop(1006)
    await settle()
    const waiting = s.notes(1, 2)
    await settle(300)
    const next = socks[1]!
    next.open(); await settle()
    next.frame(1, next.of('hello')[0]!.reqId, encodeManifest({ ...manifest, session: { ...manifest.session, token: tokenFor('j9') } }))
    await settle()
    const req = next.of('notes')[0]!
    expect(req).toEqual({ reqId: req.reqId, type: 'notes', token: tokenFor('j9'), unit: 1, scale: 2 })
    next.text({ reqId: req.reqId, type: 'notes-start', w: 1440, h: 48, scale: 2 })
    next.frame(7, req.reqId, notesBody(1, 1440, 48))
    const lease = await waiting
    expect([lease.w, lease.h]).toEqual([1440, 48])
    expect(s.stats()).toMatchObject({ inFlight: 1, inFlightBytes: notesGrant(2) })
    lease.release()
    await settle()
    expect(s.stats()).toMatchObject({ inFlight: 0, inFlightBytes: 0 })
  })

  it('a session without notes support refuses paintNotes policy-refused', async () => {
    const sf = surface()
    const m = exampleManifest()
    const doc = createRemoteDoc({ manifest: m, unit: async () => { throw new Error('x') }, text: async () => {}, close() {} }, { backend: sf.backend })
    expect((await failure(doc.paintNotes(0, 1, sf.canvas)) as StreamSessionError).code).toBe('policy-refused')
  })
})
