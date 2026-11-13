/** Stage 3b Task 5: StreamSession.print() — every page checked against the job, fail closed. */
import { describe, expect, it, vi } from 'vitest'
import { gzipSync } from 'node:zlib'
import { encodeMessage, Reassembler } from '../../src/stream/frame'
import { encodeManifest, type Manifest } from '../../src/stream/manifest'
import { encodeRasterBody } from '../../src/stream/raster'
import { encodeUnit } from '../../src/stream/unit'
import { StreamSession, type SessionOptions } from '../../src/stream/remote/session'
import { StreamSessionError } from '../../src/stream/remote/errors'
import { PRINT_JOB_MAX_BLOB_BYTES, printSizeOk } from '../../src/stream/remote/print-wire'
import { createRemoteDoc } from '../../src/stream/remote/remote-doc'

const b64url = (o: object) => btoa(JSON.stringify(o)).replace(/\+/g, '-').replace(/\//g, '_').replace(/=+$/, '')
const tokenFor = (jti: string) => `h.${b64url({ jti })}.s`
const settle = (ms = 20) => new Promise((r) => setTimeout(r, ms))

const exampleManifest = (): Manifest => ({
  v: 1, docRef: 'a'.repeat(24), kind: 'pptx', name: 'deck.pptx',
  units: { count: 3, list: [{ w: 960, h: 540 }, { w: 960, h: 540 }, { w: 960, h: 540 }] },
  layoutFingerprint: 'f'.repeat(64),
  session: { sid: '0123456789ab', token: tokenFor('j1'), expiresSec: 300, resume: 'session' },
  policy: { contentPolicy: 'visible-only', protection: 'final-raster', semantic: 'visible-only', contentPolicyId: 'visible-only@1', viewerMarkId: 'fixed' },
  can: { search: true, copy: false, print: true, textA11y: false, present: false, notes: false, links: false },
  caps: { copy: null, a11y: 'off', print: { maxPages: null, maxSessionPages: null } },
  recorder: { commands: 'ovsd-1', textMode: 'text', scalesBuckets: [1, 2], requiredFaces: [], projection: 'final-raster' },
  limits: { maxUnitPx: 8192, maxAssetBytes: 8 * 1024 * 1024, wsFrameBytes: 1 << 20, tilePx: 512 },
  notices: [],
})

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
  text(m: object) { this.deliver(JSON.stringify(m)) }
  frame(kind: number, reqId: number, payload: Uint8Array) {
    for (const f of encodeMessage({ ver: 1, kind: kind as never, reqId, payload, compress: (b) => new Uint8Array(gzipSync(b)) })) this.deliver(f.slice().buffer)
  }
  of(type: string) { return this.sent.filter((m) => m.type === type) }
  /** The drop of a live socket (no close() from the client). */
  drop(code: number) { this.readyState = 3; this.dispatchEvent(Object.assign(new Event('close'), { code, reason: '' })) }
}

function pngHeader(w: number, h: number) {
  const b = new Uint8Array(33), v = new DataView(b.buffer)
  b.set([0x89, 0x50, 0x4e, 0x47, 0x0d, 0x0a, 0x1a, 0x0a]); v.setUint32(8, 13); b.set([73, 72, 68, 82], 12)
  v.setUint32(16, w); v.setUint32(20, h); b.set([8, 6, 0, 0, 0], 24)
  return b
}
/** A PRINT_RASTER body: `u32 LE unitIndex` + the raster body. */
function page(unitIndex: number, w = 1920, h = 1080) {
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

/** Starts a 0..to print and returns its promise (observed), the reqId and the progress log. */
async function started(s: StreamSession, sock: FakeSocket, to = 1, extra: { signal?: AbortSignal } = {}) {
  const progress: Array<[number, number]> = []
  const p = s.print(0, to, { ...extra, onProgress: (d, t) => { progress.push([d, t]) } })
  void p.catch(() => {})
  await settle()
  const req = sock.of('print').at(-1)!
  return { p, reqId: req.reqId as number, progress }
}

/** The session still serves units after a failed job. */
async function usable(s: StreamSession, sock: FakeSocket) {
  const u = s.unit(0, 1)
  await settle()
  const req = sock.of('unit').at(-1)!
  sock.frame(3, req.reqId, encodeUnit({ unitId: 0, discriminator: 2, body: encodeRasterBody(2, 2, pngHeader(2, 2)) }))
  const lease = await u
  lease.release()
  expect(lease.kind).toBe('raster')
}

describe('StreamSession.print (scripted socket)', () => {
  it('happy path: print-start, two checked pages, print-done → PrintedPages in order, progress per page', async () => {
    const { s, sock } = await opened()
    const { p, reqId, progress } = await started(s, sock)
    expect(sock.of('print')[0]).toEqual({ reqId, type: 'print', token: tokenFor('j1'), from: 0, to: 1 })
    sock.text({ reqId, type: 'print-start', pages: 2, scale: 2 })
    sock.frame(8, reqId, page(0))
    sock.frame(8, reqId, page(1, 1919, 1079))
    sock.text({ reqId, type: 'print-done', pages: 2 })
    const pages = await p
    expect(pages.map((x) => [x.unitIndex, x.widthPx, x.heightPx, x.scale])).toEqual([[0, 960, 540, 2], [1, 960, 540, 2]])
    expect(pages.every((x) => x.png instanceof Blob && x.png.type === 'image/png')).toBe(true)
    expect(pages[0]!.png.size).toBe(33)
    expect(progress).toEqual([[1, 2], [2, 2]])
    expect(s.stats()).toMatchObject({ printing: false, printHeld: 0 })
    expect(sock.of('print-cancel')).toEqual([])
  })

  it('printSizeOk is the amended bound: (u - 1) * scale < w ≤ u * scale, integers only', () => {
    const u = { w: 960, h: 540 }
    expect(printSizeOk({ w: 1920, h: 1080 }, u, 2)).toBe(true)
    expect(printSizeOk({ w: 1919, h: 1079 }, u, 2)).toBe(true)
    expect(printSizeOk({ w: 1918, h: 1080 }, u, 2)).toBe(false)
    expect(printSizeOk({ w: 1921, h: 1080 }, u, 2)).toBe(false)
    expect(printSizeOk({ w: 1920, h: 1078 }, u, 2)).toBe(false)
    expect(printSizeOk({ w: 960, h: 540 }, u, 1)).toBe(true)
    expect(printSizeOk({ w: 959, h: 540 }, u, 1)).toBe(false)
    expect(printSizeOk({ w: 1919.5, h: 1080 }, u, 2)).toBe(false)
  })

  it('refusals carry limit/requested; an unknown code is policy-refused', async () => {
    const { s, sock } = await opened()
    for (const code of ['print-too-large', 'print-quota-exhausted'] as const) {
      const { p, reqId } = await started(s, sock, 2)
      sock.text({ reqId, type: 'refused', code, limit: 2, requested: 3 })
      const e = await failure(p) as StreamSessionError
      expect(e).toBeInstanceOf(StreamSessionError)
      expect({ code: e.code, limit: e.limit, requested: e.requested }).toEqual({ code, limit: 2, requested: 3 })
    }
    const a = await started(s, sock)
    sock.text({ reqId: a.reqId, type: 'refused', code: 'no-such-code' })
    expect((await failure(a.p) as StreamSessionError).code).toBe('policy-refused')
    const b = await started(s, sock)
    sock.text({ reqId: b.reqId, type: 'refused', code: 'rate-limited' })
    expect((await failure(b.p) as StreamSessionError).code).toBe('rate-limited')
    // A refusal from the server never needs a print-cancel.
    expect(sock.of('print-cancel')).toEqual([])
    await usable(s, sock)
  })

  it('a session-rotated refusal before print-start is re-issued once under a new reqId', async () => {
    const { s, sock } = await opened()
    const { p, reqId } = await started(s, sock)
    sock.text({ reqId, type: 'refused', code: 'session-rotated' })
    await settle()
    const again = sock.of('print')
    expect(again.length).toBe(2)
    expect(again[1]!.reqId).not.toBe(reqId)
    sock.text({ reqId: again[1]!.reqId, type: 'refused', code: 'session-rotated' })
    expect((await failure(p) as StreamSessionError).code).toBe('session-rotated')
  })

  const hostile: Array<[string, (sock: FakeSocket, reqId: number) => void]> = [
    ['print-start.pages ≠ to-from+1', (k, r) => k.text({ reqId: r, type: 'print-start', pages: 3, scale: 2 })],
    ['a scale other than 1 or 2', (k, r) => k.text({ reqId: r, type: 'print-start', pages: 2, scale: 3 })],
    ['a second print-start', (k, r) => { k.text({ reqId: r, type: 'print-start', pages: 2, scale: 2 }); k.text({ reqId: r, type: 'print-start', pages: 2, scale: 2 }) }],
    ['a wrong unit index', (k, r) => { k.text({ reqId: r, type: 'print-start', pages: 2, scale: 2 }); k.frame(8, r, page(1)) }],
    ['a duplicate page', (k, r) => { k.text({ reqId: r, type: 'print-start', pages: 2, scale: 2 }); k.frame(8, r, page(0)); k.frame(8, r, page(0)) }],
    ['an out-of-range unit index', (k, r) => { k.text({ reqId: r, type: 'print-start', pages: 2, scale: 2 }); k.frame(8, r, page(7)) }],
    ['a raster too wide', (k, r) => { k.text({ reqId: r, type: 'print-start', pages: 2, scale: 2 }); k.frame(8, r, page(0, 1921, 1080)) }],
    ['a raster too narrow', (k, r) => { k.text({ reqId: r, type: 'print-start', pages: 2, scale: 2 }); k.frame(8, r, page(0, 1918, 1080)) }],
    ['a raster too tall', (k, r) => { k.text({ reqId: r, type: 'print-start', pages: 2, scale: 2 }); k.frame(8, r, page(0, 1920, 1081)) }],
    ['a raster too short', (k, r) => { k.text({ reqId: r, type: 'print-start', pages: 2, scale: 2 }); k.frame(8, r, page(0, 1920, 1078)) }],
    ['a scale-2 page in a scale-1 job', (k, r) => { k.text({ reqId: r, type: 'print-start', pages: 2, scale: 1 }); k.frame(8, r, page(0)) }],
    ['a malformed page body', (k, r) => { k.text({ reqId: r, type: 'print-start', pages: 2, scale: 2 }); k.frame(8, r, page(0).subarray(0, 20)) }],
    ['a unit-kind frame on the job', (k, r) => { k.text({ reqId: r, type: 'print-start', pages: 2, scale: 2 }); k.frame(3, r, encodeUnit({ unitId: 0, discriminator: 2, body: encodeRasterBody(2, 2, pngHeader(2, 2)) })) }],
    ['a page before print-start', (k, r) => k.frame(8, r, page(0))],
    ['print-done before all pages', (k, r) => { k.text({ reqId: r, type: 'print-start', pages: 2, scale: 2 }); k.frame(8, r, page(0)); k.text({ reqId: r, type: 'print-done', pages: 2 }) }],
    ['print-done before print-start', (k, r) => k.text({ reqId: r, type: 'print-done', pages: 2 })],
    ['more pages than announced', (k, r) => { k.text({ reqId: r, type: 'print-start', pages: 2, scale: 2 }); k.frame(8, r, page(0)); k.frame(8, r, page(1)); k.frame(8, r, page(2)) }],
    ['print-done.pages ≠ n', (k, r) => { k.text({ reqId: r, type: 'print-start', pages: 2, scale: 2 }); k.frame(8, r, page(0)); k.frame(8, r, page(1)); k.text({ reqId: r, type: 'print-done', pages: 3 }) }],
    ['an unknown message type on the job', (k, r) => k.text({ reqId: r, type: 'print-page', pages: 2 })],
    ['a print-start over 48 KiB', (k, r) => k.text({ reqId: r, type: 'print-start', pages: 2, scale: 2, pad: 'x'.repeat(49 * 1024) })],
  ]
  for (const [name, script] of hostile) {
    it(`hostile server: ${name} → policy-refused, print-cancel, no pending job, session usable`, async () => {
      const { s, sock } = await opened()
      const { p, reqId, progress } = await started(s, sock)
      script(sock, reqId)
      const e = await failure(p) as StreamSessionError
      expect(e).toBeInstanceOf(StreamSessionError)
      expect(e.code).toBe('policy-refused')
      expect(progress.length).toBeLessThanOrEqual(2)
      expect(s.stats()).toMatchObject({ printing: false, printHeld: 0 })
      expect(sock.of('print-cancel')).toEqual([{ reqId, type: 'print-cancel' }])
      await usable(s, sock)
    })
  }

  it('late kind-8 frames after the job settled are rejected by expects and never revive it', async () => {
    const { s, sock } = await opened()
    const { p, reqId } = await started(s, sock, 0)
    sock.text({ reqId, type: 'print-start', pages: 1, scale: 2 })
    sock.frame(8, reqId, page(0))
    sock.text({ reqId, type: 'print-done', pages: 1 })
    expect((await p).length).toBe(1)
    const push = vi.spyOn(Reassembler.prototype, 'push')
    sock.frame(8, reqId, page(0))
    sock.text({ reqId, type: 'print-done', pages: 1 })
    await settle()
    const results = await Promise.all(push.mock.results.map((r) => r.value))
    expect(results).toEqual([{ ok: false, reqId, reason: 'unexpected-reqId' }])
    push.mockRestore()
    expect(s.stats()).toMatchObject({ printing: false, printHeld: 0 })
    // A failed job's late frames are refused the same way.
    const b = await started(s, sock, 0)
    sock.text({ reqId: b.reqId, type: 'print-start', pages: 2, scale: 2 })
    await failure(b.p)
    const push2 = vi.spyOn(Reassembler.prototype, 'push')
    sock.frame(8, b.reqId, page(0))
    await settle()
    expect(await push2.mock.results[0]!.value).toEqual({ ok: false, reqId: b.reqId, reason: 'unexpected-reqId' })
    push2.mockRestore()
    await usable(s, sock)
  })

  it('the blob ceiling (injected small) fails print-too-large with no limit, cancels, and releases held pages', async () => {
    expect(PRINT_JOB_MAX_BLOB_BYTES).toBe(192 * 1024 * 1024)
    const { s, sock } = await opened({ printMaxBlobBytes: 40 })
    const { p, reqId } = await started(s, sock)
    sock.text({ reqId, type: 'print-start', pages: 2, scale: 2 })
    sock.frame(8, reqId, page(0))
    await settle()
    expect(s.stats()).toMatchObject({ printing: true, printHeld: 1 })
    sock.frame(8, reqId, page(1))
    const e = await failure(p) as StreamSessionError
    expect(e.code).toBe('print-too-large')
    expect(e.limit).toBeUndefined()
    expect(e.requested).toBeUndefined()
    expect(sock.of('print-cancel')).toEqual([{ reqId, type: 'print-cancel' }])
    expect(s.stats()).toMatchObject({ printing: false, printHeld: 0 })
  })

  it('the ceiling hook only tightens: a larger injected value keeps 192 MiB', async () => {
    const { s } = await opened({ printMaxBlobBytes: 1e12 })
    expect(s.stats().printMaxBlobBytes).toBe(PRINT_JOB_MAX_BLOB_BYTES)
  })

  it('abort after print-start sends print-cancel, rejects with the abort reason, and removes its listener', async () => {
    const { s, sock } = await opened()
    const ac = new AbortController(), sig = ac.signal
    let added = 0, removed = 0
    const add = sig.addEventListener.bind(sig), remove = sig.removeEventListener.bind(sig)
    sig.addEventListener = ((...a: Parameters<typeof add>) => { if (a[0] === 'abort') added++; return add(...a) }) as typeof sig.addEventListener
    sig.removeEventListener = ((...a: Parameters<typeof remove>) => { if (a[0] === 'abort') removed++; return remove(...a) }) as typeof sig.removeEventListener
    // Settled jobs (success, refusal) release the signal too.
    const ok = await started(s, sock, 0, { signal: sig })
    sock.text({ reqId: ok.reqId, type: 'print-start', pages: 1, scale: 2 })
    sock.frame(8, ok.reqId, page(0))
    sock.text({ reqId: ok.reqId, type: 'print-done', pages: 1 })
    await ok.p
    const refused = await started(s, sock, 0, { signal: sig })
    sock.text({ reqId: refused.reqId, type: 'refused', code: 'rate-limited' })
    await failure(refused.p)
    const { p, reqId } = await started(s, sock, 1, { signal: sig })
    sock.text({ reqId, type: 'print-start', pages: 2, scale: 2 })
    sock.frame(8, reqId, page(0))
    await settle()
    ac.abort()
    const e = await failure(p)
    expect((e as Error).name).toBe('AbortError')
    expect(sock.of('print-cancel')).toEqual([{ reqId, type: 'print-cancel' }])
    expect(s.stats()).toMatchObject({ printing: false, printHeld: 0 })
    expect(added).toBe(3)
    expect(removed).toBe(added)
    // A page still in flight from the cancelled job is dropped.
    sock.frame(8, reqId, page(1))
    await settle()
    await usable(s, sock)
  })

  it('an already-aborted signal rejects without sending', async () => {
    const { s, sock } = await opened()
    const ac = new AbortController(); ac.abort()
    expect(((await failure(s.print(0, 0, { signal: ac.signal }))) as Error).name).toBe('AbortError')
    expect(sock.of('print')).toEqual([])
  })

  it('a socket drop mid-job rejects session-rotated and the job is never re-issued after the resume', async () => {
    const { s, sock, socks, manifest } = await opened()
    const { p, reqId } = await started(s, sock)
    sock.text({ reqId, type: 'print-start', pages: 2, scale: 2 })
    sock.frame(8, reqId, page(0))
    await settle()
    sock.drop(1006)
    expect((await failure(p) as StreamSessionError).code).toBe('session-rotated')
    expect(s.stats()).toMatchObject({ printing: false, printHeld: 0 })
    await settle(300)
    const next = socks[1]!
    next.open(); await settle()
    next.frame(1, next.of('hello')[0]!.reqId, encodeManifest(manifest))
    await settle()
    expect(next.of('print')).toEqual([])
  })

  it('a terminal socket end mid-job rejects with the session failure', async () => {
    const { s, sock } = await opened()
    const { p, reqId } = await started(s, sock)
    sock.text({ reqId, type: 'print-start', pages: 2, scale: 2 })
    await settle()
    sock.drop(4005)
    expect((await failure(p) as StreamSessionError).code).toBe('revoked')
    expect(s.stats()).toMatchObject({ printing: false, printHeld: 0 })
  })

  it('close() mid-job rejects with an AbortError', async () => {
    const { s, sock } = await opened()
    const { p } = await started(s, sock)
    s.close()
    expect(((await failure(p)) as Error).name).toBe('AbortError')
  })

  it('one job at a time: a second print() while one is pending rejects at once without sending', async () => {
    const { s, sock } = await opened()
    const { p, reqId } = await started(s, sock)
    const e = await failure(s.print(0, 0, {})) as StreamSessionError
    expect(e.code).toBe('rate-limited')
    expect(sock.of('print').length).toBe(1)
    sock.text({ reqId, type: 'print-start', pages: 2, scale: 2 })
    sock.frame(8, reqId, page(0)); sock.frame(8, reqId, page(1))
    sock.text({ reqId, type: 'print-done', pages: 2 })
    expect((await p).length).toBe(2)
  })

  it('a malformed range or print off rejects locally without sending', async () => {
    const { s, sock } = await opened()
    for (const [f, t] of [[1, 0], [-1, 0], [0, 3], [0.5, 1], [0, NaN]] as const) {
      expect((await failure(s.print(f, t, {})) as StreamSessionError).code).toBe('policy-refused')
    }
    expect(sock.of('print')).toEqual([])
    const off = await opened({}, { can: { ...exampleManifest().can, print: false }, caps: { ...exampleManifest().caps, print: null } })
    expect((await failure(off.s.print(0, 0, {})) as StreamSessionError).code).toBe('policy-refused')
    expect(off.sock.of('print')).toEqual([])
  })

  it('a throwing onProgress fails the job with its error and cancels', async () => {
    const { s, sock } = await opened()
    const boom = new Error('boom')
    const p = s.print(0, 1, { onProgress: () => { throw boom } })
    void p.catch(() => {})
    await settle()
    const reqId = sock.of('print')[0]!.reqId
    sock.text({ reqId, type: 'print-start', pages: 2, scale: 2 })
    sock.frame(8, reqId, page(0))
    expect(await failure(p)).toBe(boom)
    expect(sock.of('print-cancel')).toEqual([{ reqId, type: 'print-cancel' }])
    expect(s.stats()).toMatchObject({ printing: false, printHeld: 0 })
  })

  it('RemoteOfficeDoc.print passes the range and options through', async () => {
    const { s, sock } = await opened()
    const doc = createRemoteDoc(s)
    const progress: number[] = []
    const p = doc.print({ from: 1, to: 1 }, { onProgress: (d) => { progress.push(d) } })
    await settle()
    const req = sock.of('print')[0]!
    expect(req).toMatchObject({ from: 1, to: 1 })
    sock.text({ reqId: req.reqId, type: 'print-start', pages: 1, scale: 1 })
    sock.frame(8, req.reqId, page(1, 960, 540))
    sock.text({ reqId: req.reqId, type: 'print-done', pages: 1 })
    const pages = await p
    expect(pages.map((x) => [x.unitIndex, x.scale])).toEqual([[1, 1]])
    expect(progress).toEqual([1])
  })
})
