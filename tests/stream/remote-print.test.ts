/** Stage 3b Task 5: StreamSession.print() — every page checked against the job, fail closed. */
import { describe, expect, it, vi } from 'vitest'
import { deflateSync, gzipSync } from 'node:zlib'
import { randomBytes } from 'node:crypto'
import { encodeMessage, Reassembler } from '../../src/stream/frame'
import { encodeManifest, type Manifest } from '../../src/stream/manifest'
import { encodeRasterBody } from '../../src/stream/raster'
import { encodeUnit } from '../../src/stream/unit'
import { PRINT_PAGE_BODY_MAX, printPageCap, StreamSession, type SessionOptions } from '../../src/stream/remote/session'
import { StreamSessionError } from '../../src/stream/remote/errors'
import { PRINT_JOB_MAX_BLOB_BYTES, PRINT_JOB_MAX_DECODED, printSizeOk } from '../../src/stream/remote/print-wire'
import { crc32 } from '../../src/stream/crc32'
import { createRemoteDoc } from '../../src/stream/remote/remote-doc'
import { SESSION_TRANSIENT_BYTES, unitCost } from '../../src/stream/remote/reservation'

const MiB = 1024 * 1024
/** The page-body cap and grant of a print job over units of this size (fix round 2: encoded + 2 × body). */
const pageCapOf = (w: number, h: number) => Math.min(64 * MiB, Math.floor((SESSION_TRANSIENT_BYTES - 8 * MiB) / 2), 16 + unitCost(w, h, 2, false))
const printGrant = (w: number, h: number) => 8 * MiB + 2 * pageCapOf(w, h)
/** An incompressible page body big enough to span several 1 MiB fragments. */
function bigPage(unitIndex: number, w = 1920, h = 1080) {
  const png = new Uint8Array(randomBytes(1_500_000))
  png.set(pngHeader(w, h))
  const r = encodeRasterBody(w, h, png), out = new Uint8Array(4 + r.length)
  new DataView(out.buffer).setUint32(0, unitIndex, true); out.set(r, 4)
  return out
}

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
  /** Message types whose send throws (a transport failure). */
  throwOn = new Set<string>()
  send(d: string) {
    const m = JSON.parse(d)
    if (this.throwOn.has(m.type)) throw new Error(`send ${m.type} failed`)
    this.sent.push(m)
  }
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
  fragments(kind: number, reqId: number, payload: Uint8Array) {
    return encodeMessage({ ver: 1, kind: kind as never, reqId, payload, compress: (b) => new Uint8Array(gzipSync(b)) }).map((f) => f.slice().buffer)
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

describe('StreamSession.print — fix round 1 (admission, reassembly, throwing sends)', () => {
  const big = { units: { count: 3, list: [{ w: 4096, h: 4096 }, { w: 2000, h: 1200 }, { w: 960, h: 540 }] } }

  it('a job holds a grant from the session admission for its life, released once when it settles', async () => {
    const { s, sock } = await opened()
    const { p, reqId } = await started(s, sock, 1)
    expect(s.stats()).toMatchObject({ inFlight: 1, inFlightBytes: printGrant(960, 540) })
    sock.text({ reqId, type: 'print-start', pages: 2, scale: 2 })
    sock.frame(8, reqId, page(0)); sock.frame(8, reqId, page(1))
    sock.text({ reqId, type: 'print-done', pages: 2 })
    await p
    await settle()
    expect(s.stats()).toMatchObject({ inFlight: 0, inFlightBytes: 0 })
  })

  it('print waits behind units holding the admission and is sent when they release', async () => {
    const { s, sock } = await opened({}, big)
    expect(unitCost(4096, 4096, 1, false) + printGrant(2000, 1200)).toBeGreaterThan(SESSION_TRANSIENT_BYTES)
    const u = s.unit(0, 1)
    await settle()
    const p = s.print(1, 1, {})
    void p.catch(() => {})
    await settle()
    expect(sock.of('print')).toEqual([])
    expect(s.stats().printing).toBe(true)
    // The slot is claimed while queued: a second job is refused at once.
    expect((await failure(s.print(2, 2, {})) as StreamSessionError).code).toBe('rate-limited')
    sock.frame(3, sock.of('unit')[0]!.reqId, encodeUnit({ unitId: 0, discriminator: 2, body: encodeRasterBody(2, 2, pngHeader(2, 2)) }))
    const lease = await u
    await settle()
    expect(sock.of('print')).toEqual([])
    lease.release()
    await settle()
    const req = sock.of('print')[0]!
    expect(req).toMatchObject({ from: 1, to: 1 })
    sock.text({ reqId: req.reqId, type: 'print-start', pages: 1, scale: 2 })
    sock.frame(8, req.reqId, page(1, 4000, 2400))
    sock.text({ reqId: req.reqId, type: 'print-done', pages: 1 })
    expect((await p).length).toBe(1)
  })

  it('abort while queued for admission sends nothing, removes the listener, and frees the slot', async () => {
    const { s, sock } = await opened({}, big)
    const u = s.unit(0, 1)
    await settle()
    const ac = new AbortController(), sig = ac.signal
    let added = 0, removed = 0
    const add = sig.addEventListener.bind(sig), remove = sig.removeEventListener.bind(sig)
    sig.addEventListener = ((...a: Parameters<typeof add>) => { if (a[0] === 'abort') added++; return add(...a) }) as typeof sig.addEventListener
    sig.removeEventListener = ((...a: Parameters<typeof remove>) => { if (a[0] === 'abort') removed++; return remove(...a) }) as typeof sig.removeEventListener
    const p = s.print(1, 1, { signal: sig })
    await settle()
    ac.abort()
    expect(((await failure(p)) as Error).name).toBe('AbortError')
    expect(added).toBe(1); expect(removed).toBe(1)
    expect(s.stats().printing).toBe(false)
    sock.frame(3, sock.of('unit')[0]!.reqId, encodeUnit({ unitId: 0, discriminator: 2, body: encodeRasterBody(2, 2, pngHeader(2, 2)) }))
    ;(await u).release()
    await settle()
    expect(sock.of('print')).toEqual([])
    expect(sock.of('print-cancel')).toEqual([])
    expect(s.stats()).toMatchObject({ inFlight: 0, inFlightBytes: 0 })
  })

  it('a refusal, a hostile page and a socket drop each release the grant', async () => {
    const { s, sock } = await opened()
    const a = await started(s, sock)
    sock.text({ reqId: a.reqId, type: 'refused', code: 'print-quota-exhausted', limit: 1, requested: 2 })
    await failure(a.p); await settle()
    expect(s.stats()).toMatchObject({ inFlight: 0, inFlightBytes: 0 })
    const b = await started(s, sock)
    sock.text({ reqId: b.reqId, type: 'print-start', pages: 2, scale: 2 })
    sock.frame(8, b.reqId, page(1))
    await failure(b.p); await settle()
    expect(s.stats()).toMatchObject({ inFlight: 0, inFlightBytes: 0 })
    const c = await started(s, sock)
    sock.text({ reqId: c.reqId, type: 'refused', code: 'session-rotated' })
    await settle()
    // The retry keeps the one grant.
    expect(s.stats()).toMatchObject({ inFlight: 1, inFlightBytes: printGrant(960, 540) })
    sock.drop(1006)
    await failure(c.p); await settle()
    expect(s.stats()).toMatchObject({ inFlight: 0, inFlightBytes: 0 })
  })

  it('kind-8 inflation is bounded by the job grant, not the blanket 64 MiB', async () => {
    const { s, sock } = await opened()
    const { p, reqId } = await started(s, sock, 0)
    sock.text({ reqId, type: 'print-start', pages: 1, scale: 2 })
    // A body over the job's page cap (but well under 64 MiB) is a decoded-cap rejection.
    const over = new Uint8Array(pageCapOf(960, 540) + 1)
    const push = vi.spyOn(Reassembler.prototype, 'push')
    sock.frame(8, reqId, over)
    const e = await failure(p) as StreamSessionError
    expect(e.code).toBe('policy-refused')
    const results = (await Promise.all(push.mock.results.map((r) => r.value))).filter(Boolean)
    expect(results).toEqual([{ ok: false, reqId, reason: 'decoded-cap' }])
    push.mockRestore()
  })

  it('settlement drops a partially reassembled page: no buffered bytes remain, with no further traffic', async () => {
    const { s, sock } = await opened()
    const ac = new AbortController()
    const { p, reqId } = await started(s, sock, 1, { signal: ac.signal })
    sock.text({ reqId, type: 'print-start', pages: 2, scale: 2 })
    const frags = sock.fragments(8, reqId, bigPage(0))
    expect(frags.length).toBeGreaterThan(1)
    sock.deliver(frags[0]!)
    await settle()
    expect(s.stats().reassemblyBytes).toBeGreaterThan(0)
    ac.abort()
    await failure(p)
    await settle(1100)
    expect(s.stats()).toMatchObject({ printing: false, reassemblyBytes: 0 })
    await usable(s, sock)
  })

  it('Reassembler.drop releases a pending message and is idempotent', async () => {
    const r = new Reassembler({ acceptedVersions: new Set([1]), expects: () => true, encodedCap: () => 8 * MiB, decodedCap: () => 64 * MiB })
    const frags = new FakeSocket().fragments(8, 5, bigPage(0))
    expect(await r.push(new Uint8Array(frags[0]!))).toBeNull()
    expect(r.bufferedBytes).toBeGreaterThan(0)
    r.drop(5); r.drop(5); r.drop(99)
    expect(r.bufferedBytes).toBe(0)
    // The rest of the dropped message is out of sequence now.
    expect(await r.push(new Uint8Array(frags[1]!))).toMatchObject({ ok: false, reqId: 5, reason: 'sequence' })
  })

  it('a throwing send(print) settles the job: rejected, slot and listener freed, grant released; the next job works', async () => {
    const { s, sock } = await opened()
    const ac = new AbortController(), sig = ac.signal
    let added = 0, removed = 0
    const add = sig.addEventListener.bind(sig), remove = sig.removeEventListener.bind(sig)
    sig.addEventListener = ((...a: Parameters<typeof add>) => { if (a[0] === 'abort') added++; return add(...a) }) as typeof sig.addEventListener
    sig.removeEventListener = ((...a: Parameters<typeof remove>) => { if (a[0] === 'abort') removed++; return remove(...a) }) as typeof sig.removeEventListener
    sock.throwOn.add('print')
    const e = await failure(s.print(0, 0, { signal: sig })) as StreamSessionError
    expect(e.code).toBe('policy-refused')
    await settle()
    expect(s.stats()).toMatchObject({ printing: false, inFlight: 0, inFlightBytes: 0 })
    expect(added).toBe(1); expect(removed).toBe(1)
    sock.throwOn.delete('print')
    const { p, reqId } = await started(s, sock, 0)
    sock.text({ reqId, type: 'print-start', pages: 1, scale: 2 })
    sock.frame(8, reqId, page(0))
    sock.text({ reqId, type: 'print-done', pages: 1 })
    expect((await p).length).toBe(1)
  })

  it('a throwing send(print) on the session-rotated retry settles the job too', async () => {
    const { s, sock } = await opened()
    const { p, reqId } = await started(s, sock, 0)
    sock.throwOn.add('print')
    sock.text({ reqId, type: 'refused', code: 'session-rotated' })
    expect((await failure(p) as StreamSessionError).code).toBe('policy-refused')
    await settle()
    expect(s.stats()).toMatchObject({ printing: false, inFlight: 0, inFlightBytes: 0 })
  })

  it('a throwing send(print-cancel) still rejects the job and frees everything', async () => {
    const { s, sock } = await opened()
    const ac = new AbortController()
    const { p, reqId } = await started(s, sock, 1, { signal: ac.signal })
    sock.text({ reqId, type: 'print-start', pages: 2, scale: 2 })
    sock.throwOn.add('print-cancel')
    ac.abort()
    expect(((await failure(p)) as Error).name).toBe('AbortError')
    await settle()
    expect(s.stats()).toMatchObject({ printing: false, inFlight: 0, inFlightBytes: 0 })
    // A hostile page with a throwing cancel rejects as well.
    const b = await started(s, sock, 0)
    sock.text({ reqId: b.reqId, type: 'print-start', pages: 3, scale: 2 })
    expect((await failure(b.p) as StreamSessionError).code).toBe('policy-refused')
    expect(s.stats().printing).toBe(false)
  })

  it('fix round 2: the grant is encoded + 2 × pageCap; a 4096×4096 page body is clamped so the grant fits the aggregate', async () => {
    expect(PRINT_PAGE_BODY_MAX).toBe(44 * MiB)
    const list = [{ w: 4096, h: 4096 }, { w: 960, h: 540 }]
    expect(printPageCap(list, 1, 1)).toBe(16 + unitCost(960, 540, 2, false))
    expect(printPageCap(list, 0, 1)).toBe(PRINT_PAGE_BODY_MAX)
    const { s, sock } = await opened({}, { units: { count: 2, list } })
    const { p, reqId } = await started(s, sock, 1)
    expect(s.stats().inFlightBytes).toBe(8 * MiB + 2 * PRINT_PAGE_BODY_MAX)
    expect(s.stats().inFlightBytes).toBeLessThanOrEqual(SESSION_TRANSIENT_BYTES)
    // A body over the clamped cap is a decoded-cap rejection: the job fails and its grant is released.
    sock.text({ reqId, type: 'print-start', pages: 2, scale: 2 })
    const push = vi.spyOn(Reassembler.prototype, 'push')
    sock.frame(8, reqId, new Uint8Array(PRINT_PAGE_BODY_MAX + 1))
    expect((await failure(p) as StreamSessionError).code).toBe('policy-refused')
    const results = (await Promise.all(push.mock.results.map((r) => r.value))).filter(Boolean)
    expect(results.at(-1)).toEqual({ ok: false, reqId, reason: 'decoded-cap' })
    push.mockRestore()
    await settle()
    expect(s.stats()).toMatchObject({ printing: false, inFlight: 0, inFlightBytes: 0 })
  }, 30_000)
})

/** A valid grayscale PNG (IHDR, one deflated all-zero IDAT, IEND, real CRCs): highly compressible. */
function validPng(w: number, h: number): Uint8Array {
  const chunk = (name: string, data: Buffer) => {
    const t = Buffer.from(name), out = Buffer.alloc(data.length + 12)
    out.writeUInt32BE(data.length); t.copy(out, 4); data.copy(out, 8)
    out.writeUInt32BE(crc32(Buffer.concat([t, data])), out.length - 4)
    return out
  }
  const ih = Buffer.alloc(13); ih.writeUInt32BE(w, 0); ih.writeUInt32BE(h, 4); ih[8] = 8
  return new Uint8Array(Buffer.concat([Buffer.from([137, 80, 78, 71, 13, 10, 26, 10]), chunk('IHDR', ih), chunk('IDAT', deflateSync(Buffer.alloc((w + 1) * h))), chunk('IEND', Buffer.alloc(0))]))
}
const pngCache = new Map<string, Uint8Array>()
function validPage(unitIndex: number, w: number, h: number) {
  const key = `${w}x${h}`
  if (!pngCache.has(key)) pngCache.set(key, validPng(w, h))
  const r = encodeRasterBody(w, h, pngCache.get(key)!), out = new Uint8Array(4 + r.length)
  new DataView(out.buffer).setUint32(0, unitIndex, true); out.set(r, 4)
  return out
}
/** Counts the abort listeners a job adds to / removes from its signal. */
function countedSignal() {
  const ac = new AbortController(), sig = ac.signal, n = { added: 0, removed: 0 }
  const add = sig.addEventListener.bind(sig), remove = sig.removeEventListener.bind(sig)
  sig.addEventListener = ((...a: Parameters<typeof add>) => { if (a[0] === 'abort') n.added++; return add(...a) }) as typeof sig.addEventListener
  sig.removeEventListener = ((...a: Parameters<typeof remove>) => { if (a[0] === 'abort') n.removed++; return remove(...a) }) as typeof sig.removeEventListener
  return { sig, n }
}

describe('StreamSession.print — final fix 1: the whole-job decoded budget', () => {
  const big = (n: number, extra: Array<{ w: number; h: number }> = []) => {
    const list = [...Array.from({ length: n }, () => ({ w: 2048, h: 2048 })), ...extra]
    return { units: { count: list.length, list } }
  }

  it('PRINT_JOB_MAX_DECODED is the server twin (1 GiB)', () => {
    expect(PRINT_JOB_MAX_DECODED).toBe(1024 * 1024 * 1024)
  })

  it('exactly the budget (16 units of 2048² at scale 2 = 1 GiB of RGBA) prints every page', async () => {
    expect(16 * 4096 * 4096 * 4).toBe(PRINT_JOB_MAX_DECODED)
    const { s, sock } = await opened({}, big(16))
    const { p, reqId } = await started(s, sock, 15)
    sock.text({ reqId, type: 'print-start', pages: 16, scale: 2 })
    for (let i = 0; i < 16; i++) sock.frame(8, reqId, validPage(i, 4096, 4096))
    sock.text({ reqId, type: 'print-done', pages: 16 })
    const pages = await p
    expect(pages.map((x) => x.unitIndex)).toEqual([...Array(16).keys()])
    expect(sock.of('print-cancel')).toEqual([])
    await settle()
    expect(s.stats()).toMatchObject({ printing: false, printHeld: 0, inFlight: 0, inFlightBytes: 0, reassemblyBytes: 0 })
    s.close()
  }, 60_000)

  it('one RGBA step over the budget (16 × 2048² + one 1×1 unit at scale 2) fails at print-start, print-too-large with no limit', async () => {
    expect(16 * 4096 * 4096 * 4 + 2 * 2 * 4).toBeGreaterThan(PRINT_JOB_MAX_DECODED)
    const { s, sock } = await opened({}, big(16, [{ w: 1, h: 1 }]))
    const { sig, n } = countedSignal()
    const { p, reqId } = await started(s, sock, 16, { signal: sig })
    sock.text({ reqId, type: 'print-start', pages: 17, scale: 2 })
    const e = await failure(p) as StreamSessionError
    expect(e.code).toBe('print-too-large')
    expect(e.limit).toBeUndefined()
    expect(e.requested).toBeUndefined()
    expect(sock.of('print-cancel')).toEqual([{ reqId, type: 'print-cancel' }])
    expect(n).toEqual({ added: 1, removed: 1 })
    await settle()
    expect(s.stats()).toMatchObject({ printing: false, printHeld: 0, inFlight: 0, inFlightBytes: 0, reassemblyBytes: 0 })
    // Late pages of the failed job are ignored; the session carries on.
    sock.frame(8, reqId, validPage(0, 4096, 4096))
    await settle()
    expect(s.stats()).toMatchObject({ printing: false, reassemblyBytes: 0 })
    await usable(s, sock)
    s.close()
  }, 60_000)

  it('the highly compressible 17 × 2048² job at scale 2 (1,140,850,688 B decoded, ~280 KB of PNG) is refused before any page is kept', async () => {
    expect(17 * 4096 * 4096 * 4).toBe(1_140_850_688)
    const { s, sock } = await opened({}, big(17))
    const { sig, n } = countedSignal()
    const { p, reqId } = await started(s, sock, 16, { signal: sig })
    sock.text({ reqId, type: 'print-start', pages: 17, scale: 2 })
    for (let i = 0; i < 17; i++) sock.frame(8, reqId, validPage(i, 4096, 4096))
    sock.text({ reqId, type: 'print-done', pages: 17 })
    const e = await failure(p) as StreamSessionError
    expect(e.code).toBe('print-too-large')
    expect(e.limit).toBeUndefined()
    expect(sock.of('print-cancel')).toEqual([{ reqId, type: 'print-cancel' }])
    expect(n).toEqual({ added: 1, removed: 1 })
    await settle()
    expect(s.stats()).toMatchObject({ printing: false, printHeld: 0, inFlight: 0, inFlightBytes: 0, reassemblyBytes: 0 })
    await usable(s, sock)
    s.close()
  }, 60_000)

  it('the budget follows the announced scale: the same 17 units at scale 1 (272 MiB) print', async () => {
    const { s, sock } = await opened({}, big(17))
    const { p, reqId } = await started(s, sock, 16)
    sock.text({ reqId, type: 'print-start', pages: 17, scale: 1 })
    for (let i = 0; i < 17; i++) sock.frame(8, reqId, validPage(i, 2048, 2048))
    sock.text({ reqId, type: 'print-done', pages: 17 })
    expect((await p).length).toBe(17)
    s.close()
  }, 60_000)
})
