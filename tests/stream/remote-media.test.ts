// @vitest-environment node
/**
 * Piece 5 Task 10 (spec §4.3, §4.4; plan D1, D4, D12, D13, D16, D20): the client media transport — `mediaShow`,
 * `mediaCursor`, `mediaRequest`, `mediaEnd` on StreamSession — and the remote media lease (`openRemoteMediaLease`).
 * Kind 9 is admitted only for a live, started `media-request`, short of its declared chunks and bytes; anything
 * unsolicited or overrunning is dropped before allocation. The lease charges the show workspace as `openMediaLease`.
 */
import { gzipSync } from 'node:zlib'
import { afterEach, describe, expect, it, vi } from 'vitest'
import { encodeMessage, Kind } from '../../src/stream/frame'
import { encodeManifest, MEDIA_AUDIO_CAPS, type Manifest } from '../../src/stream/manifest'
import { MEDIA_WIRE_LIMITS } from '../../src/stream/media-wire'
import { StreamSession, type SessionOptions } from '../../src/stream/remote/session'
import { openRemoteMediaLease } from '../../src/stream/remote/media-lease'
import { createSlideShowWorkspace } from '../../src/slideshow/workspace'
import type { MediaRecord, SlideShowWorkspace } from '../../src/slideshow/types'

const SEND = MEDIA_WIRE_LIMITS.sendChunk
const b64url = (o: object) => btoa(JSON.stringify(o)).replace(/\+/g, '-').replace(/\//g, '_').replace(/=+$/, '')
const tokenFor = (jti: string) => `h.${b64url({ jti })}.s`
const settle = async (ms = 20) => { if (vi.isFakeTimers()) await vi.advanceTimersByTimeAsync(ms); else await new Promise((r) => setTimeout(r, ms)) }
const MID = 'A'.repeat(21) + 'b', MID2 = 'B'.repeat(21) + 'c'

afterEach(() => { vi.useRealTimers(); vi.restoreAllMocks() })

const exampleManifest = (media = true): Manifest => ({
  v: 1, docRef: 'a'.repeat(24), kind: 'pptx', name: 'deck.pptx',
  units: { count: 3, list: [{ w: 960, h: 540, ...(media ? { media: true as const } : {}) }, { w: 960, h: 540 }, { w: 960, h: 540, hidden: true }] },
  layoutFingerprint: 'f'.repeat(64),
  session: { sid: '0123456789ab', token: tokenFor('j1'), expiresSec: 300, resume: 'session' },
  policy: { contentPolicy: 'visible-only', protection: 'final-raster', semantic: 'visible-only', contentPolicyId: 'visible-only@1', viewerMarkId: 'fixed' },
  can: { search: true, copy: false, print: true, textA11y: false, present: true, notes: true, links: false, media },
  caps: { copy: null, a11y: 'off', print: { maxPages: null, maxSessionPages: null }, ...(media ? { media: { audio: { ...MEDIA_AUDIO_CAPS } } } : {}) },
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
  /** One MEDIA_CHUNK (kind 9), stored (D14). */
  chunk(reqId: number, payload: Uint8Array) {
    for (const f of encodeMessage({ ver: 1, kind: Kind.MEDIA_CHUNK, reqId, payload, compress: () => { throw new Error('stored') }, stored: true })) this.deliver(f.slice().buffer)
  }
  of(type: string) { return this.sent.filter((m) => m.type === type) }
  drop(code: number) { this.readyState = 3; this.dispatchEvent(Object.assign(new Event('close'), { code, reason: '' })) }
}

async function opened(o: Partial<SessionOptions> = {}, manifest: Manifest = exampleManifest()) {
  const socks: FakeSocket[] = []
  const client = () => { const s = new FakeSocket(); socks.push(s); return s as never }
  const p = StreamSession.open({ kind: 'stream', url: 'wss://x/stream' }, { client, faces: async () => [], ...o })
  const sock = socks[0]!
  sock.open(); await settle()
  sock.frame(1, sock.of('hello')[0]!.reqId, encodeManifest(manifest))
  const s = await p
  return { s, sock, socks, manifest }
}

const noise = (n: number, seed = 7) => { const b = new Uint8Array(n); let x = seed; for (let i = 0; i < n; i++) { x = (x * 1103515245 + 12345) >>> 0; b[i] = x >>> 24 } return b }
const startOf = (reqId: number, bytes: number, extra: Record<string, unknown> = {}) =>
  ({ reqId, type: 'media-start', mediaId: MID, contentType: 'audio/mp4', durationMs: 1000, bytes, chunks: Math.ceil(bytes / SEND), ...extra })
const rec = (rId = MID): MediaRecord => ({ rId, partPath: 'remote:' + rId, contentType: 'audio/mp4', declaredBytes: 0 })

/** A workspace whose peak `used()` is recorded after every lease operation. */
function watched(limit = 96 << 20) {
  const ws = createSlideShowWorkspace('desktop', { limit })
  let peak = 0
  const look = () => { peak = Math.max(peak, ws.used()) }
  const out: SlideShowWorkspace & { peak(): number } = {
    ...ws,
    used: () => ws.used(),
    reserveTerminal: (n) => ws.reserveTerminal(n),
    lease(bytes, tag) {
      const l = ws.lease(bytes, tag)
      look()
      if (!l) return l
      return {
        get tag() { return l.tag }, get bytes() { return l.bytes },
        grow(n) { const ok = l.grow(n); look(); return ok },
        shrink(n) { l.shrink(n); look() },
        transfer: (t) => l.transfer(t),
        release() { l.release(); look() },
      }
    },
    peak: () => peak,
  }
  return out
}

/** A realm: Blob URLs recorded, `canPlayType` answered as given. */
function realm(canPlay = 'maybe') {
  const made: Blob[] = [], revoked: string[] = [], asked: string[] = []
  const win = {
    URL: {
      createObjectURL: (b: Blob) => { made.push(b); return `blob:x/${made.length}` },
      revokeObjectURL: (u: string) => { revoked.push(u) },
    },
    document: { createElement: () => ({ canPlayType: (t: string) => { asked.push(t); return canPlay } }) },
  } as unknown as Window
  return { win, made, revoked, asked }
}

const concat = (parts: Uint8Array[]) => { const n = parts.reduce((a, p) => a + p.length, 0), out = new Uint8Array(n); let at = 0; for (const p of parts) { out.set(p, at); at += p.length } return out }

describe('remote media lease (spec §4.4)', () => {
  it('happy path: start, chunks, Blob of bytes; one lease; charge peak 2×bytes then bytes', async () => {
    const { s, sock } = await opened()
    const ws = watched(), r = realm(), ac = new AbortController()
    const p = openRemoteMediaLease(s, rec(), ws, r.win, ac.signal, 0)
    await settle()
    const [req] = sock.of('media-request')
    expect(req).toEqual({ reqId: req!.reqId, type: 'media-request', token: s.token, mediaId: MID, unit: 0 })
    expect(r.asked).toEqual(['audio/mp4; codecs="mp4a.40.2"'])
    const bytes = SEND + 100, a = noise(SEND, 1), b = noise(100, 2)
    sock.text(startOf(req!.reqId, bytes))
    await settle()
    expect(ws.used()).toBe(bytes) // charged at media-start
    sock.chunk(req!.reqId, a)
    sock.chunk(req!.reqId, b)
    const lease = await p
    if ('refused' in lease) throw new Error(lease.refused)
    expect(r.made).toHaveLength(1)
    expect(r.made[0]!.type).toBe('audio/mp4')
    expect(new Uint8Array(await r.made[0]!.arrayBuffer())).toEqual(concat([a, b]))
    expect(lease.url).toBe('blob:x/1')
    expect(lease.bytes).toBe(bytes)
    expect(ws.peak()).toBe(2 * bytes)
    expect(ws.used()).toBe(bytes)
    expect(s.stats().inFlightBytes).toBe(0)
    expect(s.stats().reassemblyBytes).toBe(0)
    lease.release()
    expect(ws.used()).toBe(0)
  })

  it('chunk past declared chunks → rejected, grant released', async () => {
    const { s, sock } = await opened()
    const ws = watched(), r = realm()
    const p = openRemoteMediaLease(s, rec(), ws, r.win, new AbortController().signal)
    await settle()
    const id = sock.of('media-request')[0]!.reqId
    sock.text(startOf(id, 100))
    sock.chunk(id, noise(100))
    sock.chunk(id, noise(10))
    expect(await p).toEqual({ refused: 'policy-refused' })
    expect(r.made).toHaveLength(0)
    expect(ws.used()).toBe(0)
    expect(s.stats().inFlightBytes).toBe(0)
    expect(s.stats().reassemblyBytes).toBe(0)
  })

  it('Σ chunk bytes over declared → rejected', async () => {
    const { s, sock } = await opened()
    const ws = watched(), r = realm()
    const p = openRemoteMediaLease(s, rec(), ws, r.win, new AbortController().signal)
    await settle()
    const id = sock.of('media-request')[0]!.reqId
    sock.text(startOf(id, 100))
    sock.chunk(id, noise(101))
    expect(await p).toEqual({ refused: 'policy-refused' })
    await settle()
    expect(ws.used()).toBe(0)
    expect(s.stats().inFlightBytes).toBe(0)
    // An underrun (the declared count reached short of the declared bytes) is rejected too.
    const q = openRemoteMediaLease(s, rec(), ws, r.win, new AbortController().signal)
    await settle()
    const id2 = sock.of('media-request')[1]!.reqId
    sock.text(startOf(id2, 100))
    sock.chunk(id2, noise(50))
    expect(await q).toEqual({ refused: 'policy-refused' })
    expect(ws.used()).toBe(0)
    expect(r.made).toHaveLength(0)
  })

  it('media-start over the 32 MiB audio cap, or with contentType video/mp4 → rejected before any chunk', async () => {
    const { s, sock } = await opened()
    const ws = watched()
    for (const extra of [{ bytes: MEDIA_WIRE_LIMITS.audioBytes + 1, chunks: Math.ceil((MEDIA_WIRE_LIMITS.audioBytes + 1) / SEND) }, { contentType: 'video/mp4' }]) {
      const onStart = vi.fn(() => true)
      const p = s.mediaRequest(MID, new AbortController().signal, onStart)
      await settle()
      const id = sock.of('media-request').at(-1)!.reqId
      sock.text(startOf(id, 100, extra))
      sock.chunk(id, noise(100))
      expect(await p).toEqual({ refused: 'policy-refused' })
      expect(onStart).not.toHaveBeenCalled()
      await settle()
      expect(s.stats().reassemblyBytes).toBe(0)
      expect(s.stats().inFlightBytes).toBe(0)
    }
    expect(ws.used()).toBe(0)
  })

  it('unsolicited kind 9 → protocol error, no allocation', async () => {
    const { s, sock } = await opened()
    sock.chunk(77, noise(SEND))
    await settle()
    expect(s.stats().reassemblyBytes).toBe(0)
    expect(s.stats().inFlightBytes).toBe(0)
    // A chunk for a live request before its media-start fails that request closed, unallocated.
    const p = s.mediaRequest(MID, new AbortController().signal, () => true)
    await settle()
    const id = sock.of('media-request')[0]!.reqId
    sock.chunk(id, noise(100))
    expect(await p).toEqual({ refused: 'policy-refused' })
    expect(s.stats().reassemblyBytes).toBe(0)
    expect(s.stats().inFlightBytes).toBe(0)
    // The session survives.
    expect(s.isEnded()).toBe(false)
    const q = s.mediaShow(0)
    await settle()
    sock.text({ reqId: sock.of('media-show')[0]!.reqId, type: 'media-show-ok', queued: 1 })
    expect(await q).toEqual({ queued: 1 })
  })

  it('socket close mid-delivery releases the partial charge', async () => {
    const { s, sock } = await opened()
    const ws = watched(), r = realm()
    const base = ws.used()
    const p = openRemoteMediaLease(s, rec(), ws, r.win, new AbortController().signal)
    await settle()
    const id = sock.of('media-request')[0]!.reqId
    sock.text(startOf(id, SEND + 100))
    sock.chunk(id, noise(SEND))
    await settle()
    expect(ws.used()).toBe(base + SEND + 100)
    sock.drop(1006)
    expect(await p).toEqual({ refused: 'session-rotated' })
    expect(ws.used()).toBe(base)
    expect(s.stats().inFlightBytes).toBe(0)
    expect(r.made).toHaveLength(0)
  })

  it('abort mid-delivery releases', async () => {
    const { s, sock } = await opened()
    const ws = watched(), r = realm(), ac = new AbortController()
    const p = openRemoteMediaLease(s, rec(), ws, r.win, ac.signal)
    void p.catch(() => {})
    await settle()
    const id = sock.of('media-request')[0]!.reqId
    sock.text(startOf(id, SEND + 100))
    sock.chunk(id, noise(SEND))
    await settle()
    ac.abort()
    await expect(p).rejects.toMatchObject({ name: 'AbortError' })
    expect(ws.used()).toBe(0)
    // Later chunks of the aborted request are dropped before allocation.
    sock.chunk(id, noise(100))
    await settle()
    expect(s.stats().reassemblyBytes).toBe(0)
    expect(s.stats().inFlightBytes).toBe(0)
    expect(r.made).toHaveLength(0)
  })

  it("canPlayType '' → media-codec without a request", async () => {
    const { s, sock } = await opened()
    const ws = watched(), r = realm('')
    expect(await openRemoteMediaLease(s, rec(), ws, r.win, new AbortController().signal)).toEqual({ refused: 'media-codec' })
    await settle()
    expect(sock.of('media-request')).toHaveLength(0)
    expect(ws.used()).toBe(0)
  })

  it('workspace refusal at media-start → refused, request aborted', async () => {
    const { s, sock } = await opened()
    const ws = watched(1000), r = realm()
    const p = openRemoteMediaLease(s, rec(), ws, r.win, new AbortController().signal)
    await settle()
    const id = sock.of('media-request')[0]!.reqId
    sock.text(startOf(id, 5000))
    // The workspace's refusal is the player's drop-carried-and-retry signal (player.ts `open`).
    expect(await p).toEqual({ refused: 'resource-limit' })
    sock.chunk(id, noise(5000))
    await settle()
    expect(s.stats().reassemblyBytes).toBe(0)
    expect(s.stats().inFlightBytes).toBe(0)
    expect(ws.used()).toBe(0)
    // Room for the chunks but not the Blob beside them (the 2 × bytes peak): refused the same way.
    const q = openRemoteMediaLease(s, rec(), watched(7000), r.win, new AbortController().signal)
    await settle()
    const id2 = sock.of('media-request')[1]!.reqId
    sock.text(startOf(id2, 5000))
    sock.chunk(id2, noise(5000))
    expect(await q).toEqual({ refused: 'resource-limit' })
    expect(r.made).toHaveLength(0)
  })

  it('release() revokes the Blob URL exactly once', async () => {
    const { s, sock } = await opened()
    const ws = watched(), r = realm()
    const p = openRemoteMediaLease(s, rec(), ws, r.win, new AbortController().signal)
    await settle()
    const id = sock.of('media-request')[0]!.reqId
    sock.text(startOf(id, 100))
    sock.chunk(id, noise(100))
    const lease = await p
    if ('refused' in lease) throw new Error(lease.refused)
    lease.retain(2)
    expect(lease.spans).toBe(2)
    lease.leave()
    expect(lease.spans).toBe(1)
    expect(r.revoked).toEqual([])
    lease.release()
    lease.release()
    lease.leave()
    expect(r.revoked).toEqual(['blob:x/1'])
    expect(ws.used()).toBe(0)
    expect(() => lease.retain(-1)).toThrow(RangeError)
  })
})

describe('media transport (D1, D12)', () => {
  it('media off: nothing is sent; requests answer media-unavailable', async () => {
    const { s, sock } = await opened({}, exampleManifest(false))
    const before = sock.sent.length
    expect(await s.mediaShow(0)).toEqual({ refused: 'media-unavailable' })
    expect(await s.mediaRequest(MID, new AbortController().signal, () => true)).toEqual({ refused: 'media-unavailable' })
    s.mediaCursor(1)
    s.mediaEnd()
    await settle(250)
    expect(sock.sent.length).toBe(before)
  })

  it('mediaShow: the exact message, media-show-ok, media-refused, and a stale-token retry', async () => {
    const { s, sock } = await opened()
    const p = s.mediaShow(1)
    await settle()
    const [m] = sock.of('media-show')
    expect(m).toEqual({ reqId: m!.reqId, type: 'media-show', token: s.token, unit: 1 })
    // RENEW race: refused before anything was counted; re-sent once with the fresh token.
    sock.text({ type: 'renew', token: tokenFor('j2'), jti: 'j2' })
    sock.text({ reqId: m!.reqId, type: 'refused', code: 'session-rotated' })
    await settle()
    const m2 = sock.of('media-show')[1]!
    expect(m2.reqId).not.toBe(m!.reqId)
    expect(m2.token).toBe(tokenFor('j2'))
    sock.text({ reqId: m2.reqId, type: 'media-show-ok', queued: 3 })
    expect(await p).toEqual({ queued: 3 })
    const q = s.mediaShow(0)
    await settle()
    sock.text({ reqId: sock.of('media-show')[2]!.reqId, type: 'media-refused', reason: 'media-unavailable' })
    expect(await q).toEqual({ refused: 'media-unavailable' })
  })

  it('mediaRequest: a stale-token refusal is a retry after rotation, not a refusal', async () => {
    const { s, sock } = await opened()
    const p = s.mediaRequest(MID, new AbortController().signal, () => true, 1)
    await settle()
    const [m] = sock.of('media-request')
    sock.text({ type: 'renew', token: tokenFor('j2'), jti: 'j2' })
    sock.text({ reqId: m!.reqId, type: 'refused', code: 'session-rotated' })
    await settle()
    const m2 = sock.of('media-request')[1]!
    expect(m2).toEqual({ reqId: m2.reqId, type: 'media-request', token: tokenFor('j2'), mediaId: MID, unit: 1 })
    sock.text(startOf(m2.reqId, 10))
    sock.chunk(m2.reqId, noise(10))
    const got = await p
    if ('refused' in got) throw new Error(got.refused)
    expect(got.start.bytes).toBe(10)
    expect(concat(got.chunks)).toEqual(noise(10))
    // A second stale-token refusal is a refusal.
    const q = s.mediaRequest(MID, new AbortController().signal, () => true)
    await settle()
    const n1 = sock.of('media-request')[2]!
    sock.text({ reqId: n1.reqId, type: 'refused', code: 'session-rotated' })
    await settle()
    sock.text({ reqId: sock.of('media-request')[3]!.reqId, type: 'refused', code: 'session-rotated' })
    expect(await q).toEqual({ refused: 'session-rotated' })
  })

  it('media-refused mid-delivery releases what was charged', async () => {
    const { s, sock } = await opened()
    const ws = watched(), r = realm()
    const p = openRemoteMediaLease(s, rec(), ws, r.win, new AbortController().signal)
    await settle()
    const id = sock.of('media-request')[0]!.reqId
    sock.text(startOf(id, SEND + 100))
    sock.chunk(id, noise(SEND))
    sock.text({ reqId: id, type: 'media-refused', mediaId: MID, reason: 'media-unavailable' })
    expect(await p).toEqual({ refused: 'media-unavailable' })
    expect(ws.used()).toBe(0)
    expect(s.stats().inFlightBytes).toBe(0)
  })

  it('paces media-request to 6 per minute: the 7th waits for a slot; an abort while waiting sends nothing', async () => {
    vi.useFakeTimers({ toFake: ['setTimeout', 'clearTimeout', 'setInterval', 'clearInterval', 'performance', 'Date'] })
    const { s, sock } = await opened({ now: () => Date.now() })
    const ps: Array<Promise<unknown>> = []
    for (let i = 0; i < 6; i++) ps.push(s.mediaRequest(i % 2 ? MID2 : MID, new AbortController().signal, () => true))
    const gone = new AbortController()
    const aborted = s.mediaRequest(MID, gone.signal, () => true)
    void aborted.catch(() => {})
    const seventh = s.mediaRequest(MID2, new AbortController().signal, () => true, 1)
    await settle()
    expect(sock.of('media-request')).toHaveLength(6)
    gone.abort()
    await expect(aborted).rejects.toMatchObject({ name: 'AbortError' })
    // Every one is answered (refused); a slot ages from its reply, through the server's 60 s window plus a margin.
    for (const m of sock.of('media-request')) sock.text({ reqId: m.reqId, type: 'media-refused', reason: 'media-busy' })
    await settle()
    expect(sock.of('media-request')).toHaveLength(6)
    await settle(61_000)
    expect(sock.of('media-request')).toHaveLength(6)
    await settle(2_000)
    const sent = sock.of('media-request')
    expect(sent).toHaveLength(7)
    expect(sent[6]).toMatchObject({ mediaId: MID2, unit: 1 })
    sock.text({ reqId: sent[6]!.reqId, type: 'media-refused', reason: 'media-busy' })
    expect(await seventh).toEqual({ refused: 'media-busy' })
    await Promise.all(ps)
  })

  it('mediaCursor coalesces to one per 200 ms, the latest unit on the trailing edge; the exact message', async () => {
    vi.useFakeTimers({ toFake: ['setTimeout', 'clearTimeout', 'setInterval', 'clearInterval', 'performance', 'Date'] })
    const { s, sock } = await opened({ now: () => Date.now() })
    s.mediaCursor(0)
    s.mediaCursor(1)
    s.mediaCursor(0)
    expect(sock.of('media-cursor')).toEqual([{ type: 'media-cursor', unit: 0 }])
    await settle(199)
    expect(sock.of('media-cursor')).toHaveLength(1)
    await settle(2)
    expect(sock.of('media-cursor')).toEqual([{ type: 'media-cursor', unit: 0 }, { type: 'media-cursor', unit: 0 }])
    await settle(500)
    s.mediaCursor(1)
    expect(sock.of('media-cursor')).toHaveLength(3)
    // A pending trailing cursor is cancelled by media-end.
    s.mediaCursor(0)
    s.mediaEnd()
    await settle(500)
    expect(sock.of('media-cursor')).toHaveLength(3)
    expect(sock.of('media-end')).toEqual([{ type: 'media-end' }])
  })

  it('mediaEnd settles every live media request (the server stops without a reply)', async () => {
    const { s, sock } = await opened()
    const ws = watched(), r = realm()
    const before = openRemoteMediaLease(s, rec(), ws, r.win, new AbortController().signal)
    const during = openRemoteMediaLease(s, rec(MID2), ws, r.win, new AbortController().signal)
    await settle()
    const id = sock.of('media-request')[1]!.reqId
    sock.text({ ...startOf(id, SEND + 1), mediaId: MID2 })
    sock.chunk(id, noise(SEND))
    await settle()
    s.mediaEnd()
    expect(await before).toEqual({ refused: 'media-unavailable' })
    expect(await during).toEqual({ refused: 'media-unavailable' })
    expect(ws.used()).toBe(0)
    expect(s.stats().inFlightBytes).toBe(0)
    sock.chunk(id, noise(1))
    await settle()
    expect(s.stats().reassemblyBytes).toBe(0)
  })

  it('a session that ends refuses the lease (it rejects only on abort)', async () => {
    const { s, sock } = await opened()
    const ws = watched(), r = realm()
    const p = openRemoteMediaLease(s, rec(), ws, r.win, new AbortController().signal)
    await settle()
    expect(sock.of('media-request')).toHaveLength(1)
    s.close()
    expect(await p).toEqual({ refused: 'media-unavailable' })
    expect(await openRemoteMediaLease(s, rec(), ws, r.win, new AbortController().signal)).toEqual({ refused: 'media-unavailable' })
    expect(ws.used()).toBe(0)
  })

  it('a started delivery that stalls is refused after MEDIA_CHUNK_MS; an unstarted request has no deadline', async () => {
    vi.useFakeTimers({ toFake: ['setTimeout', 'clearTimeout', 'setInterval', 'clearInterval', 'performance', 'Date'] })
    const { s, sock } = await opened({ now: () => Date.now() })
    const ws = watched(), r = realm()
    let waitingDone = false
    const waiting = s.mediaRequest(MID2, new AbortController().signal, () => true)
    void waiting.then(() => { waitingDone = true })
    const p = openRemoteMediaLease(s, rec(), ws, r.win, new AbortController().signal)
    await settle()
    const id = sock.of('media-request')[1]!.reqId
    sock.text(startOf(id, SEND + 100))
    sock.chunk(id, noise(SEND))
    await settle(29_000)
    expect(ws.used()).toBe(SEND + 100)
    await settle(3_000)
    expect(await p).toEqual({ refused: 'policy-refused' })
    expect(ws.used()).toBe(0)
    await settle(5 * 60_000)
    expect(waitingDone).toBe(false)
    s.mediaEnd()
    expect(await waiting).toEqual({ refused: 'media-unavailable' })
  })

  it('the stale-token re-send takes over its pacing slot: one slot for both sends, aged from the re-send reply', async () => {
    vi.useFakeTimers({ toFake: ['setTimeout', 'clearTimeout', 'setInterval', 'clearInterval', 'performance', 'Date'] })
    const { s, sock } = await opened({ now: () => Date.now() })
    void s.mediaRequest(MID, new AbortController().signal, () => true)
    await settle()
    sock.text({ type: 'renew', token: tokenFor('j2'), jti: 'j2' })
    sock.text({ reqId: sock.of('media-request')[0]!.reqId, type: 'refused', code: 'session-rotated' })
    await settle()
    expect(sock.of('media-request')).toHaveLength(2)
    for (let i = 0; i < 6; i++) void s.mediaRequest(MID2, new AbortController().signal, () => true)
    await settle()
    // 1 (original + re-send) + 5 = the class's 6; the last waits.
    expect(sock.of('media-request')).toHaveLength(7)
    for (const m of sock.of('media-request').slice(1)) sock.text({ reqId: m.reqId, type: 'media-refused', reason: 'media-busy' })
    await settle(63_000)
    // Every slot aged out (the original's moved to its re-send): the waiting one and five more go at once.
    for (let i = 0; i < 5; i++) void s.mediaRequest(MID2, new AbortController().signal, () => true)
    await settle()
    expect(sock.of('media-request')).toHaveLength(13)
  })

  it('resume: a request while resuming is session-rotated unsent; a media-show goes out on the new socket; slots survive', async () => {
    const { s, sock, socks, manifest } = await opened()
    const before = Array.from({ length: 6 }, () => s.mediaRequest(MID, new AbortController().signal, () => true))
    await settle()
    expect(sock.of('media-request')).toHaveLength(6)
    sock.drop(1006)
    for (const p of before) expect(await p).toEqual({ refused: 'session-rotated' })
    expect(await s.mediaRequest(MID, new AbortController().signal, () => true)).toEqual({ refused: 'session-rotated' })
    const shown = s.mediaShow(1)
    await settle(300)
    const next = socks[1]!
    next.open(); await settle()
    next.frame(1, next.of('hello')[0]!.reqId, encodeManifest(manifest))
    await settle()
    const [m] = next.of('media-show')
    expect(m).toEqual({ reqId: m!.reqId, type: 'media-show', token: s.token, unit: 1 })
    next.text({ reqId: m!.reqId, type: 'media-show-ok', queued: 2 })
    expect(await shown).toEqual({ queued: 2 })
    // The six sends before the drop still hold the media class's window (the server's sid window survives resume).
    void s.mediaRequest(MID, new AbortController().signal, () => true)
    await settle()
    expect(next.of('media-request')).toHaveLength(0)
    expect(sock.of('media-request')).toHaveLength(6)
  })

  it('a media-refused rate-limited holds the next request for a span', async () => {
    vi.useFakeTimers({ toFake: ['setTimeout', 'clearTimeout', 'setInterval', 'clearInterval', 'performance', 'Date'] })
    const { s, sock } = await opened({ now: () => Date.now() })
    const p = s.mediaRequest(MID, new AbortController().signal, () => true)
    await settle()
    sock.text({ reqId: sock.of('media-request')[0]!.reqId, type: 'media-refused', mediaId: MID, reason: 'rate-limited' })
    expect(await p).toEqual({ refused: 'rate-limited' })
    void s.mediaRequest(MID, new AbortController().signal, () => true)
    await settle(60_000)
    expect(sock.of('media-request')).toHaveLength(1)
    await settle(3_000)
    expect(sock.of('media-request')).toHaveLength(2)
  })

  it('no admission room for a chunk (8 outstanding) → resource-limit, nothing held', async () => {
    const { s, sock } = await opened()
    const ws = watched(), r = realm()
    const units = Array.from({ length: 8 }, () => s.unit(0, 1))
    for (const u of units) void u.catch(() => {})
    await settle()
    expect(s.stats().inFlight).toBe(8)
    const p = openRemoteMediaLease(s, rec(), ws, r.win, new AbortController().signal)
    await settle()
    const id = sock.of('media-request')[0]!.reqId
    sock.text(startOf(id, 100))
    sock.chunk(id, noise(100))
    expect(await p).toEqual({ refused: 'resource-limit' })
    expect(ws.used()).toBe(0)
    expect(s.stats().reassemblyBytes).toBe(0)
  })

  it('the first fragment of a multi-fragment kind 9 is not buffered when unsolicited or refused', async () => {
    const { s, sock } = await opened()
    const first = (reqId: number) => encodeMessage({ ver: 1, kind: Kind.MEDIA_CHUNK, reqId, payload: noise(2 << 20), compress: () => { throw new Error('stored') }, stored: true })[0]!
    sock.deliver(first(77).slice().buffer)
    await settle()
    expect(s.stats().reassemblyBytes).toBe(0)
    // Refused at media-start (no workspace room): its chunks are dropped before allocation.
    const p = openRemoteMediaLease(s, rec(), watched(1000), realm().win, new AbortController().signal)
    await settle()
    const id = sock.of('media-request')[0]!.reqId
    sock.text(startOf(id, 3 << 20))
    expect(await p).toEqual({ refused: 'resource-limit' })
    sock.deliver(first(id).slice().buffer)
    await settle()
    expect(s.stats().reassemblyBytes).toBe(0)
    // Admitted, the same fragment is buffered (the check above is not vacuous).
    const q = openRemoteMediaLease(s, rec(), watched(), realm().win, new AbortController().signal)
    void q
    await settle()
    const id2 = sock.of('media-request')[1]!.reqId
    sock.text(startOf(id2, 3 << 20))
    await settle()
    sock.deliver(first(id2).slice().buffer)
    await settle()
    expect(s.stats().reassemblyBytes).toBeGreaterThan(0)
    s.mediaEnd()
    await settle()
    expect(s.stats().reassemblyBytes).toBe(0)
  })
})
