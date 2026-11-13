/**
 * R28: one unit send budget per StreamSession (the server sid's lifetime), stamped
 * on every server-charged unit send from any caller. The integration cases run a
 * real StreamSession (and the real show fetcher) against a fake server that applies
 * the real server rate gates (server/src/rate.ts SessionRates).
 */
import { afterEach, beforeEach, describe, expect, it, vi } from 'vitest'
import { gzipSync } from 'node:zlib'
import { encodeMessage } from '../../src/stream/frame'
import { encodeManifest, type Manifest } from '../../src/stream/manifest'
import { encodeRasterBody } from '../../src/stream/raster'
import { encodeUnit } from '../../src/stream/unit'
import { StreamSession } from '../../src/stream/remote/session'
import { createRemoteDoc } from '../../src/stream/remote/remote-doc'
import { createShowFetcher, type ShowFetchIo } from '../../src/stream/remote/show-fetch'

// The real server rate gates, loaded by a non-literal path (see show-fetch.test.ts).
const RATE_MODULE = '../../server/src/rate'
type Rates = { check(kind: 'unit', now: number): 'ok' | 'refuse' | 'close' }
const loadSessionRates = async () => (await import(/* @vite-ignore */ RATE_MODULE) as { SessionRates: new () => Rates }).SessionRates
vi.mock('../../server/src/notes', () => ({ NOTES_RATE_PER_MIN: 30 }))
vi.mock('../../server/src/print', () => ({ PRINT_RATE_PER_MIN: 10 }))

// ── fake server over the real rate gates ──────────────────────────────────────

const b64url = (o: object) => btoa(JSON.stringify(o)).replace(/\+/g, '-').replace(/\//g, '_').replace(/=+$/, '')
const tokenFor = (jti: string) => `h.${b64url({ jti })}.s`
function pngHeader(w: number, h: number) {
  const b = new Uint8Array(33), v = new DataView(b.buffer)
  b.set([0x89, 0x50, 0x4e, 0x47, 0x0d, 0x0a, 0x1a, 0x0a]); v.setUint32(8, 13); b.set([73, 72, 68, 82], 12)
  v.setUint32(16, w); v.setUint32(20, h); b.set([8, 6, 0, 0, 0], 24)
  return b
}
const raster = (unitId: number, w = 2, h = 2) => encodeUnit({ unitId, discriminator: 2, body: encodeRasterBody(w, h, pngHeader(w, h)) })

const COUNT = 40
const manifest = (): Manifest => ({
  v: 1, docRef: 'a'.repeat(24), kind: 'pptx', name: 'deck.pptx',
  units: { count: COUNT, list: Array.from({ length: COUNT }, () => ({ w: 960, h: 540 })) },
  layoutFingerprint: 'f'.repeat(64),
  session: { sid: '0123456789ab', token: tokenFor('j1'), expiresSec: 300, resume: 'session' },
  policy: { contentPolicy: 'visible-only', protection: 'final-raster', semantic: 'visible-only', contentPolicyId: 'visible-only@1', viewerMarkId: 'fixed' },
  can: { search: false, copy: false, print: false, textA11y: false, present: true, notes: false, links: false },
  caps: { copy: null, a11y: 'off', print: { maxPages: null, maxSessionPages: null } },
  recorder: { commands: 'ovsd-1', textMode: 'text', scalesBuckets: [1, 2], requiredFaces: [], projection: 'final-raster' },
  limits: { maxUnitPx: 8192, maxAssetBytes: 8 * 1024 * 1024, wsFrameBytes: 1 << 20, tilePx: 512 },
  notices: [],
})

class ServerSocket extends EventTarget {
  binaryType: BinaryType = 'blob'
  readyState = 0
  constructor(private readonly srv: FakeServer) {
    super()
    queueMicrotask(() => { this.readyState = 1; this.dispatchEvent(new Event('open')) })
  }
  send(d: string) { this.srv.receive(this, JSON.parse(d)) }
  close(code?: number) { this.end(code ?? 1005) }
  end(code: number) {
    if (this.readyState === 3) return
    this.readyState = 3
    queueMicrotask(() => this.dispatchEvent(Object.assign(new Event('close'), { code, reason: '' })))
  }
  deliver(data: ArrayBuffer | string) { if (this.readyState === 1) this.dispatchEvent(Object.assign(new Event('message'), { data })) }
  frame(kind: number, reqId: number, payload: Uint8Array) {
    for (const f of encodeMessage({ ver: 1, kind: kind as never, reqId, payload, compress: (b) => new Uint8Array(gzipSync(b)) })) this.deliver(f.slice().buffer)
  }
}

/** Answers every unit at once (a raster), after the real 'unit' rate gate; 'close' closes 4008. */
class FakeServer {
  readonly sends: number[] = []
  readonly unitMsgs: Record<string, any>[] = []
  readonly verdicts: { at: number; result: string }[] = []
  /** Units answered with a raster too large for the request (a validation failure → one re-send). */
  invalidOnce = new Set<number>()
  sockets: ServerSocket[] = []
  /** While set, unit answers are held (the request is in flight). */
  hold = false
  constructor(readonly rates: Rates) {}
  connect() { const s = new ServerSocket(this); this.sockets.push(s); return s }
  get live() { return this.sockets[this.sockets.length - 1]! }
  receive(sock: ServerSocket, m: Record<string, any>) {
    if (m.type === 'hello') return queueMicrotask(() => sock.frame(1, m.reqId, encodeManifest(manifest())))
    if (m.type !== 'unit') return
    const now = Date.now()
    this.sends.push(now)
    this.unitMsgs.push(m)
    const result = this.rates.check('unit', now)
    if (result !== 'ok') this.verdicts.push({ at: now, result })
    if (result === 'close') return sock.end(4008)
    if (result === 'refuse') return queueMicrotask(() => sock.deliver(JSON.stringify({ reqId: m.reqId, type: 'refused', code: 'rate-limited' })))
    if (this.hold) return
    const bad = this.invalidOnce.delete(m.unitIndex)
    queueMicrotask(() => sock.frame(3, m.reqId, bad ? raster(m.unitIndex, 4000, 4000) : raster(m.unitIndex)))
  }
}

/** Lets real async work (the reassembler's DecompressionStream) finish between fake timers. */
const realFlush = async () => { for (let i = 0; i < 6; i++) await new Promise((r) => setImmediate(r)) }
async function advance(ms: number, chunk = 50) {
  for (let left = ms; left > 0; left -= chunk) {
    await vi.advanceTimersByTimeAsync(Math.min(chunk, left))
    await realFlush()
  }
}
/** Max sends inside any 60 s window (the server's own counting). */
const maxInWindow = (times: number[]) => {
  let max = 0
  for (const s of times) max = Math.max(max, times.filter((t) => t > s - 60_000 && t <= s).length)
  return max
}

async function open(prefill = 0) {
  const srv = new FakeServer(new (await loadSessionRates())())
  for (let i = 0; i < prefill; i++) srv.rates.check('unit', Date.now())
  const p = StreamSession.open({ kind: 'stream', url: 'wss://x/stream' }, { client: () => srv.connect() as never, faces: async () => [], now: () => Date.now() })
  await advance(50, 10)
  const s = await p
  return { s, srv }
}

/** The show fetcher wired as show.ts wires it: paint through the session, the session's budget. */
function showOn(s: StreamSession) {
  const io: ShowFetchIo = {
    async paint(unit, scale, _dpr, signal, charge) {
      const lease = await s.unit(unit, scale, signal, charge)
      lease.release()
      return { width: 1, height: 1 }
    },
    budget: s.unitBudget,
  }
  const f = createShowFetcher(io, { now: () => Date.now() })
  const errors: string[] = []
  f.onError((_u, c) => errors.push(c))
  f.setBucket(1, 1)
  return { f, errors }
}

describe('R28: the unit budget is the session\'s (real StreamSession, real SessionRates)', () => {
  beforeEach(() => { vi.useFakeTimers({ toFake: ['setTimeout', 'clearTimeout', 'setInterval', 'clearInterval', 'Date'] }) })
  afterEach(() => { vi.useRealTimers() })

  it('reviewer probe: 40 slides at 400 ms, dispose, restart the show on the same session — no refusal, no close', async () => {
    const { s, srv } = await open()
    for (let show = 0; show < 2; show++) {
      const { f } = showOn(s)
      for (let i = 0; i < COUNT; i++) { f.want([i]); await advance(400) }
      f.dispose()
    }
    await advance(70_000, 500)
    expect(srv.verdicts).toEqual([])
    expect(s.isEnded()).toBe(false)
    expect(maxInWindow(srv.sends)).toBeLessThanOrEqual(48)
    s.close()
  })

  it('viewer traffic before presenting counts: 30 viewer paints, then a 40-slide show — no refusal, ≤ 48 a window', async () => {
    const { s, srv } = await open()
    await Promise.all(Array.from({ length: 30 }, (_, i) => s.unit(i, 1).then((l) => l.release())))
    expect(s.unitBudget.room()).toBe(18)
    const { f } = showOn(s)
    for (let i = 0; i < COUNT; i++) { f.want([i, (i + 1) % COUNT]); await advance(400) }
    await advance(70_000, 500)
    expect(srv.verdicts).toEqual([])
    expect(maxInWindow(srv.sends)).toBeLessThanOrEqual(48)
    f.dispose(); s.close()
  })

  it('a viewer paint over budget waits for room (never sent into a full window): 70 paints never trip 4008', async () => {
    const { s, srv } = await open()
    let done = 0
    const all = Promise.all(Array.from({ length: 70 }, (_, i) => s.unit(i % COUNT, 1).then((l) => { done++; l.release() })))
    await advance(1_000)
    expect(srv.sends).toHaveLength(48)
    expect(done).toBe(48)
    await advance(60_900, 1_000)
    expect(srv.sends).toHaveLength(48) // the client window is 62 s (a receipt margin over the server's 60 s)
    await advance(2_000)
    await all
    expect(srv.sends).toHaveLength(70)
    expect(srv.verdicts).toEqual([])
    s.close()
  })

  it('a show paint stamps exactly once; a validation re-send twice; the fetcher\'s charge never stamps again', async () => {
    const { s, srv } = await open()
    const { f } = showOn(s)
    f.want([0]); await advance(100)
    expect(srv.sends).toHaveLength(1)
    expect(s.unitBudget.room()).toBe(47)
    srv.invalidOnce.add(1)
    f.want([1]); await advance(500)
    expect(srv.sends).toHaveLength(3)
    expect(s.unitBudget.room()).toBe(45)
    f.dispose(); s.close()
  })

  it('resume mid-show keeps the log: an in-flight request\'s re-send is stamped, the restarted show stays in budget', async () => {
    const { s, srv } = await open()
    let { f } = showOn(s)
    for (let i = 0; i < 20; i++) { f.want([i]); await advance(400) }
    srv.hold = true
    f.want([20]); await advance(400)
    expect(s.unitBudget.room()).toBe(48 - 21)
    srv.hold = false
    srv.live.end(1006) // abnormal drop: the session resumes and re-sends unit 20
    await advance(1_000)
    expect(srv.sockets).toHaveLength(2)
    expect(s.unitBudget.room()).toBe(48 - 22)
    for (let i = 21; i < COUNT; i++) { f.want([i]); await advance(400) }
    f.dispose()
    ;({ f } = showOn(s))
    for (let i = 0; i < COUNT; i++) { f.want([i]); await advance(400) }
    await advance(70_000, 500)
    expect(srv.verdicts).toEqual([])
    expect(s.isEnded()).toBe(false)
    f.dispose(); s.close()
  })

  it('a refusal the log cannot explain (unknown server history) holds every unit send for a whole window — no 4008', async () => {
    // The server's window is already full when the session opens: the client cannot know it.
    const { s, srv } = await open(60)
    const { f, errors } = showOn(s)
    f.want([0]); await advance(100)
    expect(srv.verdicts.map((v) => v.result)).toEqual(['refuse'])
    expect(errors).toEqual(['rate-limited'])
    const refusedAt = srv.verdicts[0]!.at
    // Viewer paints wait too: they would supply the second and third breach.
    const viewer = s.unit(5, 1).then((l) => l.release())
    for (let i = 1; i < 30; i++) { f.want([i % 2]); await advance(2_000, 100) }
    await advance(4_000, 100)
    await viewer
    expect(srv.verdicts.map((v) => v.result)).toEqual(['refuse'])
    expect(srv.sends[1]! - refusedAt).toBeGreaterThanOrEqual(62_000)
    expect(s.isEnded()).toBe(false)
    f.dispose(); s.close()
  })

  it('a rate-limited refusal for a request aborted meanwhile still blocks the log', async () => {
    const { s, srv } = await open()
    srv.hold = true
    const ac = new AbortController()
    const p = s.unit(3, 1, ac.signal).catch((e: unknown) => e)
    await advance(100)
    ac.abort() // the fetcher aborts the old target on every navigation step
    await p
    const sent = srv.unitMsgs[0]!
    srv.live.deliver(JSON.stringify({ reqId: sent.reqId, type: 'refused', code: 'rate-limited' }))
    await advance(100)
    expect(s.unitBudget.room()).toBe(0)
    // A late refusal for an id that was never a unit send blocks nothing.
    const t = await open()
    t.srv.live.deliver(JSON.stringify({ reqId: 999, type: 'refused', code: 'rate-limited' }))
    await advance(100)
    expect(t.s.unitBudget.room()).toBe(48)
    s.close(); t.s.close()
  })

  it('createRemoteDoc exposes the session budget', async () => {
    const { s } = await open()
    const doc = createRemoteDoc(s)
    expect(doc.unitBudget).toBeDefined()
    expect(doc.unitBudget).toBe(s.unitBudget)
    s.close()
  })

  it('an aborted waiting viewer paint is never sent and gives its place back', async () => {
    const { s, srv } = await open()
    await Promise.all(Array.from({ length: 48 }, (_, i) => s.unit(i % COUNT, 1).then((l) => l.release())))
    const ac = new AbortController()
    const p = s.unit(0, 1, ac.signal).catch((e: unknown) => e)
    await advance(100)
    expect(s.unitBudget.room()).toBe(0)
    ac.abort()
    expect((await p) as Error).toBeInstanceOf(Error)
    await advance(63_000, 500)
    expect(srv.sends).toHaveLength(48)
    expect(s.unitBudget.room()).toBe(48)
    s.close()
  })
})
