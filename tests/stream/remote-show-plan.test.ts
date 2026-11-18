/**
 * Piece 4 Task 10: StreamSession.showPlan / showTrack / showMorph (plan D1) and
 * decodeBundle — kind-6 admission only for a live request holding its declared
 * grant, one send-log record per charged send, fail closed.
 */
import { afterEach, describe, expect, it, vi } from 'vitest'
import { readFileSync } from 'node:fs'
import { resolve } from 'node:path'
import { gzipSync } from 'node:zlib'
import { encodeMessage, Reassembler } from '../../src/stream/frame'
import { encodeManifest, type Manifest } from '../../src/stream/manifest'
import { encodeRasterBody } from '../../src/stream/raster'
import { encodeUnit } from '../../src/stream/unit'
import { showGrant, StreamSession, type SessionOptions } from '../../src/stream/remote/session'
import { StreamSessionError } from '../../src/stream/remote/errors'
import { UNIT_REFUSAL_BACKOFF_MS, UNIT_SPAN_MS, UnitSendLog } from '../../src/stream/remote/unit-budget'
import { decodeBundle } from '../../src/stream/remote/show-bundle'
import { encodeShowMessage, type MorphPlanRecord, type ShowPlanRecord } from '../../src/stream/show-plan'
import { createSlideShowWorkspace } from '../../src/slideshow/workspace'

const MiB = 1024 * 1024
const b64url = (o: object) => btoa(JSON.stringify(o)).replace(/\+/g, '-').replace(/\//g, '_').replace(/=+$/, '')
const tokenFor = (jti: string) => `h.${b64url({ jti })}.s`
/** The show workspace the session charges retained reply payloads to. */
let WS = createSlideShowWorkspace('desktop')
const settle = (ms = 20): Promise<unknown> => vi.isFakeTimers() ? vi.advanceTimersByTimeAsync(ms) : new Promise((r) => setTimeout(r, ms))
afterEach(() => { vi.useRealTimers(); vi.restoreAllMocks(); WS = createSlideShowWorkspace('desktop') })

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
  fragments(kind: number, reqId: number, payload: Uint8Array) {
    return encodeMessage({ ver: 1, kind: kind as never, reqId, payload, compress: (b) => new Uint8Array(gzipSync(b)) }).map((f) => f.slice().buffer)
  }
  frame(kind: number, reqId: number, payload: Uint8Array) { for (const f of this.fragments(kind, reqId, payload)) this.deliver(f) }
  of(type: string) { return this.sent.filter((m) => m.type === type) }
  drop(code: number) { this.readyState = 3; this.dispatchEvent(Object.assign(new Event('close'), { code, reason: '' })) }
}

function pngHeader(w: number, h: number, extra = 0) {
  const b = new Uint8Array(33 + extra), v = new DataView(b.buffer)
  b.set([0x89, 0x50, 0x4e, 0x47, 0x0d, 0x0a, 0x1a, 0x0a]); v.setUint32(8, 13); b.set([73, 72, 68, 82], 12)
  v.setUint32(16, w); v.setUint32(20, h); b.set([8, 6, 0, 0, 0], 24)
  for (let i = 33; i < b.length; i++) b[i] = (Math.random() * 256) | 0
  return b
}
/** A kind-6 asset payload: tag + u32 unit + u32 assetId + raster body (built raw, so hostile shapes are possible). */
function asset(tag: 2 | 3 | 4, unit: number, assetId: number, w: number, h: number, extra = 0) {
  const r = encodeRasterBody(w, h, pngHeader(w, h, extra)), out = new Uint8Array(9 + r.length), v = new DataView(out.buffer)
  out[0] = tag; v.setUint32(1, unit, true); v.setUint32(5, assetId, true); out.set(r, 9)
  return out
}

/** A two-layer plan for `unit` (background asset 1, a shape asset 2, entry asset 3). */
function plan(unit: number): ShowPlanRecord {
  return {
    v: 1, slide: { index: unit, w: 960, h: 540 }, options: { showAnimation: true, useTimings: true },
    show: { transition: undefined, timing: undefined, builds: [] },
    layers: [
      { id: 0, part: 'body', bounds: { x: 0, y: 0, w: 960, h: 540 }, z: 0, assetId: 1, decodedBytes: 960 * 540 * 4, scale: 1, background: true },
      { id: 1, key: '2', part: 'fill', bounds: { x: 100, y: 100, w: 150, h: 60 }, z: 1, assetId: 2, decodedBytes: 300 * 120 * 4, scale: 2 },
    ],
    base: [['2', { x: 100, y: 100, w: 150, h: 60, rot: 0 }]],
    steps: [{ trigger: 'click', eligibility: 'layers' }],
    triggers: [], linkRegions: [], iterate: { keys: [], groups: [] }, entry: 3, diagnostics: [],
  } as unknown as ShowPlanRecord
}
/** The bundle the server sends for `unit`: plan, entry frame, layers in z order. */
const bundle = (unit: number) => [
  encodeShowMessage({ tag: 1, plan: plan(unit) }),
  asset(3, unit, 3, 960, 540), asset(2, unit, 1, 960, 540), asset(2, unit, 2, 300, 120),
]
const decodedOf = (msgs: Uint8Array[]) => msgs[0]!.length + 2 * 960 * 540 * 4 + 300 * 120 * 4
const morphRecord = (o: Partial<MorphPlanRecord> = {}): MorphPlanRecord => ({
  v: 1, from: 0, to: 1, fromStep: 2,
  items: [{ from: { assetId: 1, geometry: { x: 0, y: 0, w: 10, h: 10, rot: 0 }, bounds: { x: 0, y: 0, w: 10, h: 10 } }, to: { assetId: 2, geometry: { x: 5, y: 5, w: 10, h: 10, rot: 0 }, bounds: { x: 5, y: 5, w: 10, h: 10 } } }] as MorphPlanRecord['items'],
  over: [], pairs: [], ...o,
})

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
const codeOf = async (p: Promise<unknown>) => (await failure(p) as StreamSessionError).code

/** Starts a showPlan for `unit` and returns its observed promise and reqId. */
async function asking(s: StreamSession, sock: FakeSocket, unit = 1, signal = new AbortController().signal) {
  const p = s.showPlan(unit, 1, 'full', signal, { workspace: WS })
  void p.catch(() => {})
  await settle()
  const req = sock.of('show-plan').at(-1)
  return { p, req, reqId: req?.reqId as number }
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

/** A decode that yields a closable bitmap of the PNG's IHDR size. */
function decoder() {
  const closed: number[] = []
  const decode = vi.fn(async (png: Uint8Array) => {
    const v = new DataView(png.buffer, png.byteOffset, png.byteLength)
    const w = v.getUint32(16), h = v.getUint32(20)
    return { width: w, height: h, close: () => { closed.push(w * h) } }
  })
  return { decode, closed }
}

describe('StreamSession show requests (scripted socket)', () => {
  it('happy path: start, messages, bundle decodes; one budget record per request', async () => {
    const record = vi.spyOn(UnitSendLog.prototype, 'record')
    const { s, sock } = await opened()
    const onSend = vi.fn()
    const p = s.showPlan(1, 2, 'full', new AbortController().signal, { workspace: WS, onSend })
    await settle()
    const req = sock.of('show-plan')[0]!
    expect(req).toEqual({ reqId: req.reqId, type: 'show-plan', token: tokenFor('j1'), unit: 1, scale: 2, motion: 'full' })
    expect(record.mock.calls).toEqual([[req.reqId]])
    expect(onSend).toHaveBeenCalledTimes(1)
    // Nothing is reserved until the start declares the reply; a diag before it is tolerated.
    expect(s.stats()).toMatchObject({ inFlight: 0, inFlightBytes: 0 })
    const msgs = bundle(1), decoded = decodedOf(msgs)
    sock.text({ reqId: req.reqId, type: 'diag', code: 'forensic-unmarked', unit: 1, reason: 'layer-capacity', assetIds: [2] })
    sock.text({ reqId: req.reqId, type: 'show-plan-start', unit: 1, messages: msgs.length, decodedBytes: decoded })
    await settle()
    // Ruling 1: the grant covers only the message in flight, so a start alone holds none.
    expect(s.stats()).toMatchObject({ inFlight: 0, inFlightBytes: 0 })
    for (const m of msgs) sock.frame(6, req.reqId, m)
    const reply = await p
    expect(reply.messages.map((m) => [...m])).toEqual(msgs.map((m) => [...m]))
    // Every message's grant was released at its completion; the retained payloads are charged
    // to the show workspace at their encoded bytes instead (ruling 5).
    expect(s.stats()).toMatchObject({ inFlight: 0, inFlightBytes: 0 })
    expect(WS.used()).toBe(msgs.reduce((n, m) => n + m.length, 0))

    const ws = WS
    const { decode, closed } = decoder()
    // decodeBundle takes the reply and releases its retained payloads.
    const b = await decodeBundle(1, reply, ws, decode)
    reply.release(); reply.release()
    await settle()
    expect(b.unit).toBe(1)
    expect(b.plan.slide.index).toBe(1)
    expect([1, 2, 3].map((id) => b.assets.bitmap(id)?.width)).toEqual([960, 300, 960])
    expect([1, 2, 3].map((id) => b.assets.bytes(id))).toEqual([960 * 540 * 4, 300 * 120 * 4, 960 * 540 * 4])
    expect(b.decodedBytes).toBe(decoded)
    // Fix round 1: the bitmaps and the plan record, exactly the bundle's decodedBytes.
    expect(ws.used()).toBe(b.decodedBytes)
    b.release(); b.release()
    expect(ws.used()).toBe(0)
    expect(closed.length).toBe(3)
    expect(b.assets.bitmap(1)).toBeUndefined()

    // A track and a Morph pair are one charged send each too.
    record.mockClear()
    const t = s.showTrack(1, 0, 1, 'reduced', new AbortController().signal, { workspace: WS }); void t.catch(() => {})
    const mo = s.showMorph(0, 1, 2, 1, 'full', new AbortController().signal, { workspace: WS }); void mo.catch(() => {})
    await settle()
    const tr = sock.of('show-track')[0]!, mr = sock.of('show-morph')[0]!
    expect(tr).toEqual({ reqId: tr.reqId, type: 'show-track', token: tokenFor('j1'), unit: 1, step: 0, scale: 1, motion: 'reduced' })
    expect(mr).toEqual({ reqId: mr.reqId, type: 'show-morph', token: tokenFor('j1'), from: 0, to: 1, fromStep: 2, scale: 1, motion: 'full' })
    expect(record.mock.calls).toEqual([[tr.reqId], [mr.reqId]])
    const frames = [asset(4, 1, 0, 960, 540), asset(4, 1, 1, 960, 540)]
    sock.text({ reqId: tr.reqId, type: 'show-track-start', unit: 1, step: 0, messages: 2, decodedBytes: 2 * 960 * 540 * 4, encodedBytes: frames[0]!.length + frames[1]!.length })
    for (const f of frames) sock.frame(6, tr.reqId, f)
    const morphMsgs = [encodeShowMessage({ tag: 5, morph: morphRecord() }), asset(2, 1, 1, 10, 10), asset(2, 1, 2, 10, 10)]
    sock.text({ reqId: mr.reqId, type: 'show-morph-start', unit: 1, from: 0, to: 1, fromStep: 2, messages: 3, decodedBytes: morphMsgs[0]!.length + 800, encodedBytes: morphMsgs.reduce((n, m) => n + m.length, 0) })
    for (const m of morphMsgs) sock.frame(6, mr.reqId, m)
    const [tl, ml] = await Promise.all([t, mo])
    expect(tl.messages.length).toBe(2)
    expect(ml.messages.length).toBe(3)
    tl.release(); ml.release()
    await settle()
    expect(s.stats()).toMatchObject({ inFlight: 0, inFlightBytes: 0 })
    expect(record).toHaveBeenCalledTimes(2)
    await usable(s, sock)
  })

  it('a reply past its declared messages is rejected and releases its grant', async () => {
    const { s, sock } = await opened()
    const { p, reqId } = await asking(s, sock)
    const msgs = bundle(1)
    // Declares the plan and entry only, then sends a tag-2 layer as well.
    sock.text({ reqId, type: 'show-plan-start', unit: 1, messages: 2, decodedBytes: decodedOf(msgs) })
    for (const m of msgs.slice(0, 3)) sock.frame(6, reqId, m)
    expect(await codeOf(p)).toBe('policy-refused')
    await settle()
    expect(s.stats()).toMatchObject({ inFlight: 0, inFlightBytes: 0, reassemblyBytes: 0 })
    expect(s.isEnded()).toBe(false)
    await usable(s, sock)
  })

  it('unsolicited kind-6 → protocol error, no decode', async () => {
    const { s, sock } = await opened()
    const push = vi.spyOn(Reassembler.prototype, 'push')
    sock.frame(6, 99, encodeShowMessage({ tag: 1, plan: plan(1) }))
    await settle()
    const results = (await Promise.all(push.mock.results.map((r) => r.value))).filter(Boolean)
    expect(results).toEqual([{ ok: false, reqId: 99, reason: 'unexpected-reqId' }])
    push.mockRestore()
    // A live request before its start admits no kind 6: it fails closed, nothing is decoded.
    const { p, reqId } = await asking(s, sock)
    sock.frame(6, reqId, encodeShowMessage({ tag: 1, plan: plan(1) }))
    expect(await codeOf(p)).toBe('policy-refused')
    // Another request's kind 6 does not reach a live one.
    const b = await asking(s, sock)
    const msgs = bundle(1)
    sock.frame(6, b.reqId + 7, msgs[0]!)
    sock.text({ reqId: b.reqId, type: 'show-plan-start', unit: 1, messages: msgs.length, decodedBytes: decodedOf(msgs) })
    for (const m of msgs) sock.frame(6, b.reqId, m)
    const reply = await b.p
    expect(reply.messages.length).toBe(4)
    reply.release()
    await usable(s, sock)
  })

  it('asset for another unit → rejected', async () => {
    const { s, sock } = await opened()
    const { p, reqId } = await asking(s, sock)
    const msgs = bundle(1)
    sock.text({ reqId, type: 'show-plan-start', unit: 1, messages: msgs.length, decodedBytes: decodedOf(msgs) })
    sock.frame(6, reqId, msgs[0]!)
    sock.frame(6, reqId, asset(3, 0, 3, 960, 540))
    expect(await codeOf(p)).toBe('policy-refused')
    await settle()
    expect(s.stats()).toMatchObject({ inFlight: 0, inFlightBytes: 0 })
    // decodeBundle holds every message to the unit as well, and leaves nothing charged.
    const ws = createSlideShowWorkspace('desktop')
    const bad = [...msgs.slice(0, 3), asset(2, 0, 2, 300, 120)]
    expect((await failure(decodeBundle(1, bad, ws, decoder().decode)) as StreamSessionError).code).toBe('policy-refused')
    expect(ws.used()).toBe(0)
    await usable(s, sock)
  })

  it('Σ decodedBytes over 128 MiB declared in show-plan-start → refused before any message', async () => {
    const { s, sock } = await opened()
    const acquire = vi.spyOn(Object.getPrototypeOf((s as any).admission), 'tryAcquire')
    const { p, reqId } = await asking(s, sock)
    sock.text({ reqId, type: 'show-plan-start', unit: 1, messages: 4, decodedBytes: 128 * MiB + 1 })
    expect(await codeOf(p)).toBe('policy-refused')
    expect(acquire).not.toHaveBeenCalled()
    // Its messages are not admitted.
    const push = vi.spyOn(Reassembler.prototype, 'push')
    sock.frame(6, reqId, bundle(1)[0]!)
    await settle()
    const results = (await Promise.all(push.mock.results.map((r) => r.value))).filter(Boolean)
    expect(results).toEqual([{ ok: false, reqId, reason: 'unexpected-reqId' }])
    expect(s.stats()).toMatchObject({ inFlight: 0, inFlightBytes: 0 })
    await usable(s, sock)
  })

  it('a message whose encoded total exceeds the start\'s encoded allowance is rejected before buffering', async () => {
    const { s, sock } = await opened()
    const push = vi.spyOn(Reassembler.prototype, 'push')
    const { p, reqId } = await asking(s, sock)
    // decodedBytes 1: the grant's encoded part is ceil(1.01) + 65536 bytes; the message is ~1.5 MB of noise.
    sock.text({ reqId, type: 'show-plan-start', unit: 1, messages: 2, decodedBytes: 1 })
    const parts = sock.fragments(6, reqId, asset(3, 1, 3, 960, 540, 1_500_000))
    expect(parts.length).toBeGreaterThan(1)
    sock.deliver(parts[0]!)
    await settle()
    const results = (await Promise.all(push.mock.results.map((r) => r.value))).filter(Boolean)
    expect(results).toEqual([{ ok: false, reqId, reason: 'encoded-cap' }])
    expect(await codeOf(p)).toBe('policy-refused')
    expect(s.stats()).toMatchObject({ inFlight: 0, inFlightBytes: 0, reassemblyBytes: 0 })
    expect(WS.used()).toBe(0)
    for (const f of parts.slice(1)) sock.deliver(f)
    await settle()
    expect(s.stats().reassemblyBytes).toBe(0)
    await usable(s, sock)
  })

  it('abort mid-reply drops reassembly and releases', async () => {
    const { s, sock } = await opened()
    const ac = new AbortController()
    const { p, reqId } = await asking(s, sock, 1, ac.signal)
    const msgs = [...bundle(1).slice(0, 2), asset(2, 1, 1, 960, 540, 1_500_000)]
    sock.text({ reqId, type: 'show-plan-start', unit: 1, messages: 4, decodedBytes: decodedOf(msgs) + 1_500_000 })
    sock.frame(6, reqId, msgs[0]!)
    sock.frame(6, reqId, msgs[1]!)
    const parts = sock.fragments(6, reqId, msgs[2]!)
    expect(parts.length).toBeGreaterThan(1)
    sock.deliver(parts[0]!)
    await settle()
    expect(s.stats().reassemblyBytes).toBeGreaterThan(0)
    expect(s.stats().inFlightBytes).toBeGreaterThan(0)
    const reason = new Error('navigated away')
    ac.abort(reason)
    expect(await failure(p)).toBe(reason)
    await settle()
    expect(s.stats()).toMatchObject({ inFlight: 0, inFlightBytes: 0, reassemblyBytes: 0 })
    for (const f of parts.slice(1)) sock.deliver(f)
    await settle()
    await usable(s, sock)
  })

  it('doc.ended rejects pending requests with ended', async () => {
    const { s, sock } = await opened()
    const a = await asking(s, sock)
    const msgs = bundle(1)
    sock.text({ reqId: a.reqId, type: 'show-plan-start', unit: 1, messages: msgs.length, decodedBytes: decodedOf(msgs) })
    sock.frame(6, a.reqId, msgs[0]!)
    const t = s.showTrack(1, 0, 1, 'full', new AbortController().signal, { workspace: WS }); void t.catch(() => {})
    await settle()
    expect(WS.used()).toBe(msgs[0]!.length)
    sock.drop(4005)
    const ended = await s.ended
    expect(ended?.code).toBe('revoked')
    expect(await failure(a.p)).toBe(ended)
    expect(await failure(t)).toBe(ended)
    expect(s.stats()).toMatchObject({ inFlight: 0, inFlightBytes: 0 })
    expect(WS.used()).toBe(0)
    expect(await failure(s.showPlan(1, 1, 'full', new AbortController().signal, { workspace: WS }))).toBe(ended)
  })

  it('rate-limited refusal backs off UNIT_REFUSAL_BACKOFF_MS', async () => {
    vi.useFakeTimers()
    const refused = vi.spyOn(UnitSendLog.prototype, 'refused')
    const { s, sock } = await opened({ now: () => Date.now() })
    const { p, reqId } = await asking(s, sock)
    sock.text({ reqId, type: 'error', code: 'rate-limited' })
    expect(await codeOf(p)).toBe('rate-limited')
    expect(refused).toHaveBeenCalledTimes(1)
    expect(s.unitBudget.room()).toBe(0)
    // The next charged send (a show request) waits out the block: nothing goes out inside the back-off.
    const next = s.showMorph(0, 1, 0, 1, 'full', new AbortController().signal, { workspace: WS }); void next.catch(() => {})
    await settle(UNIT_REFUSAL_BACKOFF_MS - 50)
    expect(sock.of('show-morph')).toEqual([])
    // Ruling 2: the hold is the whole span, not 2 s. UnitSendLog.refused() takes the short
    // UNIT_REFUSAL_BACKOFF_MS only when the log holds the server's whole limit (≥ 60 sends) and so
    // explains the refusal; the client's own caps (48 first sends, 56 re-sends) never get there, so
    // a refusal always means unknown server history: no charged send for UNIT_SPAN_MS.
    await settle(UNIT_SPAN_MS - UNIT_REFUSAL_BACKOFF_MS - 100)
    expect(sock.of('show-morph')).toEqual([])
    await settle(500)
    expect(sock.of('show-morph').length).toBe(1)
    // A refusal for a show request aborted meanwhile still blocks the log.
    const ac = new AbortController()
    const gone = s.showTrack(1, 0, 1, 'full', ac.signal, { workspace: WS }); void gone.catch(() => {})
    await settle()
    const tr = sock.of('show-track')[0]!
    ac.abort()
    sock.text({ reqId: tr.reqId, type: 'error', code: 'rate-limited' })
    await settle()
    expect(refused).toHaveBeenCalledTimes(2)
    s.close()
  })

  it('refusal codes surface typed: still, show-budget, show-limits, show-plan-timeout, morph-fallback; unknown → policy-refused', async () => {
    const { s, sock } = await opened()
    for (const code of ['still', 'show-budget', 'show-limits', 'show-plan-timeout', 'invalid-unit', 'policy-refused', 'nonsense']) {
      const { p, reqId } = await asking(s, sock)
      sock.text({ reqId, type: 'error', code })
      expect(await codeOf(p)).toBe(code === 'nonsense' ? 'policy-refused' : code)
    }
    const m = s.showMorph(0, 1, 1, 1, 'full', new AbortController().signal, { workspace: WS })
    await settle()
    sock.text({ reqId: sock.of('show-morph')[0]!.reqId, type: 'error', code: 'morph-fallback' })
    expect(await codeOf(m)).toBe('morph-fallback')
    expect(s.stats()).toMatchObject({ inFlight: 0, inFlightBytes: 0 })
    await usable(s, sock)
  })

  it('a stale-token refusal is re-sent once, uncharged: its log slot moves to the re-send', async () => {
    const record = vi.spyOn(UnitSendLog.prototype, 'record')
    const transfer = vi.spyOn(UnitSendLog.prototype, 'transfer')
    const { s, sock } = await opened()
    const onSend = vi.fn()
    const p = s.showPlan(1, 1, 'full', new AbortController().signal, { workspace: WS, onSend })
    await settle()
    const first = sock.of('show-plan')[0]!
    sock.text({ reqId: 0, type: 'renew', token: tokenFor('j2'), jti: 'j2', exp: 9e9 })
    sock.text({ reqId: first.reqId, type: 'error', code: 'session-rotated' })
    await settle()
    const second = sock.of('show-plan')[1]!
    expect(second).toMatchObject({ type: 'show-plan', token: tokenFor('j2'), unit: 1 })
    expect(second.reqId).not.toBe(first.reqId)
    expect(record.mock.calls).toEqual([[first.reqId]])
    expect(transfer.mock.calls).toEqual([[first.reqId, second.reqId]])
    expect(onSend).toHaveBeenCalledTimes(1)
    // A second stale-token refusal is not re-sent.
    sock.text({ reqId: second.reqId, type: 'error', code: 'session-rotated' })
    expect(await codeOf(p)).toBe('session-rotated')
    expect(sock.of('show-plan').length).toBe(2)
  })

  it('the start and the Morph record are held to the request', async () => {
    const { s, sock } = await opened()
    const cases: Array<[string, () => Promise<unknown>, (reqId: number) => void]> = [
      ['a start for another unit', () => s.showPlan(1, 1, 'full', new AbortController().signal, { workspace: WS }), (r) => sock.text({ reqId: r, type: 'show-plan-start', unit: 0, messages: 2, decodedBytes: 4 * MiB })],
      ['a start of another kind', () => s.showPlan(1, 1, 'full', new AbortController().signal, { workspace: WS }), (r) => sock.text({ reqId: r, type: 'show-track-start', unit: 1, step: 0, messages: 2, decodedBytes: 4 * MiB, encodedBytes: 2 * MiB })],
      ['a track start for another step', () => s.showTrack(1, 1, 1, 'full', new AbortController().signal, { workspace: WS }), (r) => sock.text({ reqId: r, type: 'show-track-start', unit: 1, step: 0, messages: 2, decodedBytes: 4 * MiB, encodedBytes: 2 * MiB })],
      ['zero messages', () => s.showPlan(1, 1, 'full', new AbortController().signal, { workspace: WS }), (r) => sock.text({ reqId: r, type: 'show-plan-start', unit: 1, messages: 0, decodedBytes: 4 * MiB })],
      ['more messages than a plan can carry', () => s.showPlan(1, 1, 'full', new AbortController().signal, { workspace: WS }), (r) => sock.text({ reqId: r, type: 'show-plan-start', unit: 1, messages: 131, decodedBytes: 4 * MiB })],
      ['a second start', () => s.showPlan(1, 1, 'full', new AbortController().signal, { workspace: WS }), (r) => { sock.text({ reqId: r, type: 'show-plan-start', unit: 1, messages: 4, decodedBytes: 4 * MiB }); sock.text({ reqId: r, type: 'show-plan-start', unit: 1, messages: 4, decodedBytes: 4 * MiB }) }],
      ['a start over 64 KiB', () => s.showPlan(1, 1, 'full', new AbortController().signal, { workspace: WS }), (r) => sock.text({ reqId: r, type: 'show-plan-start', unit: 1, messages: 4, decodedBytes: 4 * MiB, pad: 'x'.repeat(65536) })],
      ['an unexpected text type', () => s.showPlan(1, 1, 'full', new AbortController().signal, { workspace: WS }), (r) => sock.text({ reqId: r, type: 'show-plan-done' })],
      ['a plan reply whose first message is not the plan', () => s.showPlan(1, 1, 'full', new AbortController().signal, { workspace: WS }), (r) => { sock.text({ reqId: r, type: 'show-plan-start', unit: 1, messages: 4, decodedBytes: 8 * MiB }); sock.frame(6, r, asset(3, 1, 3, 960, 540)) }],
      ['a track reply carrying a layer', () => s.showTrack(1, 0, 1, 'full', new AbortController().signal, { workspace: WS }), (r) => { sock.text({ reqId: r, type: 'show-track-start', unit: 1, step: 0, messages: 2, decodedBytes: 8 * MiB, encodedBytes: 2 * MiB }); sock.frame(6, r, asset(2, 1, 0, 960, 540)) }],
      ['a Morph start for another pair', () => s.showMorph(0, 1, 2, 1, 'full', new AbortController().signal, { workspace: WS }), (r) => sock.text({ reqId: r, type: 'show-morph-start', unit: 1, from: 0, to: 1, fromStep: 1, messages: 3, decodedBytes: 4 * MiB, encodedBytes: 2 * MiB })],
      ['a Morph record for another fromStep', () => s.showMorph(0, 1, 1, 1, 'full', new AbortController().signal, { workspace: WS }), (r) => { sock.text({ reqId: r, type: 'show-morph-start', unit: 1, from: 0, to: 1, fromStep: 1, messages: 3, decodedBytes: 4 * MiB, encodedBytes: 2 * MiB }); sock.frame(6, r, encodeShowMessage({ tag: 5, morph: morphRecord({ fromStep: 2 }) })) }],
      ['a Morph record for another from', () => s.showMorph(0, 1, 2, 1, 'full', new AbortController().signal, { workspace: WS }), (r) => { sock.text({ reqId: r, type: 'show-morph-start', unit: 1, from: 0, to: 1, fromStep: 2, messages: 3, decodedBytes: 4 * MiB, encodedBytes: 2 * MiB }); sock.frame(6, r, encodeShowMessage({ tag: 5, morph: morphRecord({ from: 1, to: 0 }) })) }],
      ['a composited Morph whose frame count is not the declared reply', () => s.showMorph(0, 1, 2, 1, 'full', new AbortController().signal, { workspace: WS }), (r) => { sock.text({ reqId: r, type: 'show-morph-start', unit: 1, from: 0, to: 1, fromStep: 2, messages: 3, decodedBytes: 8 * MiB, encodedBytes: 2 * MiB }); sock.frame(6, r, encodeShowMessage({ tag: 5, morph: morphRecord({ items: [], composited: { fps: 15, count: 5 } }) })) }],
    ]
    for (const [what, ask, hostile] of cases) {
      const before = sock.sent.length
      const p = ask(); void p.catch(() => {})
      await settle()
      const reqId = sock.sent.slice(before).find((m) => m.type.startsWith('show-'))!.reqId
      hostile(reqId)
      expect({ what, code: await codeOf(p) }).toEqual({ what, code: 'policy-refused' })
      await settle()
      expect({ what, stats: s.stats() }).toMatchObject({ what, stats: { inFlight: 0, inFlightBytes: 0 } })
    }
    await usable(s, sock)
  })

  it('refuses without sending: no present, a bad scale or motion, a unit out of range or hidden, a bad step or pair', async () => {
    const { s, sock } = await opened()
    const sig = () => new AbortController().signal
    expect(await codeOf(s.showPlan(3, 1, 'full', sig(), { workspace: WS }))).toBe('invalid-unit')
    expect(await codeOf(s.showPlan(2, 1, 'full', sig(), { workspace: WS }))).toBe('invalid-unit')
    expect(await codeOf(s.showPlan(0.5, 1, 'full', sig(), { workspace: WS }))).toBe('invalid-unit')
    expect(await codeOf(s.showTrack(1, 64, 1, 'full', sig(), { workspace: WS }))).toBe('invalid-unit')
    expect(await codeOf(s.showTrack(1, -1, 1, 'full', sig(), { workspace: WS }))).toBe('invalid-unit')
    expect(await codeOf(s.showMorph(1, 1, 0, 1, 'full', sig(), { workspace: WS }))).toBe('invalid-unit')
    expect(await codeOf(s.showMorph(2, 1, 0, 1, 'full', sig(), { workspace: WS }))).toBe('invalid-unit')
    expect(await codeOf(s.showMorph(0, 1, 65, 1, 'full', sig(), { workspace: WS }))).toBe('invalid-unit')
    expect(await codeOf(s.showPlan(1, 3 as never, 'full', sig(), { workspace: WS }))).toBe('policy-refused')
    expect(await codeOf(s.showPlan(1, 1, 'slow' as never, sig(), { workspace: WS }))).toBe('policy-refused')
    const ac = new AbortController(); ac.abort(new Error('gone'))
    expect(((await failure(s.showPlan(1, 1, 'full', ac.signal, { workspace: WS }))) as Error).message).toBe('gone')
    expect(sock.sent.filter((m) => String(m.type).startsWith('show-'))).toEqual([])
    const off = await opened({}, { can: { ...exampleManifest().can, present: false } })
    expect(await codeOf(off.s.showPlan(1, 1, 'full', sig(), { workspace: WS }))).toBe('policy-refused')
    expect(off.sock.sent.filter((m) => String(m.type).startsWith('show-'))).toEqual([])
  })

  it('two concurrent ≥32 MiB-declared replies both succeed, each message granted in turn within the 96 MiB total', async () => {
    const { s, sock } = await opened()
    const a = await asking(s, sock), b = await asking(s, sock)
    const msgs = bundle(1)
    // Each declares 40 MiB: one message's grant is 72 MiB, two at once would be 144 MiB.
    expect(showGrant(40 * MiB) * 2).toBeGreaterThan(96 * MiB)
    sock.text({ reqId: a.reqId, type: 'show-plan-start', unit: 1, messages: msgs.length, decodedBytes: 40 * MiB })
    sock.text({ reqId: b.reqId, type: 'show-plan-start', unit: 1, messages: msgs.length, decodedBytes: 40 * MiB })
    await settle()
    expect(s.stats()).toMatchObject({ inFlight: 0, inFlightBytes: 0 })
    // Interleaved messages: each completes (and frees its grant) before the next one starts.
    for (const m of msgs) { sock.frame(6, a.reqId, m); sock.frame(6, b.reqId, m) }
    const [ra, rb] = await Promise.all([a.p, b.p])
    expect([ra.messages.length, rb.messages.length]).toEqual([4, 4])
    expect(s.stats()).toMatchObject({ inFlight: 0, inFlightBytes: 0 })
    expect(WS.used()).toBe(2 * msgs.reduce((n, m) => n + m.length, 0))
    ra.release(); rb.release()
    expect(WS.used()).toBe(0)
    await usable(s, sock)
  })

  it('an admission refusal surfaces as `admission` and does not trigger UNIT_REFUSAL_BACKOFF_MS', async () => {
    const refused = vi.spyOn(UnitSendLog.prototype, 'refused')
    const { s, sock } = await opened()
    const a = await asking(s, sock), b = await asking(s, sock)
    const big = asset(3, 1, 3, 960, 540, 1_500_000)
    sock.text({ reqId: a.reqId, type: 'show-plan-start', unit: 1, messages: 2, decodedBytes: 40 * MiB })
    sock.text({ reqId: b.reqId, type: 'show-plan-start', unit: 1, messages: 2, decodedBytes: 40 * MiB })
    sock.frame(6, a.reqId, encodeShowMessage({ tag: 1, plan: plan(1) }))
    // a's second message is in flight (one fragment of several), holding its 72 MiB grant…
    const parts = sock.fragments(6, a.reqId, big)
    sock.deliver(parts[0]!)
    await settle()
    expect(s.stats().inFlightBytes).toBe(showGrant(40 * MiB))
    // …so b's first message cannot be admitted: a local, transient refusal.
    sock.frame(6, b.reqId, encodeShowMessage({ tag: 1, plan: plan(1) }))
    expect(await codeOf(b.p)).toBe('admission')
    expect(refused).not.toHaveBeenCalled()
    expect(s.unitBudget.room()).toBeGreaterThan(0)
    // The next charged send goes out at once (no server rate pause).
    const c = s.showPlan(1, 1, 'full', new AbortController().signal, { workspace: WS }); void c.catch(() => {})
    await settle()
    expect(sock.of('show-plan').length).toBe(3)
    // a completes normally.
    for (const f of parts.slice(1)) sock.deliver(f)
    const ra = await a.p
    expect(ra.messages.length).toBe(2)
    ra.release()
    s.close()
  })

  it('a drop with the request sent rejects session-rotated (never re-issued)', async () => {
    const { s, sock } = await opened()
    const { p, reqId } = await asking(s, sock)
    sock.text({ reqId, type: 'show-plan-start', unit: 1, messages: 4, decodedBytes: 8 * MiB })
    await settle()
    sock.drop(1006)
    expect(await codeOf(p)).toBe('session-rotated')
    expect(s.stats()).toMatchObject({ inFlight: 0, inFlightBytes: 0 })
    s.close()
  })
})

describe('track and Morph starts declare the D23 window (final review finding 1)', () => {
  const WINDOW = 8
  const sumOf = (msgs: Uint8Array[]) => msgs.reduce((n, m) => n + m.length, 0)
  const frames = (n: number, w: number, h: number) => Array.from({ length: n }, (_, i) => asset(4, 1, i, w, h))

  it('a 90-frame 960×720 track (Σ decoded over 128 MiB) is admitted on its window bound and its encoded bytes', async () => {
    const { s, sock } = await opened()
    const t = s.showTrack(1, 0, 1, 'full', new AbortController().signal, { workspace: WS }); void t.catch(() => {})
    await settle()
    const r = sock.of('show-track').at(-1)!.reqId
    const msgs = frames(90, 960, 720)
    expect(90 * 960 * 720 * 4).toBeGreaterThan(128 * MiB)
    sock.text({ reqId: r, type: 'show-track-start', unit: 1, step: 0, messages: 90, decodedBytes: WINDOW * 960 * 720 * 4, encodedBytes: sumOf(msgs) })
    for (const m of msgs) sock.frame(6, r, m)
    const reply = await t
    expect(reply.messages.length).toBe(90)
    reply.release()
    await settle()
    expect(s.stats()).toMatchObject({ inFlight: 0, inFlightBytes: 0 })
    await usable(s, sock)
  })

  it('a composited Morph whose Σ is over 128 MiB is admitted on its window bound plus its record', async () => {
    const { s, sock } = await opened()
    const mo = s.showMorph(0, 1, 2, 1, 'full', new AbortController().signal, { workspace: WS }); void mo.catch(() => {})
    await settle()
    const r = sock.of('show-morph').at(-1)!.reqId
    const rec = encodeShowMessage({ tag: 5, morph: morphRecord({ items: [], composited: { fps: 15, count: 60 } }) })
    const msgs = [rec, ...frames(60, 960, 720)]
    expect(60 * 960 * 720 * 4).toBeGreaterThan(128 * MiB)
    sock.text({ reqId: r, type: 'show-morph-start', unit: 1, from: 0, to: 1, fromStep: 2, messages: 61, decodedBytes: WINDOW * 960 * 720 * 4 + rec.length, encodedBytes: sumOf(msgs) })
    for (const m of msgs) sock.frame(6, r, m)
    const reply = await mo
    expect(reply.messages.length).toBe(61)
    reply.release()
    await settle()
    expect(s.stats()).toMatchObject({ inFlight: 0, inFlightBytes: 0 })
  })

  it('hostile track and Morph starts are refused before any message', async () => {
    const { s, sock } = await opened()
    const track = () => s.showTrack(1, 0, 1, 'full', new AbortController().signal, { workspace: WS })
    const morph = () => s.showMorph(0, 1, 2, 1, 'full', new AbortController().signal, { workspace: WS })
    const T = { type: 'show-track-start', unit: 1, step: 0, messages: 2, decodedBytes: 4 * MiB, encodedBytes: 2 * MiB }
    const M = { type: 'show-morph-start', unit: 1, from: 0, to: 1, fromStep: 2, messages: 3, decodedBytes: 4 * MiB, encodedBytes: 2 * MiB }
    const P = { type: 'show-plan-start', unit: 1, messages: 4, decodedBytes: 4 * MiB }
    const plan = () => s.showPlan(1, 1, 'full', new AbortController().signal, { workspace: WS })
    const { encodedBytes: _t, ...noEncT } = T
    const { encodedBytes: _m, ...noEncM } = M
    const cases: Array<[string, () => Promise<unknown>, Record<string, unknown>]> = [
      ['a track start over 128 MiB encoded', track, { ...T, messages: 90, encodedBytes: 128 * MiB + 1 }],
      ['a track start over messages × 8 MiB encoded', track, { ...T, encodedBytes: 16 * MiB + 1 }],
      ['a track start over 128 MiB decoded', track, { ...T, decodedBytes: 128 * MiB + 1 }],
      ['a track start with no encodedBytes', track, noEncT],
      ['a track start with a zero encodedBytes', track, { ...T, encodedBytes: 0 }],
      ['a track start with a fractional encodedBytes', track, { ...T, encodedBytes: 1.5 }],
      ['a track start with an extra key', track, { ...T, extra: 1 }],
      ['a Morph start over 128 MiB encoded', morph, { ...M, messages: 61, encodedBytes: 128 * MiB + 1 }],
      ['a Morph start over messages × 8 MiB encoded', morph, { ...M, encodedBytes: 24 * MiB + 1 }],
      ['a Morph start over 128 MiB decoded', morph, { ...M, decodedBytes: 128 * MiB + 1 }],
      ['a Morph start with no encodedBytes', morph, noEncM],
      ['a Morph start with an extra key', morph, { ...M, scale: 1 }],
      ['a plan start with encodedBytes (not its key set)', plan, { ...P, encodedBytes: 2 * MiB }],
      ['a plan start with an extra key', plan, { ...P, step: 0 }],
    ]
    const push = vi.spyOn(Reassembler.prototype, 'push')
    for (const [what, ask, start] of cases) {
      const before = sock.sent.length
      const p = ask(); void p.catch(() => {})
      await settle()
      const reqId = sock.sent.slice(before).find((m) => m.type.startsWith('show-'))!.reqId
      sock.text({ reqId, ...start })
      expect({ what, code: await codeOf(p) }).toEqual({ what, code: 'policy-refused' })
      await settle()
      expect({ what, stats: s.stats() }).toMatchObject({ what, stats: { inFlight: 0, inFlightBytes: 0 } })
    }
    expect(push).not.toHaveBeenCalled()
    await usable(s, sock)
  })

  it('a track reply past its declared encoded bytes is refused and releases what it held', async () => {
    const { s, sock } = await opened()
    const t = s.showTrack(1, 0, 1, 'full', new AbortController().signal, { workspace: WS }); void t.catch(() => {})
    await settle()
    const r = sock.of('show-track').at(-1)!.reqId
    const msgs = frames(3, 960, 720)
    sock.text({ reqId: r, type: 'show-track-start', unit: 1, step: 0, messages: 3, decodedBytes: 3 * 960 * 720 * 4, encodedBytes: sumOf(msgs) - 1 })
    for (const m of msgs) sock.frame(6, r, m)
    expect(await codeOf(t)).toBe('policy-refused')
    await settle()
    expect(WS.used()).toBe(0)
    expect(s.stats()).toMatchObject({ inFlight: 0, inFlightBytes: 0 })
  })
})

/** The recorded harness sessions (Task 9, `tests/fixtures/show-plan/*.bin`; format in scripts/record-show-fixtures.ts). */
interface Rec { op: number; unit: number; a: number; b: number; bytes: Uint8Array }
function records(name: string): Rec[] {
  const buf = new Uint8Array(readFileSync(resolve(__dirname, '../fixtures/show-plan', `${name}.bin`)))
  const v = new DataView(buf.buffer, buf.byteOffset, buf.byteLength)
  const out: Rec[] = []
  for (let at = 0; at < buf.length;) {
    const len = v.getUint32(at + 13, true)
    out.push({ op: v.getUint8(at), unit: v.getUint32(at + 1, true), a: v.getUint32(at + 5, true), b: v.getUint32(at + 9, true), bytes: buf.subarray(at + 17, at + 17 + len) })
    at += 17 + len
  }
  return out
}
/** Real replies grouped by request: op 1 plan (unit), op 2 track (unit, step a), op 3 Morph (to = unit, from a, fromStep b). */
function replies(name: string) {
  const out = new Map<string, { op: number; unit: number; a: number; b: number; messages: Uint8Array[] }>()
  for (const r of records(name)) {
    if (r.op < 1 || r.op > 3) continue
    const key = `${r.op}:${r.unit}:${r.op === 1 ? 0 : r.a}:${r.op === 3 ? r.b : 0}`
    if (!out.has(key)) out.set(key, { op: r.op, unit: r.unit, a: r.a, b: r.b, messages: [] })
    out.get(key)!.messages.push(r.bytes)
  }
  return [...out.values()]
}
const rgbaOf = (msgs: Uint8Array[]) => msgs.reduce((n, m) => (m[0] === 1 || m[0] === 5 ? n + m.length : n + new DataView(m.buffer, m.byteOffset).getUint32(9, true) * new DataView(m.buffer, m.byteOffset).getUint32(13, true) * 4), 0)

describe('recorded server replies', () => {
  for (const name of ['fly', 'morph']) {
    it(`${name}: every recorded reply passes the session and every slide bundle decodes`, async () => {
      const all = replies(name)
      expect(all.some((r) => r.op === 1)).toBe(true)
      const units = Math.max(...all.map((r) => r.unit), ...all.filter((r) => r.op === 3).map((r) => r.a)) + 1
      const { s, sock } = await opened({}, { units: { count: units, list: Array.from({ length: units }, () => ({ w: 960, h: 540 })) } })
      for (const r of all) {
        const sig = new AbortController().signal
        const p = r.op === 1 ? s.showPlan(r.unit, 1, 'full', sig, { workspace: WS }) : r.op === 2 ? s.showTrack(r.unit, r.a, 1, 'full', sig, { workspace: WS }) : s.showMorph(r.a, r.unit, r.b, 1, 'full', sig, { workspace: WS })
        void p.catch(() => {})
        await settle()
        const kind = r.op === 1 ? 'plan' : r.op === 2 ? 'track' : 'morph'
        const req = sock.of(`show-${kind}`).at(-1)!
        const echo = r.op === 1 ? {} : r.op === 2 ? { step: r.a } : { from: r.a, to: r.unit, fromStep: r.b }
        sock.text({ reqId: req.reqId, type: `show-${kind}-start`, unit: r.unit, ...echo, messages: r.messages.length, decodedBytes: rgbaOf(r.messages), ...(r.op === 1 ? {} : { encodedBytes: r.messages.reduce((n, m) => n + m.length, 0) }) })
        for (const m of r.messages) sock.frame(6, req.reqId, m)
        const reply = await p
        expect(reply.messages.length).toBe(r.messages.length)
        if (r.op === 1) {
          const ws = createSlideShowWorkspace('desktop'), { decode } = decoder()
          const b = await decodeBundle(r.unit, reply.messages, ws, decode)
          expect(ws.used()).toBeGreaterThan(0)
          b.release()
          expect(ws.used()).toBe(0)
        }
        reply.release()
      }
      await settle()
      expect(s.stats()).toMatchObject({ inFlight: 0, inFlightBytes: 0 })
      s.close()
    })
  }
})

describe('decodeBundle', () => {
  const ws = () => createSlideShowWorkspace('desktop')
  const cases: Array<[string, Uint8Array[]]> = [
    ['no messages', []],
    ['a first message that is not the plan', bundle(1).slice(1)],
    ['a duplicate assetId', [...bundle(1).slice(0, 3), asset(2, 1, 1, 960, 540)]],
    ['a layer asset missing', bundle(1).slice(0, 3)],
    ['an extra asset', [...bundle(1), asset(2, 1, 9, 10, 10)]],
    ['an entry frame that is not the plan\'s', [bundle(1)[0]!, asset(3, 1, 7, 960, 540), ...bundle(1).slice(2)]],
    ['an asset larger than its layer', [...bundle(1).slice(0, 3), asset(2, 1, 2, 301, 120)]],
    ['a keyframe in a slide bundle', [...bundle(1), asset(4, 1, 8, 10, 10)]],
    ['a plan for another unit', [encodeShowMessage({ tag: 1, plan: plan(0) }), ...bundle(1).slice(1)]],
  ]
  for (const [what, msgs] of cases) {
    it(`rejects ${what}, charging nothing`, async () => {
      const w = ws(), { decode, closed } = decoder()
      expect((await failure(decodeBundle(1, msgs, w, decode)) as StreamSessionError).code).toBe('policy-refused')
      expect(w.used()).toBe(0)
      expect(closed.length).toBe(decode.mock.calls.length)
    })
  }

  it('a bitmap not the asset\'s size, or a failed decode, rejects and releases everything', async () => {
    const w = ws(), closed: number[] = []
    let n = 0
    const lying = async () => ({ width: n++ === 1 ? 5 : 960, height: 540, close: () => { closed.push(1) } })
    expect((await failure(decodeBundle(1, bundle(1), w, lying)) as StreamSessionError).code).toBe('policy-refused')
    expect(w.used()).toBe(0)
    expect(closed.length).toBe(2)
    const throwing = async () => { throw new Error('bad png') }
    expect((await failure(decodeBundle(1, bundle(1), w, throwing)) as StreamSessionError).code).toBe('policy-refused')
    expect(w.used()).toBe(0)
  })

  it('the plan record stays charged for the bundle\'s life: bitmaps that fit without it are refused', async () => {
    const msgs = bundle(1), bitmaps = 2 * 960 * 540 * 4 + 300 * 120 * 4
    const tight = createSlideShowWorkspace('desktop', { limit: bitmaps + msgs[0]!.length - 1 })
    const { decode, closed } = decoder()
    expect((await failure(decodeBundle(1, msgs, tight, decode)) as StreamSessionError).code).toBe('workspace-full')
    expect(tight.used()).toBe(0)
    expect(closed.length).toBe(decode.mock.calls.length)
    const exact = createSlideShowWorkspace('desktop', { limit: bitmaps + msgs[0]!.length })
    const b = await decodeBundle(1, msgs, exact, decoder().decode)
    expect(exact.used()).toBe(b.decodedBytes)
    b.release(); b.release()
    expect(exact.used()).toBe(0)
  })

  it('several retained bundles: the workspace holds the sum of their decodedBytes', async () => {
    const w = ws()
    const a = await decodeBundle(1, bundle(1), w, decoder().decode)
    const b = await decodeBundle(0, bundle(0), w, decoder().decode)
    expect(w.used()).toBe(a.decodedBytes + b.decodedBytes)
    a.release()
    expect(w.used()).toBe(b.decodedBytes)
    b.release()
    expect(w.used()).toBe(0)
  })

  it('no workspace room → workspace-full, nothing kept', async () => {
    const w = createSlideShowWorkspace('desktop', { limit: 960 * 540 * 4 + 10 })
    const { decode, closed } = decoder()
    expect((await failure(decodeBundle(1, bundle(1), w, decode)) as StreamSessionError).code).toBe('workspace-full')
    expect(w.used()).toBe(0)
    expect(closed.length).toBe(decode.mock.calls.length)
  })
})
