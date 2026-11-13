/** Stage 3b Task 3 — print over the wire: whole-job check, caps, quota, audit, two renders at a time, PRINT_RASTER. */
import { afterAll, describe, expect, test } from 'bun:test'
import { createCanvas, loadImage } from '@napi-rs/canvas'
import { decodeRasterBody } from '../../src/stream/raster'
import { AuditTable } from '../src/audit'
import { AuditStore } from '../src/audit-store'
import type { AuthDecision } from '../src/hooks'
import { AdmissionRefused, type Pool } from '../src/pool'
import { BUN_DEFAULT_BACKPRESSURE_LIMIT, planPrint, PRINT_CONCURRENCY, PRINT_ENCODED_CAP, PRINT_JOB_MAX_DECODED, PRINT_MAX_FRAMED_BYTES, PRINT_RATE_PER_MIN, PRINT_SEND_HIGH_WATER, printSizeOk } from '../src/print'
import { Client, corpus, disposePool, pool, start } from './harness'

afterAll(() => disposePool())

const KIB48 = 48 * 1024
/** Six visible slides, 960 × 720 (pypptx-test.pptx has only one unit). */
const DECK = 'lo-tdf104788.pptx'
type Policy = AuthDecision['policy']

/** Per-sid instrumentation of the shared pool's `render`, optionally gated. */
function instrumented() {
  const real = pool()
  const by = new Map<string, { calls: number; inFlight: number; max: number; afterTeardown: number }>()
  let gate: Promise<void> | undefined, open = () => {}
  let live: (sid: string) => boolean = () => true
  let fail: { unitIndex: number; error: Error } | undefined
  let shrink = 0
  const render: Pool['render'] = async (sid, job, est) => {
    let e = by.get(sid); if (!e) by.set(sid, e = { calls: 0, inFlight: 0, max: 0, afterTeardown: 0 })
    e.calls++; if (!live(sid)) e.afterTeardown++
    e.inFlight++; e.max = Math.max(e.max, e.inFlight)
    try {
      if (gate) await gate
      if (fail && job.unitIndex === fail.unitIndex) throw fail.error
      const out = await real.render(sid, job, est)
      if (!shrink || out.kind !== 'raster' || job.scale !== 2) return out
      // As if the raw page were fractional: the renderer's ceil(raw × 2) is under 2 × the manifest size.
      const img = await loadImage(Buffer.from(out.png)), w = out.w - shrink, h = out.h - shrink
      const cv = createCanvas(w, h); cv.getContext('2d').drawImage(img, 0, 0)
      return { ...out, w, h, png: new Uint8Array(cv.toBuffer('image/png')) }
    } finally { e.inFlight-- }
  }
  const proxy = new Proxy(real, { get(t, k) { if (k === 'render') return render; const v = Reflect.get(t, k); return typeof v === 'function' ? v.bind(t) : v } })
  return {
    pool: proxy as Pool, of: (sid: string) => by.get(sid) ?? { calls: 0, inFlight: 0, max: 0, afterTeardown: 0 },
    hold() { gate = new Promise<void>((r) => { open = r }) }, release() { const o = open; gate = undefined; o() },
    liveWhen(f: (sid: string) => boolean) { live = f },
    failOn(unitIndex: number, error: Error) { fail = { unitIndex, error } },
    /** Scale-2 rasters come back `px` smaller on each side. */
    shrinkBy(px: number) { shrink = px },
    /** Waits until no render of `sid` is in flight. */
    async settled(sid: string) { while ((by.get(sid)?.inFlight ?? 0) > 0) await Bun.sleep(10) },
  }
}

async function session(o: Parameters<typeof start>[0] = {}) {
  const store = AuditStore.open(':memory:')
  const h = await start({ doc: corpus(DECK), audit: new AuditTable({ store }), ...o }), c = new Client(h.url)
  const m = await c.hello(); await c.ready()
  return { h, c, m, store, sid: m.session.sid, stop: () => { c.close(); h.stop() } }
}
const refused = (reqId: number, code: string, nums: object = {}) => ({ reqId, type: 'refused', code, ...nums })
const printOf = (c: Client, reqId: number, from: unknown, to: unknown, extra: object = {}) => c.text({ reqId, type: 'print', from, to, ...extra }, { partial: [] })
const bodyIndex = (b: Uint8Array) => new DataView(b.buffer, b.byteOffset, b.byteLength).getUint32(0, true)

describe('planPrint', () => {
  const pages = (n: number, w: number, h: number) => Array.from({ length: n }, () => ({ w, h }))
  test('scale 2 when every page and the job budget fit', () => {
    expect(planPrint(pages(3, 960, 720), 0, 2)).toEqual({ scale: 2 })
    expect(planPrint(pages(40, 1000, 1000), 0, 39)).toEqual({ scale: 2 })
  })
  test('falls back to scale 1 when a page or the job budget fails at 2', () => {
    // One page at 3000² is 144 MB decoded at scale 2 (over 64 MiB), 36 MB at 1.
    expect(planPrint([{ w: 960, h: 720 }, { w: 3000, h: 3000 }], 0, 1)).toEqual({ scale: 1 })
    // Over the side cap at 2 (5000 → 10000 > 8192), fine at 1 (side only; 5000×3000×4 = 60 MB).
    expect(planPrint([{ w: 5000, h: 3000 }], 0, 0)).toEqual({ scale: 1 })
    // 80 pages × 16 MB = 1.28 GB > 1 GiB at 2; 320 MB at 1.
    expect(planPrint(pages(80, 1000, 1000), 0, 79)).toEqual({ scale: 1 })
    expect(PRINT_JOB_MAX_DECODED).toBe(1024 * 1024 * 1024)
  })
  test('too-large when nothing fits; only the asked range counts', () => {
    expect(planPrint([{ w: 960, h: 720 }, { w: 8192, h: 8192 }], 0, 1)).toEqual({ refused: 'too-large' })
    expect(planPrint([{ w: 960, h: 720 }, { w: 8192, h: 8192 }], 0, 0)).toEqual({ scale: 2 })
    expect(planPrint(pages(300, 1000, 1000), 0, 299)).toEqual({ refused: 'too-large' })
  })
})

describe('print send bound', () => {
  test('a page sent at the high-water mark fits under Bun\'s default backpressure limit, with ~4 MiB to spare', () => {
    expect(PRINT_SEND_HIGH_WATER).toBe(4 * 1024 * 1024)
    expect(BUN_DEFAULT_BACKPRESSURE_LIMIT).toBe(16 * 1024 * 1024)
    expect(PRINT_MAX_FRAMED_BYTES).toBe(PRINT_ENCODED_CAP + 256 * 23 + 10)
    const spare = BUN_DEFAULT_BACKPRESSURE_LIMIT - PRINT_SEND_HIGH_WATER - PRINT_MAX_FRAMED_BYTES
    expect(spare).toBeGreaterThan(4 * 1024 * 1024 - 8 * 1024)
  })
})

describe('printSizeOk', () => {
  test('the renderer size bound: (unit - 1) × scale < w ≤ unit × scale, both sides', () => {
    const u = { w: 961, h: 721 }
    expect(printSizeOk({ w: 1921, h: 1441 }, u, 2)).toBe(true)
    expect(printSizeOk({ w: 1922, h: 1442 }, u, 2)).toBe(true)
    expect(printSizeOk({ w: 1920, h: 1442 }, u, 2)).toBe(false)
    expect(printSizeOk({ w: 1923, h: 1442 }, u, 2)).toBe(false)
    expect(printSizeOk({ w: 1922, h: 1440 }, u, 2)).toBe(false)
    expect(printSizeOk({ w: 961, h: 721 }, u, 1)).toBe(true)
    expect(printSizeOk({ w: 960, h: 721 }, u, 1)).toBe(false)
  })
})

/** A TEST-ONLY transport: real sends unless told to drop; a fake buffered amount. */
function transport() {
  const t = {
    buffered: 0, sock: undefined as unknown, dropKind8After: Infinity, kind8Sent: 0, kind8Attempts: 0,
    sendBinary(ws: { sendBinary(f: Uint8Array): number }, f: Uint8Array) {
      t.sock = ws
      if (f[5] === 8) { t.kind8Attempts++; if (t.kind8Sent >= t.dropKind8After) return 0; t.kind8Sent++ }
      return ws.sendBinary(f)
    },
    bufferedAmount(ws: unknown) { t.sock = ws; return t.buffered },
  }
  return t
}

describe('print over the wire', () => {
  test('happy path: print-start, ascending kind-8 pages at scale 2, print-done; audited and delivered', async () => {
    const { c, m, store, sid, stop } = await session()
    expect(m.units.count).toBe(6)
    const r = await c.print(10, 0, 1)
    expect(r.start).toEqual({ reqId: 10, type: 'print-start', pages: 2, scale: 2 })
    expect(r.pages.map(bodyIndex)).toEqual([0, 1])
    for (const [i, b] of r.pages.entries()) {
      const u = m.units.list[i]!, d = decodeRasterBody(b.subarray(4))
      expect([d.w, d.h]).toEqual([Math.ceil(u.w * 2), Math.ceil(u.h * 2)])
    }
    expect(r.end).toEqual({ reqId: 10, type: 'print-done', pages: 2 })
    for (const t of [r.start, r.end]) expect(new TextEncoder().encode(JSON.stringify(t)).length).toBeLessThanOrEqual(KIB48)
    expect(store.prints(sid)).toEqual([{ sid, pages: 2, fromUnit: 0, toUnit: 1, scale: 2, at: expect.any(Number) }])
    const ds = store.deliveries(sid).filter((d) => d.scale === 2)
    expect(ds.map((d) => d.unitIndex).sort()).toEqual([0, 1])
    // No diag or other text rides on the print reqId.
    await Bun.sleep(100)
    expect(c.msgs.filter((x) => ('text' in x && x.text.reqId === 10) || ('binary' in x && x.binary.reqId === 10))).toEqual([])
    stop()
  }, 60_000)

  test('per-job cap boundary: maxPages allowed, maxPages+1 refused print-too-large, nothing charged', async () => {
    const { c, store, sid, stop } = await session({ policy: { print: { maxPages: 2 } } })
    expect((await c.print(10, 0, 1)).end).toMatchObject({ type: 'print-done', pages: 2 })
    expect((await c.print(11, 0, 2)).end).toEqual(refused(11, 'print-too-large', { limit: 2, requested: 3 }))
    expect((await c.print(12, 0, 1)).end).toMatchObject({ type: 'print-done', pages: 2 })
    expect(store.prints(sid).map((p) => [p.fromUnit, p.toUnit])).toEqual([[0, 1], [0, 1]])
    stop()
  }, 60_000)

  test('session cap boundary: 2 + 2 allowed under 4, the next page refused print-quota-exhausted', async () => {
    const { c, h, sid, stop } = await session({ policy: { print: { maxPages: 5, maxSessionPages: 4 } } })
    expect((await c.print(10, 0, 1)).end.type).toBe('print-done')
    expect((await c.print(11, 2, 3)).end.type).toBe('print-done')
    expect((await c.print(12, 0, 0)).end).toEqual(refused(12, 'print-quota-exhausted', { limit: 4, requested: 1 }))
    expect(h.stream.audit.get(sid)!.printed).toEqual({ jobs: 2, pages: 4 })
    stop()
  }, 60_000)

  test('the total is a running sum: printing page 0 twice counts twice', async () => {
    const { c, stop } = await session({ policy: { print: { maxPages: 5, maxSessionPages: 2 } } })
    expect((await c.print(10, 0, 0)).end.type).toBe('print-done')
    expect((await c.print(11, 0, 0)).end.type).toBe('print-done')
    expect((await c.print(12, 0, 0)).end).toEqual(refused(12, 'print-quota-exhausted', { limit: 2, requested: 1 }))
    stop()
  }, 60_000)

  test('resume keeps the printed total and measures it against the new caps', async () => {
    let policy: Policy = { print: { maxPages: 5, maxSessionPages: 4 } }
    const { h, c, m, sid, stop } = await session({ authorize: async () => ({ viewerId: 'alice', policy }) })
    expect((await c.print(10, 0, 2)).end).toMatchObject({ type: 'print-done', pages: 3 })
    policy = { print: { maxPages: 5, maxSessionPages: 3 } }
    const b = new Client(h.url)
    const m2 = await b.hello({ resume: { sid, resumeProof: m.session.token } }); await b.ready()
    expect(m2.session.sid).toBe(sid)
    expect(m2.caps.print).toEqual({ maxPages: 5, maxSessionPages: 3 })
    expect((await b.print(20, 0, 0)).end).toEqual(refused(20, 'print-quota-exhausted', { limit: 3, requested: 1 }))
    b.close(); stop()
  }, 60_000)

  test('malformed ranges and a bad token are refused policy-refused; nothing charged', async () => {
    const { c, store, sid, stop } = await session({ policy: { print: { maxPages: 6, maxSessionPages: 1 } } })
    const bad: Array<[unknown, unknown, object?]> = [[2, 1], [0, 6], [-1, 0], [0, 1.5], [1.5, 2], [0, undefined], ['0', 0], [0, 0, { token: 'nope' }]]
    for (const [i, [from, to, extra]] of bad.entries()) {
      expect((await printOf(c, 50 + i, from, to, extra)).end, JSON.stringify([from, to, extra])).toEqual(refused(50 + i, 'policy-refused'))
    }
    expect(store.prints(sid)).toEqual([])
    // Nothing was charged against the 1-page session cap.
    expect((await c.print(70, 0, 0)).end.type).toBe('print-done')
    stop()
  }, 60_000)

  test('print off is refused policy-refused', async () => {
    const { c, m, store, sid, stop } = await session({ policy: { print: false } })
    expect(m.can.print).toBe(false)
    expect((await c.print(10, 0, 0)).end).toEqual(refused(10, 'policy-refused'))
    expect(store.prints(sid)).toEqual([])
    stop()
  }, 60_000)

  test('rate: 10 jobs a minute, the 11th refused rate-limited; the socket stays open and units still flow', async () => {
    let t = Date.now()
    const { c, store, sid, stop } = await session({ now: () => t })
    expect(PRINT_RATE_PER_MIN).toBe(10)
    for (let i = 0; i < 10; i++) { t += 200; expect((await c.print(10 + i, 0, 0)).end.type).toBe('print-done') }
    t += 200; expect((await c.print(30, 0, 0)).end).toEqual(refused(30, 'rate-limited'))
    expect(store.prints(sid).length).toBe(10)
    expect(c.ws.readyState).toBe(WebSocket.OPEN)
    t += 200; c.unit(40, 1, 1)
    expect((await c.unitMessages(40)).unit.unitId).toBe(1)
    // A window later, print is allowed again.
    t += 61_000; expect((await c.print(41, 0, 0)).end.type).toBe('print-done')
    stop()
  }, 120_000)

  test('one job at a time: a second print while one runs is refused rate-limited; the first completes', async () => {
    const ins = instrumented()
    const { c, sid, stop } = await session({ pool: ins.pool })
    ins.hold()
    const first = c.print(10, 0, 5)
    // A render has started only after print-start went out.
    while (ins.of(sid).calls === 0) await Bun.sleep(10)
    expect((await c.print(11, 0, 0)).end).toEqual(refused(11, 'rate-limited'))
    ins.release()
    const r = await first
    expect(r.pages.map(bodyIndex)).toEqual([0, 1, 2, 3, 4, 5])
    expect(r.end).toEqual({ reqId: 10, type: 'print-done', pages: 6 })
    // Once done, the next job runs.
    expect((await c.print(12, 0, 0)).end.type).toBe('print-done')
    stop()
  }, 60_000)

  test('cancel: no print-done, at most PRINT_CONCURRENCY pages after the cancel, the job stays charged', async () => {
    const ins = instrumented()
    const { c, store, sid, stop } = await session({ pool: ins.pool })
    ins.hold()
    c.send({ reqId: 10, type: 'print', token: c.manifest!.session.token, from: 0, to: 5 })
    expect(((await c.next((x) => 'text' in x && x.text.reqId === 10)) as { text: object }).text).toEqual({ reqId: 10, type: 'print-start', pages: 6, scale: 2 })
    c.send({ reqId: 10, type: 'print-cancel' })
    // Unknown and finished reqIds are ignored silently.
    c.send({ reqId: 99, type: 'print-cancel' })
    await Bun.sleep(100)
    ins.release()
    await ins.settled(sid); await Bun.sleep(100)
    const mine = c.msgs.filter((x) => ('text' in x && (x.text.reqId === 10 || x.text.reqId === 99)) || ('binary' in x && x.binary.reqId === 10))
    expect(mine.filter((x) => 'text' in x)).toEqual([])
    expect(mine.length).toBeLessThanOrEqual(PRINT_CONCURRENCY)
    expect(ins.of(sid).calls).toBeLessThanOrEqual(PRINT_CONCURRENCY)
    expect(store.prints(sid)).toEqual([{ sid, pages: 6, fromUnit: 0, toUnit: 5, scale: 2, at: expect.any(Number) }])
    // The sid is free for a new job.
    expect((await c.print(11, 0, 0)).end.type).toBe('print-done')
    stop()
  }, 60_000)

  test('a page failing mid-job: the pages before it, one refusal, nothing more; the job stays charged', async () => {
    for (const [reqId, error, code] of [[10, new Error('worker died'), 'policy-refused'], [20, new AdmissionRefused('session-queue'), 'rate-limited']] as const) {
      const ins = instrumented()
      const { c, store, sid, stop } = await session({ pool: ins.pool })
      ins.failOn(1, error)
      const r = await c.print(reqId, 0, 5)
      expect(r.start).toEqual({ reqId, type: 'print-start', pages: 6, scale: 2 })
      expect(r.pages.map(bodyIndex)).toEqual([0])
      expect(r.end).toEqual(refused(reqId, code))
      await ins.settled(sid); await Bun.sleep(100)
      expect(c.msgs.filter((x) => ('text' in x && x.text.reqId === reqId) || ('binary' in x && x.binary.reqId === reqId))).toEqual([])
      expect(store.prints(sid)).toEqual([{ sid, pages: 6, fromUnit: 0, toUnit: 5, scale: 2, at: expect.any(Number) }])
      // A following job runs (unit 1 still fails, so print another page).
      expect((await c.print(reqId + 1, 2, 2)).end.type).toBe('print-done')
      stop()
    }
  }, 60_000)

  test('scale 1 fallback over the wire: print-start scale 1, 1× rasters, audited and delivered at 1', async () => {
    const { c, m, store, sid, stop } = await session({ planPrint: () => ({ scale: 1 }) })
    const r = await c.print(10, 2, 3)
    expect(r.start).toEqual({ reqId: 10, type: 'print-start', pages: 2, scale: 1 })
    expect(r.pages.map(bodyIndex)).toEqual([2, 3])
    for (const [i, b] of r.pages.entries()) {
      const u = m.units.list[2 + i]!, d = decodeRasterBody(b.subarray(4))
      expect([d.w, d.h]).toEqual([u.w, u.h])
    }
    expect(r.end).toEqual({ reqId: 10, type: 'print-done', pages: 2 })
    expect(store.prints(sid)).toEqual([{ sid, pages: 2, fromUnit: 2, toUnit: 3, scale: 1, at: expect.any(Number) }])
    expect(store.deliveries(sid).map((d) => [d.unitIndex, d.scale]).sort()).toEqual([[2, 1], [3, 1]])
    stop()
  }, 60_000)

  test('fractional page size: a page inside (unit − 1) × scale < w ≤ unit × scale is sent; one outside fails the job', async () => {
    // No corpus or builder doc yields a fractional raw size (pptx, docx and xlsx layouts round),
    // so the pool returns the 1 px smaller raster ceil(raw × 2) gives for raw = unit − 0.5.
    const ins = instrumented()
    const { c, m, store, sid, stop } = await session({ pool: ins.pool })
    ins.shrinkBy(1)
    const u = m.units.list[0]!
    const r = await c.print(10, 0, 0)
    expect(r.end).toEqual({ reqId: 10, type: 'print-done', pages: 1 })
    const d = decodeRasterBody(r.pages[0]!.subarray(4))
    expect([d.w, d.h]).toEqual([u.w * 2 - 1, u.h * 2 - 1])
    expect(d.w).not.toBe(Math.ceil(u.w * 2))
    expect((u.w - 1) * 2 < d.w && d.w <= u.w * 2 && (u.h - 1) * 2 < d.h && d.h <= u.h * 2).toBe(true)
    // 2 px under is outside the contract: the job fails closed before marking or sending that page.
    ins.shrinkBy(2)
    const bad = await c.print(11, 1, 1)
    expect(bad.pages).toEqual([])
    expect(bad.end).toEqual(refused(11, 'policy-refused'))
    expect(store.deliveries(sid).filter((x) => x.unitIndex === 1)).toEqual([])
    stop()
  }, 60_000)

  test('backpressure: no page goes out above the high-water mark; drain resumes the job', async () => {
    const t = transport()
    const { h, c, stop } = await session({ transport: t as never })
    t.buffered = PRINT_SEND_HIGH_WATER + 1
    const job = c.print(10, 0, 2)
    await Bun.sleep(600)
    expect(t.kind8Attempts).toBe(0)
    // Below the mark again, but only a drain (not a timer) wakes the job.
    t.buffered = 0
    await Bun.sleep(300)
    expect(t.kind8Attempts).toBe(0)
    ;(h.stream.websocket as { drain(ws: unknown): void }).drain(t.sock)
    const r = await job
    expect(r.start).toEqual({ reqId: 10, type: 'print-start', pages: 3, scale: 2 })
    expect(r.pages.map(bodyIndex)).toEqual([0, 1, 2])
    expect(r.end).toEqual({ reqId: 10, type: 'print-done', pages: 3 })
    // Cancel breaks a drain wait: the job ends without print-done and frees the sid.
    t.buffered = PRINT_SEND_HIGH_WATER + 1
    c.send({ reqId: 11, type: 'print', token: c.manifest!.session.token, from: 0, to: 0 })
    expect(((await c.next((x) => 'text' in x && x.text.reqId === 11)) as { text: object }).text).toMatchObject({ type: 'print-start' })
    c.send({ reqId: 11, type: 'print-cancel' })
    await Bun.sleep(200)
    t.buffered = 0
    expect((await c.print(12, 1, 1)).end).toEqual({ reqId: 12, type: 'print-done', pages: 1 })
    expect(c.msgs.filter((x) => ('text' in x && x.text.reqId === 11) || ('binary' in x && x.binary.reqId === 11))).toEqual([])
    stop()
  }, 60_000)

  test('a dropped send fails the job: one refused, no print-done, nothing more rendered or sent', async () => {
    const t = transport(), ins = instrumented()
    const { c, store, sid, stop } = await session({ transport: t as never, pool: ins.pool })
    t.dropKind8After = 1
    const r = await c.print(10, 0, 5)
    expect(r.start).toEqual({ reqId: 10, type: 'print-start', pages: 6, scale: 2 })
    expect(r.pages.map(bodyIndex)).toEqual([0])
    expect(r.end).toEqual(refused(10, 'policy-refused'))
    await ins.settled(sid); await Bun.sleep(200)
    expect(c.msgs.filter((x) => ('text' in x && x.text.reqId === 10) || ('binary' in x && x.binary.reqId === 10))).toEqual([])
    // The dropped fragment was the last kind-8 send tried; at most the two in-flight pages were rendered beyond it.
    expect(t.kind8Attempts).toBe(2)
    expect(ins.of(sid).calls).toBeLessThanOrEqual(2 + PRINT_CONCURRENCY)
    expect(store.prints(sid).length).toBe(1)
    stop()
  }, 60_000)

  test('too large: an injected planPrint refusal is print-too-large with no limit; nothing charged', async () => {
    const { c, h, store, sid, stop } = await session({ planPrint: () => ({ refused: 'too-large' }) })
    expect((await c.print(10, 0, 5)).end).toEqual(refused(10, 'print-too-large'))
    expect(store.prints(sid)).toEqual([])
    expect(h.stream.audit.get(sid)!.printed).toEqual({ jobs: 0, pages: 0 })
    stop()
  }, 60_000)

  test('concurrency: a 6-page job never has more than 2 renders in flight for the sid', async () => {
    const ins = instrumented()
    const { c, sid, stop } = await session({ pool: ins.pool })
    const r = await c.print(10, 0, 5)
    expect(r.end).toEqual({ reqId: 10, type: 'print-done', pages: 6 })
    expect(r.pages.map(bodyIndex)).toEqual([0, 1, 2, 3, 4, 5])
    expect(ins.of(sid).calls).toBe(6)
    expect(ins.of(sid).max).toBe(PRINT_CONCURRENCY)
    stop()
  }, 60_000)

  test('teardown: closing the socket mid-job starts no further renders for the sid', async () => {
    const ins = instrumented()
    const { h, c, sid, stop } = await session({ pool: ins.pool })
    ins.liveWhen((s) => h.stream.table.get(s) !== undefined && h.stream.table.get(s)!.status !== 'closed')
    ins.hold()
    c.send({ reqId: 10, type: 'print', token: c.manifest!.session.token, from: 0, to: 5 })
    await c.next((x) => 'text' in x && x.text.reqId === 10 && x.text.type === 'print-start')
    c.close()
    await c.closed
    await Bun.sleep(100)
    expect(h.stream.table.get(sid)).toBeUndefined()
    ins.release()
    await ins.settled(sid); await Bun.sleep(100)
    expect(ins.of(sid).calls).toBeLessThanOrEqual(PRINT_CONCURRENCY)
    expect(ins.of(sid).afterTeardown).toBe(0)
    expect(ins.of(sid).inFlight).toBe(0)
    stop()
  }, 60_000)
})

describe('print: cancel before validation (final fix 3)', () => {
  const ofReq = (c: Client, reqId: number) => c.msgs.filter((x) => ('text' in x && x.text.reqId === reqId) || ('binary' in x && x.binary.reqId === reqId))

  test('print then print-cancel back to back: nothing is sent, charged, rated, audited or rendered; the slot is free', async () => {
    const ins = instrumented()
    // A session cap of 1 page: had the cancelled job been charged, the next print would be refused.
    const { c, h, store, sid, stop } = await session({ pool: ins.pool, policy: { print: { maxPages: 5, maxSessionPages: 1 } } })
    c.send({ reqId: 120, type: 'print', token: c.manifest!.session.token, from: 0, to: 5 })
    c.send({ reqId: 120, type: 'print-cancel' })
    await Bun.sleep(1500)
    expect(ofReq(c, 120)).toEqual([])
    expect(store.prints(sid)).toEqual([])
    expect(h.stream.audit.get(sid)!.printed).toEqual({ jobs: 0, pages: 0 })
    expect(ins.of(sid).calls).toBe(0)
    // The slot is free and nothing was charged: the one allowed page still prints.
    const r = await c.print(121, 0, 0)
    expect(r.start).toEqual({ reqId: 121, type: 'print-start', pages: 1, scale: 2 })
    expect(r.end).toEqual({ reqId: 121, type: 'print-done', pages: 1 })
    stop()
  }, 60_000)

  test('rate: a job cancelled before validation takes no print rate slot', async () => {
    const { c, stop } = await session()
    c.send({ reqId: 200, type: 'print', token: c.manifest!.session.token, from: 0, to: 0 })
    c.send({ reqId: 200, type: 'print-cancel' })
    await Bun.sleep(300)
    for (let i = 0; i < PRINT_RATE_PER_MIN; i++) expect((await c.print(201 + i, 0, 0)).end.type).toBe('print-done')
    stop()
  }, 60_000)

  test('a second print during the first one\'s validation is refused rate-limited; the first completes', async () => {
    const { c, stop } = await session()
    c.send({ reqId: 130, type: 'print', token: c.manifest!.session.token, from: 0, to: 0 })
    const second = c.print(131, 0, 0)
    expect((await second).end).toEqual(refused(131, 'rate-limited'))
    const end = await c.next((x) => 'text' in x && x.text.reqId === 130 && x.text.type === 'print-done')
    expect((end as { text: object }).text).toEqual({ reqId: 130, type: 'print-done', pages: 1 })
    stop()
  }, 60_000)

  test('a cancel for another reqId during validation does not stop the job', async () => {
    const { c, stop } = await session()
    c.send({ reqId: 140, type: 'print', token: c.manifest!.session.token, from: 0, to: 0 })
    c.send({ reqId: 141, type: 'print-cancel' })
    const end = await c.next((x) => 'text' in x && x.text.reqId === 140 && x.text.type === 'print-done')
    expect((end as { text: object }).text).toEqual({ reqId: 140, type: 'print-done', pages: 1 })
    expect(ofReq(c, 141)).toEqual([])
    stop()
  }, 60_000)

  test('a pending job refused in validation (bad token) frees the slot; a cancel after it is ignored', async () => {
    const { c, stop } = await session()
    c.send({ reqId: 150, type: 'print', token: 'not.a.token', from: 0, to: 0 })
    c.send({ reqId: 150, type: 'print-cancel' })
    await Bun.sleep(300)
    // A malformed token fails before the cancel is read: at most that one refusal.
    expect(ofReq(c, 150).map((x) => 'text' in x && x.text)).toEqual([refused(150, 'policy-refused')])
    expect((await c.print(151, 0, 0)).end.type).toBe('print-done')
    stop()
  }, 60_000)
})
