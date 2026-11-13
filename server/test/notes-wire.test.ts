/** Stage 3c Task 2 — speaker notes over the wire: presenters only, unit rules, the 'notes' rate, audit, NOTES_RASTER. */
import { afterAll, describe, expect, test } from 'bun:test'
import { buildDocx, buildPptxDeck } from '../../src/testdata/ooxml-builders'
import { OfficePackage } from '../../src/core/zip'
import { decodeRasterBody } from '../../src/stream/raster'
import { AuditTable } from '../src/audit'
import { AuditStore } from '../src/audit-store'
import type { AuthDecision } from '../src/hooks'
import { serverBackend } from '../src/backend'
import { decodePng } from '../src/mark/png'
import { drawNotes, layoutNotes, NOTES_CONTINUATION, NOTES_MAX_PX, NOTES_RATE_PER_MIN, readSlideNotes } from '../src/notes'
import type { Pool } from '../src/pool'
import { hiddenSlideDeck } from './fixtures'
import { Client, corpus, disposePool, pool, start } from './harness'

afterAll(() => disposePool())

/** Six slides, 960 × 720; every slide has a notes part, units 4 and 5 carry notes text. */
const DECK = 'lo-tdf104788.pptx'
type Policy = AuthDecision['policy']
const presenter = (policy: Policy = {}): AuthDecision => ({ viewerId: 'alice', role: 'presenter', policy })
const audience = (policy: Policy = {}): AuthDecision => ({ viewerId: 'alice', policy })
const error = (reqId: number, code: string) => ({ reqId, type: 'error', code })

/** `tick()` advances the server clock past the per-second message burst (many refusals in a row). */
async function session(o: Parameters<typeof start>[0] & { decision?: AuthDecision } = {}) {
  const store = AuditStore.open(':memory:')
  const decision = o.decision ?? presenter()
  let t = Date.now()
  const h = await start({ doc: corpus(DECK), audit: new AuditTable({ store }), authorize: async () => decision, now: () => t, ...o }), c = new Client(h.url)
  const m = await c.hello(); await c.ready()
  return { h, c, m, store, sid: m.session.sid, tick: () => { t += 120 }, stop: () => { c.close(); h.stop() } }
}
/** A unit round trip after `reqId`'s answer: everything for `reqId` has arrived by then. */
async function fence(c: Client, reqId = 9000) { c.unit(reqId, 0, 1); await c.unitMessages(reqId) }
const framesFor = (c: Client, reqId: number) => c.msgs.filter((m) => 'binary' in m && m.binary.reqId === reqId)
function body(b: Uint8Array) {
  const unitIndex = new DataView(b.buffer, b.byteOffset, b.byteLength).getUint32(0, true)
  return { unitIndex, ...decodeRasterBody(b.subarray(4)) }
}
/** A refusal: exactly one `error`, no notes-start, no frame, no audit row. */
async function expectRefused(s: Awaited<ReturnType<typeof session>>, unit: unknown, scale: unknown, code: string) {
  s.tick()
  const r = await s.c.notes(unit, scale)
  expect(r.end).toEqual(error(r.reqId, code))
  expect(r.start).toBeUndefined()
  expect(r.extra).toEqual([])
  s.tick()
  await fence(s.c, 9000 + r.reqId)
  expect(framesFor(s.c, r.reqId)).toEqual([])
  expect(s.c.msgs.filter((m) => 'text' in m && m.text.reqId === r.reqId)).toEqual([])
  expect(s.store.notesDeliveries(s.sid)).toEqual([])
}

describe('notes over the wire: the presenter path', () => {
  test.each([[1], [2]] as const)('unit 4 at scale %p: notes-start then one kind-7 image 720 × scale wide; one audit row', async (scale) => {
    const s = await session()
    expect(s.m.can.notes).toBe(true)
    const r = await s.c.notes(4, scale)
    expect(r.start).toEqual({ reqId: r.reqId, type: 'notes-start', w: 720 * scale, h: expect.any(Number), scale })
    expect(r.extra).toEqual([])
    const b = body(r.body!)
    expect(b.unitIndex).toBe(4)
    expect([b.w, b.h]).toEqual([720 * scale, r.start!.h])
    const png = decodePng(b.png)
    expect([png.w, png.h]).toEqual([720 * scale, r.start!.h])
    const rows = s.store.notesDeliveries(s.sid)
    expect(rows).toEqual([{ sid: s.sid, unitIndex: 4, scale, charsDrawn: expect.any(Number), at: expect.any(Number) }])
    expect(rows[0]!.charsDrawn).toBeGreaterThan(300)
    expect(s.h.stream.audit.get(s.sid)!.notesShown).toEqual({ images: 1, chars: rows[0]!.charsDrawn })
    // Notes are not unit deliveries (they never pose as a slide to the tracer).
    expect(s.store.deliveries(s.sid).filter((d) => d.unitIndex === 4)).toEqual([])
    s.stop()
  }, 60_000)

  test('unit 0 (an empty notes body) answers a valid white image; audited with chars_drawn 0; every request re-audits (C4)', async () => {
    const s = await session()
    const a = await s.c.notes(0, 1), b = await s.c.notes(0, 1)
    for (const r of [a, b]) {
      expect(r.start).toMatchObject({ type: 'notes-start', w: 720, scale: 1 })
      const d = body(r.body!)
      expect([d.unitIndex, d.w, d.h]).toEqual([0, 720, r.start!.h])
    }
    expect(s.store.notesDeliveries(s.sid).map((x) => [x.unitIndex, x.scale, x.charsDrawn])).toEqual([[0, 1, 0], [0, 1, 0]])
    expect(s.h.stream.audit.get(s.sid)!.notesShown).toEqual({ images: 2, chars: 0 })
    s.stop()
  }, 60_000)

  test('a deck with no notes parts answers the same empty image', async () => {
    const s = await session({ doc: await hiddenSlideDeck(false) })
    const r = await s.c.notes(0, 2)
    expect(r.start).toMatchObject({ type: 'notes-start', w: 1440, scale: 2 })
    expect(body(r.body!).unitIndex).toBe(0)
    expect(s.store.notesDeliveries(s.sid).map((x) => x.charsDrawn)).toEqual([0])
    s.stop()
  }, 60_000)
})

describe('notes over the wire: refusals send no frame and write no audit row', () => {
  test('an audience gets policy-refused, whatever its notes policy', async () => {
    for (const decision of [audience(), audience({ notes: true }), { viewerId: 'alice', role: 'audience', policy: {} } as AuthDecision]) {
      const s = await session({ decision })
      expect(s.m.can.notes).toBe(false)
      await expectRefused(s, 4, 1, 'policy-refused')
      s.stop()
    }
  }, 60_000)

  test('a presenter with notes: false gets policy-refused', async () => {
    const s = await session({ decision: presenter({ notes: false }) })
    await expectRefused(s, 4, 1, 'policy-refused')
    s.stop()
  }, 60_000)

  test('a presenter resumed to notes: false gets policy-refused; resumed back on, notes flow again', async () => {
    let decision = presenter()
    const store = AuditStore.open(':memory:')
    const h = await start({ doc: corpus(DECK), audit: new AuditTable({ store }), authorize: async () => decision })
    const a = new Client(h.url), m = await a.hello(); await a.ready()
    expect((await a.notes(4, 1)).start).toBeDefined()
    decision = presenter({ notes: false })
    const b = new Client(h.url), m2 = await b.hello({ resume: { sid: m.session.sid, resumeProof: m.session.token } }); await b.ready()
    expect(m2.can.notes).toBe(false)
    const before = store.notesDeliveries(m.session.sid).length
    const r = await b.notes(4, 1)
    expect(r.end).toEqual(error(r.reqId, 'policy-refused'))
    await fence(b)
    expect(framesFor(b, r.reqId)).toEqual([])
    expect(store.notesDeliveries(m.session.sid).length).toBe(before)
    decision = presenter()
    const c = new Client(h.url); await c.hello({ resume: { sid: m.session.sid, resumeProof: m2.session.token } }); await c.ready()
    expect((await c.notes(4, 1)).start).toBeDefined()
    a.close(); b.close(); c.close(); h.stop()
  }, 60_000)

  test('a docx presenter session gets policy-refused (no slides, no notes)', async () => {
    const s = await session({ doc: await buildDocx([{ runs: [{ text: 'Hello' }] }]) })
    await expectRefused(s, 0, 1, 'policy-refused')
    s.stop()
  }, 60_000)

  test('a hidden unit (\'full\') is invalid-unit; a visible one next to it is served', async () => {
    const s = await session({ doc: await hiddenSlideDeck(), decision: presenter({ contentPolicy: 'full' }) })
    expect(s.m.units.list[1]!.hidden).toBe(true)
    await expectRefused(s, 1, 1, 'invalid-unit')
    expect((await s.c.notes(2, 1)).start).toMatchObject({ type: 'notes-start' })
    s.stop()
  }, 60_000)

  test('out-of-range and non-integer units are invalid-unit; a bad scale is policy-refused', async () => {
    const s = await session()
    for (const unit of [6, -1, 1.5, '0', null, undefined, 2 ** 32]) await expectRefused(s, unit, 1, 'invalid-unit')
    for (const scale of [0, 3, 1.5, '1', null, undefined]) await expectRefused(s, 4, scale, 'policy-refused')
    s.stop()
  }, 60_000)

  test('the checks run in order: policy before the unit, the unit before the scale', async () => {
    const s = await session({ decision: audience() })
    await expectRefused(s, 99, 7, 'policy-refused')
    s.stop()
    const t = await session()
    await expectRefused(t, 99, 7, 'invalid-unit')
    t.stop()
  }, 60_000)

  test('the token: garbage is policy-refused, a stale jti session-rotated — both in the error shape', async () => {
    const h = await start({ doc: corpus(DECK), authorize: async () => presenter() })
    const a = new Client(h.url), m = await a.hello(); await a.ready()
    const b = new Client(h.url); await b.hello({ resume: { sid: m.session.sid, resumeProof: m.session.token } }); await b.ready()
    const stale = await b.notes(4, 1, { token: m.session.token })
    expect(stale.end).toEqual(error(stale.reqId, 'session-rotated'))
    const bad = await b.notes(4, 1, { token: 'garbage' })
    expect(bad.end).toEqual(error(bad.reqId, 'policy-refused'))
    expect((await b.notes(4, 1)).start).toBeDefined()
    a.close(); b.close(); h.stop()
  }, 60_000)

  test('rate: 30 notes a minute, the 31st rate-limited (no frame, no row); the socket stays open; a window later notes flow', async () => {
    let t = Date.now()
    const s = await session({ now: () => t })
    expect(NOTES_RATE_PER_MIN).toBe(30)
    for (let i = 0; i < 30; i++) { t += 200; expect((await s.c.notes(0, 1)).start).toBeDefined() }
    t += 200
    const r = await s.c.notes(0, 1)
    expect(r.end).toEqual(error(r.reqId, 'rate-limited'))
    expect(s.store.notesDeliveries(s.sid).length).toBe(30)
    expect(s.c.ws.readyState).toBe(WebSocket.OPEN)
    t += 200; await fence(s.c)
    expect(framesFor(s.c, r.reqId)).toEqual([])
    t += 61_000
    expect((await s.c.notes(0, 1)).start).toBeDefined()
    s.stop()
  }, 120_000)

  test('a notes image over the kind-7 encoded cap is refused policy-refused: no notes-start, no frame', async () => {
    const real = pool()
    const noise = new Uint8Array(9 * 1024 * 1024); for (let i = 0; i < noise.length; i += 65536) crypto.getRandomValues(noise.subarray(i, i + 65536))
    const big = new Proxy(real, {
      get(t, k) {
        if (k === 'notes') return async (sid: string, job: { draw: boolean }, est: number) => {
          const out = await (t.notes as Pool['notes'])(sid, job as never, est)
          return job.draw && out.kind === 'raster' ? { ...out, png: noise } : out
        }
        const v = Reflect.get(t, k); return typeof v === 'function' ? v.bind(t) : v
      },
    }) as Pool
    const s = await session({ pool: big, decision: presenter({ forensicMark: false }) })
    const r = await s.c.notes(4, 1)
    expect(r.end).toEqual(error(r.reqId, 'policy-refused'))
    expect(r.start).toBeUndefined()
    await fence(s.c)
    expect(framesFor(s.c, r.reqId)).toEqual([])
    s.stop()
  }, 60_000)
})

/** Notes text of `unit` in the original deck. */
async function notesText(doc: Uint8Array, unit: number) {
  const pkg = await OfficePackage.load(doc)
  return readSlideNotes(pkg, `ppt/slides/slide${unit + 1}.xml`, unit)
}

describe('notes wire audit: no notes text on the wire', () => {
  test('the kind-7 bodies, the manifest and every text message carry none of the notes text', async () => {
    const s = await session()
    const raw: string[] = []
    s.c.ws.addEventListener('message', (e) => { if (typeof e.data === 'string') raw.push(e.data) })
    const needles: string[] = []
    for (const unit of [4, 5]) for (const p of await notesText(corpus(DECK), unit)) if (p.text.trim().length >= 12) needles.push(p.text.slice(0, 12))
    // Non-vacuous: both slides carry real notes text.
    expect(needles.length).toBeGreaterThanOrEqual(6)
    const bodies: Uint8Array[] = []
    for (const unit of [4, 5]) for (const scale of [1, 2]) {
      const r = await s.c.notes(unit, scale)
      bodies.push(r.body!)
    }
    await fence(s.c)
    const texts = [...raw, ...s.c.msgs.flatMap((m) => 'text' in m ? [JSON.stringify(m.text)] : []), JSON.stringify(s.c.manifest)]
    for (const n of needles) {
      const utf8 = Buffer.from(n), json = JSON.stringify(n).slice(1, -1)
      for (const b of bodies) expect(Buffer.from(b).indexOf(utf8), `kind-7 body carries "${n}"`).toBe(-1)
      for (const t of texts) { expect(t.includes(n), `text message carries "${n}"`).toBe(false); expect(t.includes(json)).toBe(false) }
    }
    s.stop()
  }, 60_000)
})

describe('long notes', () => {
  // Short enough that all 2,000 stay under the reader's 65,536-character cap.
  const paragraphs = Array.from({ length: 2000 }, (_, i) => [`Note ${i}: say this next.`])
  const long = () => buildPptxDeck({ slides: [{ notes: paragraphs, shapes: [] }] })

  test('layout: 2,000 paragraphs truncate at 8,192 device px and end with the continuation line', async () => {
    const backend = serverBackend()
    const notes = paragraphs.map(([text]) => ({ level: 0, text: text! }))
    for (const scale of [1, 2] as const) {
      const d = drawNotes(notes, scale, backend)
      expect([d.w, d.h, d.truncated]).toEqual([720 * scale, NOTES_MAX_PX, true])
      expect(d.lines.at(-1)!.text).toBe(NOTES_CONTINUATION)
      expect(d.chars).toBeGreaterThan(0)
      expect(d.chars).toBeLessThan(notes.reduce((n, p) => n + p.text.length, 0))
    }
    // Short notes fit and are not truncated.
    const short = layoutNotes([{ level: 0, text: 'Hello' }], 1, (s) => s.length * 8)
    expect([short.truncated, short.chars, short.lines.map((l) => l.text)]).toEqual([false, 5, ['Hello']])
    // A paragraph wider than the line wraps by words; a bullet prefix is not counted as notes text.
    const wrapped = layoutNotes([{ level: 1, bullet: '-', text: 'alpha beta gamma' }], 1, (s) => s.length * 80)
    expect(wrapped.lines.map((l) => l.text)).toEqual(['- alpha', 'beta', 'gamma'])
    expect(wrapped.chars).toBe('alphabetagamma'.length)
  })

  test('over the wire: height 8,192 and the last 60 rows are not white (the continuation line), marks off', async () => {
    const s = await session({ doc: await long(), decision: presenter({ visibleMark: 'off', forensicMark: false }) })
    expect((await notesText(await long(), 0)).length).toBe(2000)
    const r = await s.c.notes(0, 1)
    expect(r.start).toEqual({ reqId: r.reqId, type: 'notes-start', w: 720, h: 8192, scale: 1 })
    const png = decodePng(body(r.body!).png)
    expect([png.w, png.h]).toEqual([720, 8192])
    const nonWhite = (from: number, to: number) => { let n = 0; for (let y = from; y < to; y++) for (let x = 0; x < png.w; x++) { const i = (y * png.w + x) * 4; if (png.rgba[i]! < 250 || png.rgba[i + 1]! < 250 || png.rgba[i + 2]! < 250) n++ } return n }
    expect(nonWhite(8192 - 60, 8192)).toBeGreaterThan(50)
    const row = s.store.notesDeliveries(s.sid)[0]!
    expect(row.charsDrawn).toBeGreaterThan(0)
    expect(row.charsDrawn).toBeLessThan(paragraphs.reduce((n, [t]) => n + t!.length, 0))
    s.stop()
  }, 60_000)
})
