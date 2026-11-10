/** Stage 2b Task 5: trace() — one §13.2 decision over every candidate. */
import { afterAll, beforeAll, describe, expect, it } from 'bun:test'
import { Database } from 'bun:sqlite'
import { createCanvas } from '@napi-rs/canvas'
import { mkdtempSync, rmSync } from 'node:fs'
import { tmpdir } from 'node:os'
import { join } from 'node:path'
import JSZip from 'jszip'
import { decodeRasterBody } from '../../src/stream/raster'
import { average, crop, rotate, scale } from '../../scripts/stream-s0/mark/attacks'
import { AuditTable } from '../src/audit'
import { AuditStore } from '../src/audit-store'
import { decodeImage } from '../src/detect/image'
import { MAX_CANDIDATES, TAU_TABLE, tauFor } from '../src/detect/tau'
import { trace, type TraceAnswer, type TraceInput, type TraceObservation } from '../src/detect/trace'
import { deriveKeys, type Keys } from '../src/keys'
import { decodePng } from '../src/mark/png'
import type { Bitmap } from '../src/mark/types'
import { DocumentHost } from '../src/render'
import { Client, corpus, disposePool, start, type Harness } from './harness'

const dirs: string[] = []
afterAll(() => { disposePool(); for (const d of dirs.splice(0)) rmSync(d, { recursive: true, force: true }) })
const dbPath = () => { const d = mkdtempSync(join(tmpdir(), 'ov-trace-')); dirs.push(d); return join(d, 'audit.db') }

const doc = corpus('pypptx-test.pptx')
const host = new DocumentHost() // shared: per-trace times below exclude the one-off document parse
const keys1 = await deriveKeys(new Uint8Array(32).fill(9), 1)
const keys2 = await deriveKeys(new Uint8Array(32).fill(2), 2)
const ring1: ReadonlyMap<number, Keys> = new Map([[1, keys1]])
async function timed(_name: string, input: Omit<TraceInput, 'host'>): Promise<TraceAnswer> {
  return trace({ ...input, host })
}

/** Opens a session (viewer `who`) and fetches the given (unitIndex, scale) units; returns sid and received PNGs. */
async function session(h: Harness, setViewer: (v: string) => void, who: string, units: Array<[number, number]> = [[0, 1]]) {
  setViewer(who)
  const c = new Client(h.url), m = await c.hello(); await c.ready()
  const pngs: Uint8Array[] = []
  let reqId = 10
  for (const [u, s] of units) {
    c.unit(reqId, u, s)
    const { unit } = await c.unitMessages(reqId++)
    if (unit.discriminator !== 2) throw new Error('expected raster')
    pngs.push(decodeRasterBody(unit.body).png)
  }
  return { sid: m.session.sid, pngs, client: c }
}
function jpegBytes(b: Bitmap, q: number): Uint8Array {
  const c = createCanvas(b.width, b.height), ctx = c.getContext('2d'), img = ctx.createImageData(b.width, b.height)
  img.data.set(b.data); ctx.putImageData(img, 0, 0)
  return new Uint8Array(c.encodeSync('jpeg', q))
}

describe('tau', () => {
  it('declared table and lookup', () => {
    expect(TAU_TABLE.map((t) => t.m)).toEqual([2 ** 16, 2 ** 20, 2 ** 24])
    expect(TAU_TABLE[0]).toMatchObject({ abs: 7.99, margin: 3.85, source: 'stage-0 provisional' })
    expect(TAU_TABLE[1]!.abs).toBeCloseTo(Math.sqrt(7.99 ** 2 + 2 * Math.log(16)), 12)
    expect(TAU_TABLE[2]).toMatchObject({ abs: 8.66, margin: 3.85 })
    expect(tauFor(1).m).toBe(2 ** 16); expect(tauFor(2 ** 16).m).toBe(2 ** 16)
    expect(tauFor(2 ** 16 + 1).m).toBe(2 ** 20); expect(tauFor(MAX_CANDIDATES).m).toBe(2 ** 24)
    expect(MAX_CANDIDATES).toBe(2 ** 22)
    expect(() => tauFor(2 ** 24 + 1)).toThrow()
  })
})

describe('trace', () => {
  const path = dbPath()
  let store: AuditStore, h: Harness, reader: AuditStore
  let A: Awaited<ReturnType<typeof session>>, B: Awaited<ReturnType<typeof session>>, C: Awaited<ReturnType<typeof session>>, A2: Awaited<ReturnType<typeof session>>
  beforeAll(async () => {
    let current = 'alice'
    const set = (v: string) => { current = v }
    store = AuditStore.open(path)
    h = await start({ keys: keys1, audit: new AuditTable({ store }), authorize: async () => ({ viewerId: current, policy: { visibleMark: 'viewer' } }) })
    A = await session(h, set, 'alice', [[0, 1]])
    B = await session(h, set, 'bob', [[0, 1], [0, 2]])
    C = await session(h, set, 'carol', [[0, 1]])
    A2 = await session(h, set, 'alice', [[0, 1]])
    reader = AuditStore.openReadOnlyForTrace(path)
  }, 60_000)
  afterAll(() => { for (const s of [A, B, C, A2]) s?.client.close(); h?.stop(); reader?.close(); store?.close() })
  const input = async (capture: Bitmap, unitIndex = 0) => ({ capture, doc, unitIndex, keys: ring1, store: reader })

  it('decodeImage matches the PNG decoder', async () => {
    const d = decodePng(A.pngs[0]!), b = await decodeImage(A.pngs[0]!)
    expect([b.width, b.height]).toEqual([d.w, d.h])
    expect(b.data).toEqual(d.rgba)
  })

  it('traces the leaking session', async () => {
    const a = await timed('leaking session', await input(await decodeImage(A.pngs[0]!)))
    expect(a).toMatchObject({ kind: 'traced', sid: A.sid, viewerId: 'alice', keyGeneration: 1, unitIndex: 0, scale: 1, ip: expect.any(String), startedAt: expect.any(Number) })
    const b = await timed('leaking session B', await input(await decodeImage(B.pngs[0]!)))
    expect(b).toMatchObject({ kind: 'traced', sid: B.sid, viewerId: 'bob' })
  }, 60_000)

  it('same viewer twice: both candidates, right one named', async () => {
    const a = await timed('same viewer A2', await input(await decodeImage(A2.pngs[0]!)))
    expect(a).toMatchObject({ kind: 'traced', sid: A2.sid, viewerId: 'alice' })
  }, 60_000)

  it('bucket-2 capture at half zoom traces', async () => {
    const a = await timed('bucket-2 at 0.5', await input(scale(await decodeImage(B.pngs[1]!), 0.5)))
    expect(a).toMatchObject({ kind: 'traced', sid: B.sid, scale: 2 })
    // The scale-1 registration is out of scope for this capture: its 4 candidates stay in the
    // decision at the null expectation 0, so the margin is z - max(other scores, 0) = z, never Infinity.
    if (a.kind !== 'traced') throw new Error('not traced')
    expect(Number.isFinite(a.margin)).toBe(true)
    expect(a.margin).toBe(a.z)
  }, 60_000)

  it('zoom 1.25 + jpeg q70 traces', async () => {
    const capture = await decodeImage(jpegBytes(scale(await decodeImage(C.pngs[0]!), 1.25), 70))
    const a = await timed('zoom 1.25 + q70', await input(capture))
    expect(a).toMatchObject({ kind: 'traced', sid: C.sid, scale: 1 })
  }, 60_000)

  it('average of two is never a single innocent', async () => {
    const a = await timed('average A+B', await input(average([await decodeImage(A.pngs[0]!), await decodeImage(B.pngs[0]!)])))
    expect(['collusion', 'no-match', 'traced']).toContain(a.kind)
    if (a.kind === 'traced') expect([A.sid, B.sid]).toContain(a.sid)
    if (a.kind === 'collusion') for (const c of a.candidates) expect([A.sid, B.sid]).toContain(c.sid)
  }, 60_000)

  it('unmarked canonical is not traced', async () => {
    const canonical = h.stream.canonicalOf(A.sid, 0, 1)!
    const a = await timed('unmarked canonical', await input(await decodeImage(canonical)))
    expect(['no-match', 'no-signal']).toContain(a.kind)
  }, 60_000)

  it('rotated capture', async () => {
    const a = await timed('rotate 5', await input(rotate(await decodeImage(A.pngs[0]!), 5)))
    expect(a).toEqual({ kind: 'no-signal', reason: 'geometry' })
  }, 60_000)

  it('observe reports the decision inputs (diagnostic, no behaviour change)', async () => {
    const seen: TraceObservation[] = []
    const a = await trace({ ...(await input(await decodeImage(A.pngs[0]!))), host, observe: (o) => seen.push(o) })
    expect(a).toMatchObject({ kind: 'traced', sid: A.sid })
    expect(seen).toHaveLength(1)
    const o = seen[0]!
    // A, B (two scales), C, A2: five l1 deliveries in one list, τ for 2^16.
    expect(o.scores.length).toBe(5); expect(o.sids).toHaveLength(5); expect(o.tau.m).toBe(2 ** 16)
    expect(o.sids[o.scores.indexOf(Math.max(...o.scores))]).toBe(A.sid)
    expect(o.poses.map((p) => p.scale).sort()).toEqual([1, 2])
    for (const p of o.poses) expect(typeof p.residual).toBe('number')
    expect(o.poses.find((p) => p.scale === 1)!.inScope).toBe(true)
    expect(o.unscored).toBe(o.poses.find((p) => p.scale === 2)!.inScope ? 0 : 1)
    if (a.kind !== 'traced') throw new Error('not traced')
    expect(o.scores[o.sids.indexOf(A.sid)]).toBe(a.z)
    const r: TraceObservation[] = []
    expect(await trace({ ...(await input(rotate(await decodeImage(A.pngs[0]!), 5))), host, observe: (x) => r.push(x) })).toEqual({ kind: 'no-signal', reason: 'geometry' })
    expect(r[0]!.geometry).toBe('unsupported'); expect(r[0]!.poses.every((p) => !p.inScope)).toBe(true); expect(r[0]!.unscored).toBe(5)
  }, 60_000)

  it('unrequested page', async () => {
    const a = await timed('unrequested page', await input(await decodeImage(A.pngs[0]!), 1))
    expect(a).toEqual({ kind: 'no-signal', reason: 'no-candidates' })
  })
})

describe('trace fail-closed and durability', () => {
  it('never-marked page', async () => {
    const zip = await JSZip.loadAsync(doc)
    const pres = await zip.file('ppt/presentation.xml')!.async('string')
    zip.file('ppt/presentation.xml', pres.replace(/<p:sldSz[^>]*\/>/, '<p:sldSz cx="952500" cy="476250"/>'))
    const tiny = new Uint8Array(await zip.generateAsync({ type: 'uint8array' }))
    const store = AuditStore.open(':memory:')
    const h = await start({ keys: keys1, doc: tiny, audit: new AuditTable({ store }) })
    const s = await session(h, () => {}, 'alice')
    const a = await timed('never-marked', { capture: await decodeImage(s.pngs[0]!), doc: tiny, unitIndex: 0, keys: ring1, store })
    expect(a).toEqual({ kind: 'no-signal', reason: 'never-marked' })
    s.client.close(); h.stop(); store.close()
  }, 60_000)

  it('reference mismatch fails closed', async () => {
    const path = dbPath(), store = AuditStore.open(path)
    const h = await start({ keys: keys1, audit: new AuditTable({ store }) })
    const s = await session(h, () => {}, 'alice'), t = await session(h, () => {}, 'alice')
    const db = new Database(path)
    db.query('UPDATE deliveries SET ref_sha256 = ? WHERE sid = ?').run('00'.repeat(32), t.sid)
    db.close()
    const a = await timed('reference mismatch', { capture: await decodeImage(s.pngs[0]!), doc, unitIndex: 0, keys: ring1, store })
    expect(a).toEqual({ kind: 'no-signal', reason: 'reference-mismatch' })
    s.client.close(); t.client.close(); h.stop(); store.close()
  }, 60_000)

  it('traces across key generations', async () => {
    const path = dbPath()
    let store = AuditStore.open(path)
    let h = await start({ keys: keys1, audit: new AuditTable({ store }) })
    const s1 = await session(h, () => {}, 'alice')
    s1.client.close(); h.stop(); store.close()
    store = AuditStore.open(path)
    h = await start({ keys: keys2, audit: new AuditTable({ store }) })
    const s2 = await session(h, () => {}, 'alice')
    s2.client.close(); h.stop(); store.close()
    const reader = AuditStore.openReadOnlyForTrace(path), keys = new Map([[1, keys1], [2, keys2]])
    const a1 = await timed('gen 1 capture', { capture: await decodeImage(s1.pngs[0]!), doc, unitIndex: 0, keys, store: reader })
    expect(a1).toMatchObject({ kind: 'traced', sid: s1.sid, keyGeneration: 1 })
    const a2 = await timed('gen 2 capture', { capture: await decodeImage(s2.pngs[0]!), doc, unitIndex: 0, keys, store: reader })
    expect(a2).toMatchObject({ kind: 'traced', sid: s2.sid, keyGeneration: 2 })
    reader.close()
  }, 90_000)

  it('a gen-2 capture registers with gen-2 chips (blank page, cropped)', async () => {
    // A blank slide with no visible mark: registration can only synchronise on the carriers,
    // so anchoring a gen-2 capture on gen-1 chips misplaces it.
    const zip = await JSZip.loadAsync(doc)
    const slide = await zip.file('ppt/slides/slide1.xml')!.async('string')
    zip.file('ppt/slides/slide1.xml', slide.replace(/<p:sp>[\s\S]*<\/p:sp>/, ''))
    const blank = new Uint8Array(await zip.generateAsync({ type: 'uint8array' }))
    const path = dbPath()
    let store = AuditStore.open(path)
    let h = await start({ keys: keys1, doc: blank, policy: { visibleMark: 'off' }, audit: new AuditTable({ store }) })
    const s1 = await session(h, () => {}, 'alice')
    s1.client.close(); h.stop(); store.close()
    store = AuditStore.open(path)
    h = await start({ keys: keys2, doc: blank, policy: { visibleMark: 'off' }, audit: new AuditTable({ store }) })
    const s2 = await session(h, () => {}, 'alice')
    s2.client.close(); h.stop(); store.close()
    const reader = AuditStore.openReadOnlyForTrace(path), keys = new Map([[1, keys1], [2, keys2]])
    const a = await timed('blank gen 2 cropped', { capture: crop(await decodeImage(s2.pngs[0]!), 0.1, 'bottom-right'), doc: blank, unitIndex: 0, keys, store: reader })
    expect(a).toMatchObject({ kind: 'traced', sid: s2.sid, keyGeneration: 2 })
    reader.close()
  }, 90_000)

  it('traces a session served before a restart', async () => {
    const path = dbPath(), store = AuditStore.open(path)
    const h = await start({ keys: keys1, audit: new AuditTable({ store }) })
    const s = await session(h, () => {}, 'alice'), t = await session(h, () => {}, 'alice')
    s.client.close(); t.client.close(); h.stop(); store.close()
    const reader = AuditStore.openReadOnlyForTrace(path)
    const a = await timed('after restart', { capture: await decodeImage(t.pngs[0]!), doc, unitIndex: 0, keys: ring1, store: reader })
    expect(a).toMatchObject({ kind: 'traced', sid: t.sid, keyGeneration: 1 })
    reader.close()
  }, 60_000)
})
