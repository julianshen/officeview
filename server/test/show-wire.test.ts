/**
 * Piece 4 Task 7 — the show ops over the wire (plan D1, D4, D6, D7, D11): `show-plan`, `show-track` and
 * `show-morph`, each one charged `'unit'` send, Layer-1 per asset at send time, a `show_deliveries` row, then
 * `*-start` and the kind-6 messages.
 */
import { afterAll, describe, expect, test } from 'bun:test'
import { decodeShowMessage, SHOW_PLAN_LIMITS, isRasterLayer, type MorphPlanRecord, type ShowMessage, type ShowPlanRecord } from '../../src/stream/show-plan'
import { flyInDeck, mediaEndRevealDeck, morphDeck } from '../../src/testdata/show-decks'
import { Pool, RenderTimeout } from '../src/pool'
import { canvasBackend } from '../../src/core/canvas-backend'
import { morphFrame, type MorphFrameItem } from '../../src/slideshow/morph/render'
import { createCanvas2DCompositor } from '../../src/slideshow/render/canvas2d'
import type { LayerHandle, Lease, SlideShowWorkspace } from '../../src/slideshow/types'
import { createDiagnosticSink } from '../../src/slideshow/diagnostics'
import { DocumentHost } from '../src/render'
import { AuditTable } from '../src/audit'
import { decodePng } from '../src/mark/png'
import { AuditStore } from '../src/audit-store'
import type { AuthDecision } from '../src/hooks'
import { appearMoveDeck, hiddenLinkDeck, manyObjectsMorphDeck, morphLateDeck, morphSweepDeck, revealThenMoveDeck, tinyBoxDeck } from './show-fixtures'
import { tauFor } from '../src/detect/tau'
import { deriveKeys } from '../src/keys'
import { basis, deriveBasisKey } from '../src/mark/basis'
import { extendedCodeword } from '../src/mark/bch'
import { detect } from '../src/mark/detect'
import { messageFor } from '../src/mark/message'
import { DEFAULT_PARAMS } from '../src/mark/params'
import { Client, disposePool, start } from './harness'

afterAll(() => disposePool())

type Policy = AuthDecision['policy']
const error = (reqId: number, code: string) => ({ reqId, type: 'error', code })

async function session(doc: Uint8Array, o: Parameters<typeof start>[0] & { policy?: Policy } = {}) {
  const store = AuditStore.open(':memory:')
  let t = Date.now()
  const h = await start({ doc, audit: new AuditTable({ store }), now: () => t, policy: { contentPolicy: 'full', ...o.policy }, ...o }), c = new Client(h.url)
  const m = await c.hello(); await c.ready()
  return { h, c, m, store, sid: m.session.sid, tick: (ms = 120) => { t += ms }, stop: () => { c.close(); h.stop() } }
}
const decodeAll = (unit: number, messages: readonly Uint8Array[]): ShowMessage[] => messages.map((p) => decodeShowMessage(p, { unit, maxDecoded: SHOW_PLAN_LIMITS.assetDecoded }))
const tags = (ms: readonly ShowMessage[]): number[] => ms.map((m) => m.tag)
const planOf = (ms: readonly ShowMessage[]) => (ms[0] as { plan: ShowPlanRecord }).plan
const morphOf = (ms: readonly ShowMessage[]) => (ms[0] as { morph: MorphPlanRecord }).morph
const pngs = (ms: readonly ShowMessage[]) => new Map(ms.filter((m): m is Extract<ShowMessage, { png: Uint8Array }> => 'png' in m).map((m) => [m.assetId, Buffer.from(m.png).toString('hex')]))

describe('show wire', () => {
  test('show-plan: start then messages = 1 + 1 + layers; decodes under the client codec', async () => {
    const s = await session(await appearMoveDeck())
    expect(s.m.units.list[0]!.anim).toBe(true)
    const r = await s.c.showPlan(0, 1, 'full')
    expect(r.end).toBeUndefined()
    expect(r.start).toMatchObject({ type: 'show-plan-start', unit: 0, messages: r.messages.length })
    expect(r.start!.decodedBytes).toBeGreaterThan(0)
    const ms = decodeAll(0, r.messages)
    const plan = planOf(ms)
    const rasters = plan.layers.filter(isRasterLayer)
    expect(r.messages.length).toBe(1 + 1 + rasters.length)
    expect(tags(ms)).toEqual([1, 3, ...rasters.map(() => 2)])
    expect(ms.slice(2).map((m) => (m as { assetId: number }).assetId)).toEqual(rasters.map((l) => l.assetId))
    expect(s.store.showDeliveries(s.sid)).toEqual([expect.objectContaining({ unitIndex: 0, kind: 'plan', assetCount: 1 + rasters.length })])
    s.stop()
  }, 60_000)

  test("each show-plan, show-track and show-morph is one 'unit' charge; the 61st in 60 s refuses; the third refusal closes 4008", async () => {
    const s = await session(await morphDeck(), { policy: { forensicMark: false } })
    for (let i = 0; i < 20; i++) {
      s.tick(); expect((await s.c.showPlan(1, 1, 'full')).start).toBeDefined()
      // Unit 1 has no click step: refused after its charge.
      s.tick(); expect((await s.c.showTrack(1, 0, 1, 'full')).end?.code).toBe('invalid-unit')
      s.tick(); expect((await s.c.showMorph(0, 1, 0, 1, 'full')).start).toBeDefined()
    }
    s.tick()
    const a = await s.c.showPlan(1, 1, 'full')
    expect(a.end).toEqual(error(a.reqId, 'rate-limited'))
    s.tick()
    const b = await s.c.showMorph(0, 1, 0, 1, 'full')
    expect(b.end).toEqual(error(b.reqId, 'rate-limited'))
    s.tick()
    s.c.showTrack(1, 0, 1, 'full').catch(() => {})
    expect((await s.c.closed).code).toBe(4008)
    s.stop()
  }, 120_000)

  test('hidden unit → invalid-unit; non-anim unit → invalid-unit', async () => {
    const s = await session(await hiddenLinkDeck())
    expect(s.m.units.list.map((u) => !!u.anim)).toEqual([true, false, false])
    const hidden = await s.c.showPlan(1, 1, 'full')
    expect(hidden.end).toEqual(error(hidden.reqId, 'invalid-unit'))
    const still = await s.c.showPlan(2, 1, 'full')
    expect(still.end).toEqual(error(still.reqId, 'invalid-unit'))
    const out = await s.c.showPlan(3, 1, 'full')
    expect(out.end).toEqual(error(out.reqId, 'invalid-unit'))
    expect(s.store.showDeliveries(s.sid)).toEqual([])
    s.stop()
  }, 60_000)

  test('show-morph for a non-Morph slide → invalid-unit', async () => {
    const s = await session(await hiddenLinkDeck())
    const r = await s.c.showMorph(0, 2, 0, 1, 'full')
    expect(r.end).toEqual(error(r.reqId, 'invalid-unit'))
    const self = await s.c.showMorph(0, 0, 0, 1, 'full')
    expect(self.end).toEqual(error(self.reqId, 'invalid-unit'))
    const hidden = await s.c.showMorph(1, 2, 0, 1, 'full')
    expect(hidden.end).toEqual(error(hidden.reqId, 'invalid-unit'))
    s.stop()
  }, 60_000)

  test('fromStep out of range → invalid-unit', async () => {
    const s = await session(await morphDeck())
    for (const step of [-1, 1.5, 'x', 65]) {
      const r = await s.c.showMorph(0, 1, step, 1, 'full')
      expect(r.end).toEqual(error(r.reqId, 'invalid-unit'))
    }
    // Slide 0 has no click step: fromStep 1 is past its steps.
    const past = await s.c.showMorph(0, 1, 1, 1, 'full')
    expect(past.end).toEqual(error(past.reqId, 'invalid-unit'))
    const ok = await s.c.showMorph(0, 1, 0, 1, 'full')
    expect(ok.start).toMatchObject({ type: 'show-morph-start', from: 0, to: 1, fromStep: 0 })
    s.stop()
  }, 60_000)

  test('Morph from-side object hidden at fromStep has no asset', async () => {
    const s = await session(await morphLateDeck())
    const sideKeys = (m: MorphPlanRecord) => m.items.flatMap((it) => (it.from?.key ? [it.from.key] : []))
    const before = await s.c.showMorph(0, 1, 0, 1, 'full')
    expect(before.end).toBeUndefined()
    const ms = decodeAll(1, before.messages)
    const rec = morphOf(ms)
    expect(rec.composited).toBeUndefined()
    expect(sideKeys(rec)).toContain('sp:2')
    expect(sideKeys(rec)).not.toContain('sp:3')
    // Exactly the assets the items name, every one tagged for the `to` unit.
    const named = rec.items.flatMap((it) => [it.from?.assetId, it.to?.assetId]).filter((x): x is number => x !== undefined)
    expect(ms.slice(1).every((m) => m.tag === 2)).toBe(true)
    expect(ms.slice(1).map((m) => (m as { assetId: number }).assetId).sort()).toEqual([...named].sort())
    const after = await s.c.showMorph(0, 1, 1, 1, 'full')
    expect(sideKeys(morphOf(decodeAll(1, after.messages)))).toContain('sp:3')
    s.stop()
  }, 60_000)

  test('unproven Morph (item sweeps a masked layer) → one charge, tag 5 with composited, tag-4 frames ≤ 90, no tag-2; diagnostic morph-composited', async () => {
    const s = await session(await morphSweepDeck())
    const r = await s.c.showMorph(0, 1, 0, 1, 'full')
    expect(r.end).toBeUndefined()
    const ms = decodeAll(1, r.messages)
    const rec = morphOf(ms)
    expect(rec.composited).toEqual({ fps: 15, count: ms.length - 1 })
    expect(rec.items).toEqual([])
    expect(ms.length - 1).toBeLessThanOrEqual(90)
    expect(tags(ms)).toEqual([5, ...ms.slice(1).map(() => 4)])
    expect(r.extra).toContainEqual(expect.objectContaining({ type: 'diag', code: 'morph-composited', reason: 'swept' }))
    expect(s.store.showDeliveries(s.sid)).toEqual([expect.objectContaining({ unitIndex: 1, kind: 'morph', assetCount: ms.length - 1 })])
    s.stop()
  }, 60_000)

  test('a layered Morph rebuilt from the wire alone (records + PNGs) composes as the server does at progress 0, 0.5, 1', async () => {
    const deck = await morphDeck()
    const s = await session(deck, { policy: { visibleMark: 'off', forensicMark: false } })
    const r = await s.c.showMorph(0, 1, 0, 1, 'full')
    const ms = decodeAll(1, r.messages)
    const rec = morphOf(ms)
    expect(rec.composited).toBeUndefined()
    // The server's own composition of the same pair (its rasters, marks off).
    const host = new DocumentHost()
    const opened = await host.open('ref', deck, 'full')
    const ref = await host.showMorphReference(opened.docKey, { from: 0, to: 1, fromStep: 0, scale: 1, motion: 'full', mark: { kind: 'off' } }, [0, 0.5, 1])
    // The client side: only decoded records and PNGs.
    const lease = (bytes: number, tag: string): Lease => { const l: Lease = { tag, bytes, grow: () => true, shrink() {}, transfer: () => l, release() {} }; return l }
    const ws: SlideShowWorkspace = { limit: Number.POSITIVE_INFINITY, used: () => 0, reserveTerminal: () => undefined, lease }
    const factory = (w: number, h: number) => canvasBackend().createCanvas(Math.max(1, Math.ceil(w)), Math.max(1, Math.ceil(h))) as never
    const canvas = factory(ref.w, ref.h) as unknown as HTMLCanvasElement
    const c = createCanvas2DCompositor(canvas, factory, ws, createDiagnosticSink())
    c.setViewport({ matrix: [1, 0, 0, 1, 0, 0], deviceWidth: ref.w, deviceHeight: ref.h })
    const pngOf = new Map(ms.filter((m): m is Extract<ShowMessage, { png: Uint8Array }> => 'png' in m).map((m) => [m.assetId, m]))
    const handles = new Map<number, LayerHandle>()
    const upload = (side: NonNullable<MorphPlanRecord['items'][number]['from']>): LayerHandle => {
      const have = handles.get(side.assetId)
      if (have) return have
      const a = pngOf.get(side.assetId)!, d = decodePng(a.png)
      const cv = canvasBackend().createCanvas(d.w, d.h), ctx = cv.getContext('2d') as unknown as CanvasRenderingContext2D
      const img = ctx.createImageData(d.w, d.h); img.data.set(d.rgba); ctx.putImageData(img, 0, 0)
      const h = c.upload({ canvas: cv as never, bounds: side.bounds, scale: d.w / side.bounds.w, bytes: d.w * d.h * 4 }, lease(0, 't'))
      if ('refused' in h) throw new Error('refused')
      handles.set(side.assetId, h)
      return h
    }
    const pairs: MorphFrameItem[] = rec.items.map(({ from, to, ...rest }) => ({
      ...rest,
      ...(from ? { from: { handle: upload(from), geometry: from.geometry, ...(from.state ? { state: from.state } : {}), ...(from.part ? { part: from.part } : {}), ...(from.key ? { key: from.key } : {}) } } : {}),
      ...(to ? { to: { handle: upload(to), geometry: to.geometry, ...(to.state ? { state: to.state } : {}), ...(to.part ? { part: to.part } : {}), ...(to.key ? { key: to.key } : {}) } } : {}),
    }))
    const ctx = canvas.getContext('2d') as unknown as CanvasRenderingContext2D
    ;[0, 0.5, 1].forEach((p, k) => {
      morphFrame({ pairs }, p, c)
      const got = ctx.getImageData(0, 0, ref.w, ref.h).data, want = ref.frames[k]!
      let diff = 0
      for (let i = 0; i < want.length; i += 4) if (got[i] !== want[i] || got[i + 1] !== want[i + 1] || got[i + 2] !== want[i + 2] || got[i + 3] !== want[i + 3]) diff++
      expect({ progress: p, diff }).toEqual({ progress: p, diff: 0 })
    })
    c.dispose(); host.drop(opened.docKey)
    s.stop()
  }, 60_000)

  test('Morph past 256 objects → refused morph-fallback', async () => {
    const s = await session(await manyObjectsMorphDeck(257))
    const r = await s.c.showMorph(0, 1, 0, 1, 'full', { ms: 60_000 })
    expect(r.end).toEqual(error(r.reqId, 'morph-fallback'))
    expect(s.store.showDeliveries(s.sid)).toEqual([])
    s.stop()
  }, 120_000)

  test('every asset carries Layer-1 or an audited unmarked reason', async () => {
    const deck = await tinyBoxDeck()
    const s = await session(deck)
    const r = await s.c.showPlan(0, 1, 'full')
    const ms = decodeAll(0, r.messages)
    const ids = ms.slice(1).map((m) => (m as { assetId: number }).assetId)
    const diag = r.extra.find((x) => x.code === 'forensic-unmarked')
    expect(diag).toMatchObject({ type: 'diag', reason: 'layer-capacity', unit: 0 })
    const unmarked: number[] = diag!.assetIds
    expect(unmarked.length).toBeGreaterThan(0)
    expect(s.store.showDeliveries(s.sid)).toEqual([expect.objectContaining({ kind: 'plan', assetCount: ids.length, unmarked: unmarked.length, reason: 'layer-capacity' })])
    // Against the same bundle without Layer-1: the marked assets differ, the unmarked ones are the canonical bytes.
    const off = await session(deck, { testOnly: { forensicOff: true } })
    const plain = pngs(decodeAll(0, (await off.c.showPlan(0, 1, 'full')).messages))
    const sent = pngs(ms)
    for (const id of ids) {
      if (unmarked.includes(id)) expect(sent.get(id)).toBe(plain.get(id)!)
      else expect(sent.get(id)).not.toBe(plain.get(id)!)
    }
    expect(off.store.showDeliveries(off.sid)).toEqual([expect.objectContaining({ assetCount: ids.length, unmarked: ids.length, reason: 'forensic-off' })])
    s.stop(); off.stop()
  }, 60_000)

  test('testOnly.forensicOff cannot be set by authorize', async () => {
    const deck = await appearMoveDeck()
    const s = await session(deck, { authorize: async () => ({ viewerId: 'alice', policy: { contentPolicy: 'full', forensicOff: true } as Policy }) })
    const ms = decodeAll(0, (await s.c.showPlan(0, 1, 'full')).messages)
    const off = await session(deck, { testOnly: { forensicOff: true } })
    const plain = pngs(decodeAll(0, (await off.c.showPlan(0, 1, 'full')).messages))
    const sent = pngs(ms)
    expect(sent.size).toBeGreaterThan(0)
    for (const [id, png] of sent) expect(png).not.toBe(plain.get(id)!)
    expect(s.store.showDeliveries(s.sid)).toEqual([expect.objectContaining({ unmarked: 0 })])
    s.stop(); off.stop()
  }, 60_000)

  test('strictMarks option → moving steps composited on the wire', async () => {
    const deck = await appearMoveDeck()
    const loose = await session(deck)
    expect(planOf(decodeAll(0, (await loose.c.showPlan(0, 1, 'full')).messages)).steps[0]!.eligibility).toBe('layers')
    const strict = await session(deck, { animation: { strictMarks: true } })
    const ms = decodeAll(0, (await strict.c.showPlan(0, 1, 'full')).messages)
    expect(planOf(ms).steps[0]).toMatchObject({ eligibility: 'composited' })
    expect(tags(ms)).toEqual([1, 3])
    const track = await strict.c.showTrack(0, 0, 1, 'full')
    expect(track.start).toMatchObject({ type: 'show-track-start', unit: 0, step: 0, messages: planOf(ms).steps[0]!.track!.count })
    expect(tags(decodeAll(0, track.messages)).every((t) => t === 4)).toBe(true)
    loose.stop(); strict.stop()
  }, 60_000)

  test('a slide the classifier cannot prove → error still, no row (the client plays the piece-3 still)', async () => {
    const s = await session(await revealThenMoveDeck())
    expect(s.m.units.list[0]!.anim).toBe(true)
    const r = await s.c.showPlan(0, 1, 'full')
    expect(r.end).toEqual(error(r.reqId, 'still'))
    expect(r.messages).toEqual([])
    expect(s.store.showDeliveries(s.sid)).toEqual([])
    s.stop()
  }, 60_000)

  test('a slide holding media → error still (so no track), and no Morph pair into it (PR #78 round 12)', async () => {
    // The client ends every media node at once (no media plays in a streamed show); no server history does.
    const s = await session(await mediaEndRevealDeck('onEnd'))
    const r = await s.c.showPlan(0, 1, 'full')
    expect(r.end).toEqual(error(r.reqId, 'still'))
    expect(r.messages).toEqual([])
    s.stop()
    const m = await session(await morphDeck({ media: true }))
    const pair = await m.c.showMorph(0, 1, 0, 1, 'full', { ms: 60_000 })
    expect(pair.end).toEqual(error(pair.reqId, 'morph-fallback'))
    expect(m.store.showDeliveries(m.sid)).toEqual([])
    m.stop()
  }, 60_000)

  test("show-plan Layer-1 attributes to the session's code on the page basis `unit@scale`", async () => {
    const deck = await appearMoveDeck()
    const s = await session(deck)
    const entry = decodeAll(0, (await s.c.showPlan(0, 1, 'full')).messages)[1] as Extract<ShowMessage, { png: Uint8Array }>
    const off = await session(deck, { testOnly: { forensicOff: true } })
    const plain = decodeAll(0, (await off.c.showPlan(0, 1, 'full')).messages)[1] as Extract<ShowMessage, { png: Uint8Array }>
    const keys = await deriveKeys(new Uint8Array(32).fill(9))
    const own = s.h.stream.forensicOf(s.sid)!.code
    const codes = [own, ...Array.from({ length: 31 }, (_, i) => extendedCodeword(messageFor(keys.kMsg, 1, 2000 + i, 0)))]
    const chips = basis(deriveBasisKey(keys.kBasisRoot, s.m.docRef, '0@1'), entry.w, entry.h, DEFAULT_PARAMS)
    const bmp = (png: Uint8Array) => { const d = decodePng(png); return { width: d.w, height: d.h, data: new Uint8ClampedArray(d.rgba) } }
    expect(detect(bmp(entry.png), bmp(plain.png), chips, codes, DEFAULT_PARAMS, tauFor(codes.length))).toMatchObject({ kind: 'attributed', index: 0 })
    expect(detect(bmp(plain.png), bmp(plain.png), chips, codes, DEFAULT_PARAMS, tauFor(codes.length)).kind).not.toBe('attributed')
    s.stop(); off.stop()
  }, 60_000)

  test('show-morph under reduced motion → invalid-unit (the Morph plays as a fade, D13)', async () => {
    const s = await session(await morphDeck())
    const r = await s.c.showMorph(0, 1, 0, 1, 'reduced')
    expect(r.end).toEqual(error(r.reqId, 'invalid-unit'))
    const bad = await s.c.showMorph(0, 1, 0, 1, 'slow')
    expect(bad.end).toEqual(error(bad.reqId, 'policy-refused'))
    s.stop()
  }, 60_000)

  test('revoked session mid-bundle stops sending', async () => {
    let armed = false, sent = 0
    let revoke = () => {}
    const s = await session(await appearMoveDeck(), {
      transport: {
        sendBinary: (ws, f) => { if (armed) { sent++; if (sent === 1) revoke() } return (ws as { sendBinary(f: Uint8Array): number }).sendBinary(f) },
        bufferedAmount: (ws) => (ws as { getBufferedAmount(): number }).getBufferedAmount(),
      },
    })
    revoke = () => { s.h.stream.revoke(s.sid) }
    armed = true
    s.c.showPlan(0, 1, 'full').catch(() => {})
    expect((await s.c.closed).code).toBe(4005)
    expect(sent).toBe(1)
    s.stop()
  }, 60_000)
})

describe('show job refusal cache (controller ruling)', () => {
  const job = { scale: 1 as const, motion: 'full' as const, mark: { kind: 'fixed' as const }, strictMarks: false }
  const timedOut = async (p: Promise<unknown>) => { try { await p; return false } catch (e) { return e instanceof RenderTimeout } }

  test('a timed-out track is answered from the pool cache: no build, no second retirement; re-register forgets it', async () => {
    const pool = new Pool({ workers: 1, testHooks: true, showDeadlineMs: 400 })
    try {
      await pool.open('fly', await flyInDeck(), 'full')
      expect(await timedOut(pool.showTrack('s1', { docRef: 'fly', contentPolicy: 'full', ...job, unitIndex: 0, step: 0, hangMs: 1500 }, 1024))).toBe(true)
      expect(pool.stats().retired).toBe(1)
      const t0 = performance.now()
      expect(await timedOut(pool.showTrack('s2', { docRef: 'fly', contentPolicy: 'full', ...job, unitIndex: 0, step: 0 }, 1024))).toBe(true)
      expect(performance.now() - t0).toBeLessThan(100)
      expect(pool.stats()).toMatchObject({ retired: 1, reserved: 0, running: 0 })
      // Another step is its own entry: it runs on a worker (and, hung, times out and retires it).
      expect(await timedOut(pool.showTrack('s2', { docRef: 'fly', contentPolicy: 'full', ...job, unitIndex: 0, step: 1, hangMs: 1500 }, 1024))).toBe(true)
      expect(pool.stats().retired).toBe(2)
      await pool.open('fly', await flyInDeck(), 'full')
      expect(await timedOut(pool.showTrack('s1', { docRef: 'fly', contentPolicy: 'full', ...job, unitIndex: 0, step: 0, hangMs: 1500 }, 1024))).toBe(true)
      expect(pool.stats().retired).toBe(3)
    } finally { pool.dispose() }
  }, 60_000)

  test('a timed-out Morph pair is answered from the pool cache; morph-fallback is cached too; forget clears both', async () => {
    const pool = new Pool({ workers: 1, testHooks: true, showDeadlineMs: 400 })
    try {
      await pool.open('m', await morphDeck(), 'full')
      const pair = { docRef: 'm', contentPolicy: 'full' as const, ...job, from: 0, to: 1, fromStep: 0 }
      expect(await timedOut(pool.showMorph('s1', { ...pair, hangMs: 1500 }, 1024))).toBe(true)
      expect(pool.stats().retired).toBe(1)
      const t0 = performance.now()
      expect(await timedOut(pool.showMorph('s2', { ...pair, mark: { kind: 'off' } }, 1024))).toBe(true)
      expect(performance.now() - t0).toBeLessThan(100)
      expect(pool.stats()).toMatchObject({ retired: 1, reserved: 0, running: 0 })
      // Forgotten, then registered again: the pair runs on a worker again.
      pool.forget('m')
      await pool.open('m', await morphDeck(), 'full')
      expect(await timedOut(pool.showMorph('s1', { ...pair, hangMs: 1500 }, 1024))).toBe(true)
      expect(pool.stats().retired).toBe(2)
    } finally { pool.dispose() }
    const roomy = new Pool({ workers: 1 })
    try {
      await roomy.open('many', await manyObjectsMorphDeck(257), 'full')
      const many = { docRef: 'many', contentPolicy: 'full' as const, ...job, from: 0, to: 1, fromStep: 0 }
      expect(await roomy.showMorph('s1', many, 1024)).toEqual({ refused: 'morph-fallback' })
      const pages = roomy.stats().workers[0]!.pages
      const t1 = performance.now()
      expect(await roomy.showMorph('s1', many, 1024)).toEqual({ refused: 'morph-fallback' })
      expect(performance.now() - t1).toBeLessThan(100)
      expect(roomy.stats().workers[0]!.pages).toBe(pages)
    } finally { roomy.dispose() }
  }, 120_000)
})
