/**
 * Piece 4 Task 7 — the show ops over the wire (plan D1, D4, D6, D7, D11): `show-plan`, `show-track` and
 * `show-morph`, each one charged `'unit'` send, Layer-1 per asset at send time, a `show_deliveries` row, then
 * `*-start` and the kind-6 messages.
 */
import { afterAll, describe, expect, test } from 'bun:test'
import { decodeShowMessage, SHOW_PLAN_LIMITS, isRasterLayer, type MorphPlanRecord, type ShowMessage, type ShowPlanRecord } from '../../src/stream/show-plan'
import { morphDeck } from '../../src/testdata/show-decks'
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
