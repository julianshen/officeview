/**
 * Task 11b (owner decision 2026-11-19): a slide that falls back to its still plays its sounds. With media on, a
 * classifier still (`refused: 'still'`) that has audio uses carries a sound-only plan (`sounds`): its records, bound
 * per session as any plan's; its transition sound; and the no-input audio schedule (the shared scheduler and
 * `slideMediaEnds` on the initial cursor). Click-, trigger- and event-started sounds ship refused `media-unavailable`.
 * Nothing drawn changes: the refusal, the classification and every asset are the same with media on and off.
 */
import { afterAll, describe, expect, it } from 'bun:test'
import { mkdtempSync, rmSync } from 'node:fs'
import { tmpdir } from 'node:os'
import { join } from 'node:path'
import { randomBytes } from 'node:crypto'
import { decodeShowMessage, encodeShowMessage, MEDIA_ID, SHOW_PLAN_LIMITS, type ShowPlanRecord } from '../../src/stream/show-plan'
import type { BehaviourNode, TimeNode } from '../../src/slideshow/types'
import { flyInDeck, mediaDeck, mixedMediaDeck, narrationDeck, repeatDeck, stillSoundDeck, videoDeck } from '../../src/testdata/show-decks'
import { DocumentHost } from '../src/render'
import { Pool } from '../src/pool'
import { createMediaTemp } from '../src/media/temp'
import { createMediaServices } from '../src/media/services'
import { bindMedia, mediaIdentity, sessionMediaId } from '../src/media/bind'
import type { BuiltShowPlan, ShowPlanJob, ShowRefusal } from '../src/show/build'
import { Client, disposePool, start } from './harness'
import { ffmpegForTests } from './media/ffmpeg-gate'

const info = await ffmpegForTests()
const root = mkdtempSync(join(tmpdir(), 'ov-t11b-'))
const temp = await createMediaTemp(root)
let mediaPool: Pool | undefined
const poolOn = () => (mediaPool ??= new Pool({ workers: 1, media: { info: info!, dir: temp.dir } }))
afterAll(async () => { mediaPool?.dispose(); disposePool(); await temp.remove(); rmSync(root, { recursive: true, force: true }) })

const inkBox = { id: 2, name: 'Ink', prst: 'rect', off: ['3619500', '2476500'] as [string, string], ext: ['1905000', '1905000'] as [string, string], fill: '13F00D' }
const job = (over: Partial<ShowPlanJob> = {}): ShowPlanJob => ({ unitIndex: 0, scale: 1, motion: 'full', mark: { kind: 'fixed' }, strictMarks: false, media: true, ...over })
async function opened(bytes: Uint8Array) {
  const h = new DocumentHost()
  h.configureMedia({ info: info!, dir: temp.dir })
  const { docKey } = await h.open(`d${Math.random()}`, bytes, 'visible-only')
  return { h, docKey }
}
async function plan(bytes: Uint8Array, o: Partial<ShowPlanJob> = {}) {
  const { h, docKey } = await opened(bytes)
  return await h.showPlan(docKey, job(o))
}
/** The still refusal and its sound-only plan. */
function sounds(r: BuiltShowPlan | ShowRefusal) {
  if (!('refused' in r)) throw new Error('expected a still')
  expect(r.refused).toBe('still')
  if (!r.sounds) throw new Error('expected a sound-only plan')
  return r.sounds
}
const bind = (s: { plan: ShowPlanRecord; mediaRefs: BuiltShowPlan['mediaRefs'] }, k = new Uint8Array(randomBytes(32))) => bindMedia(s.plan, (ref) => sessionMediaId(k, mediaIdentity(ref)), s.mediaRefs)
const viaClient = (p: ShowPlanRecord) => (decodeShowMessage(encodeShowMessage({ tag: 1, plan: p }), { unit: p.slide.index, maxDecoded: SHOW_PLAN_LIMITS.showDecoded }) as { plan: ShowPlanRecord }).plan
function leaves(p: ShowPlanRecord): Array<{ node: BehaviourNode; begin: number }> {
  const out: Array<{ node: BehaviourNode; begin: number }> = []
  const delay = (n: TimeNode) => { const d = n.ctn.start[0]?.delayMs; return typeof d === 'number' ? d : NaN }
  const walk = (n: TimeNode, at: number, top: boolean) => {
    const here = top ? at : at + delay(n)
    if ('children' in n) n.children.forEach(c => walk(c, here, false)); else out.push({ node: n, begin: here })
  }
  if (p.show.timing && !('refused' in p.show.timing)) walk(p.show.timing.root, 0, true)
  return out
}
const drawnless = (p: ShowPlanRecord) => {
  expect(p.fallback).toBe('still')
  expect({ layers: p.layers, steps: p.steps, base: p.base, triggers: p.triggers, linkRegions: p.linkRegions, iterate: p.iterate, builds: p.show.builds }).toEqual({ layers: [], steps: [], base: [], triggers: [], linkRegions: [], iterate: { keys: [], groups: [] }, builds: [] })
}

describe.skipIf(!info)('still slides play their sounds (Task 11b)', () => {
  it('a video slide with a narration that starts on entry plays the narration; the video stays a still', async () => {
    const s = sounds(await plan(await mixedMediaDeck('video')))
    drawnless(s.plan)
    expect(s.plan.media).toEqual([{ mediaId: '@a1.4', kind: 'audio', spid: '4', durationMs: expect.any(Number) }])
    expect(s.plan.media![0]!.durationMs).toBeGreaterThan(0)
    // The schedule: the narration's leaf at its begin; no video node, no build.
    expect(leaves(s.plan).map(l => [l.node.kind, l.node.ctn.id, l.begin])).toEqual([['audio', 4, 0]])
    const b = viaClient(bind(s))
    expect(b.media![0]!.mediaId).toMatch(MEDIA_ID)
    expect(JSON.stringify(b)).not.toContain('@a')
    // Bound per session exactly as a plan is: the same identity gives the same id under one key.
    const k = new Uint8Array(randomBytes(32))
    expect(bind(s, k).media![0]!.mediaId).toBe(sessionMediaId(k, mediaIdentity(s.mediaRefs[0]!.ref)))
  })

  it('a slide stilled by an audio-end dependency plays its auto-start sound and refuses its click-started one', async () => {
    const s = sounds(await plan(await stillSoundDeck('audio-end')))
    drawnless(s.plan)
    const byId = new Map(s.plan.media!.map(r => [r.mediaId, r]))
    expect(byId.get('@a0.3')).toEqual({ mediaId: '@a0.3', kind: 'audio', spid: '3', durationMs: expect.any(Number) })
    // The effect sound begins only with click 1: on a still a click advances, so it never plays.
    expect(byId.get('@a1')).toEqual({ mediaId: '@a1', kind: 'audio', durationMs: 0, refused: 'media-unavailable' })
    // The narration at its computed begin (250 ms); unbounded, as the client's own clock ends it where the server's
    // does (begin + durationMs), and records that end in its history (the build proves it before shipping).
    const l = leaves(s.plan)
    expect(l.map(x => [x.node.kind, x.node.ctn.id, x.begin])).toEqual([['audio', 3, 250]])
    const root = s.plan.show.timing && !('refused' in s.plan.show.timing) ? s.plan.show.timing.root : undefined
    const wrap = root && 'children' in root ? root.children[0]! : undefined
    expect(wrap?.ctn.durMs).toBe('indefinite')
    expect(wrap?.ctn.start).toEqual([{ delayMs: 250 }])
    expect(viaClient(bind(s)).media).toHaveLength(2)
  })

  it('a transition sound plays on entry to a still', async () => {
    const s = sounds(await plan(await stillSoundDeck('transition')))
    drawnless(s.plan)
    // k 1: the video's part sorts first (D21); it ships no record (D22).
    expect(s.plan.media).toEqual([{ mediaId: '@a1', kind: 'audio', durationMs: expect.any(Number) }])
    expect(s.plan.show.transition?.sound).toEqual({ start: { mediaId: '@a1', loop: false } })
    expect(leaves(s.plan)).toEqual([])
    const b = viaClient(bind(s))
    expect(b.show.transition!.sound!.start!.mediaId).toBe(b.media![0]!.mediaId)
  })

  it('fix round 1 (Important 2): a shape command the schedule would retarget onto another audio leaf leaves the still silent', async () => {
    // Shape 3 holds two audio leaves: cTn 3 (first: the one its commands drive) begins twice and so is not scheduled,
    // cTn 4 begins once. A playFrom on shape 3 at 300 ms drives cTn 3 locally; in a schedule without cTn 3 it would
    // seek cTn 4. The ink box is revealed on cTn 4's end (the still).
    const tgt3 = '<p:tgtEl><p:spTgt spid="3"/></p:tgtEl>'
    const twice = `<p:audio><p:cMediaNode vol="80000"><p:cTn id="3" dur="100" restart="always" fill="hold"><p:stCondLst><p:cond delay="0"/><p:cond delay="500"/></p:stCondLst></p:cTn>${tgt3}</p:cMediaNode></p:audio>`
    const once = `<p:audio><p:cMediaNode vol="80000"><p:cTn id="4" fill="hold"><p:stCondLst><p:cond delay="0"/></p:stCondLst></p:cTn>${tgt3}</p:cMediaNode></p:audio>`
    const seek = `<p:cmd type="call" cmd="playFrom(0.0)"><p:cBhvr><p:cTn id="5" dur="1" fill="hold"><p:stCondLst><p:cond delay="300"/></p:stCondLst></p:cTn>${tgt3}</p:cBhvr></p:cmd>`
    const reveal = '<p:par><p:cTn id="6" fill="hold"><p:stCondLst><p:cond evt="onEnd" delay="0"><p:tn val="4"/></p:cond></p:stCondLst><p:childTnLst><p:set><p:cBhvr><p:cTn id="7" dur="1" fill="hold"><p:stCondLst><p:cond delay="0"/></p:stCondLst></p:cTn><p:tgtEl><p:spTgt spid="2"/></p:tgtEl><p:attrNameLst><p:attrName>style.visibility</p:attrName></p:attrNameLst></p:cBhvr><p:to><p:strVal val="visible"/></p:to></p:set></p:childTnLst></p:cTn></p:par>'
    const deck = (cmds: string) => mediaDeck({ slides: [{ boxes: [{ ...inkBox }], media: [{ spid: 3, file: 'a.m4a' }], timing: twice + once + cmds + reveal }] })
    const r = await plan(await deck(seek))
    expect('refused' in r && r.refused).toBe('still')
    expect(r).not.toHaveProperty('sounds')
    // Control: without the command the same still sounds (cTn 4 scheduled, cTn 3 not): the command is what silences it.
    expect(leaves(sounds(await plan(await deck(''))).plan).map(l => l.node.ctn.id)).toEqual([4])
  })

  it('fix round 1 (minor 7): a sound-only plan carries the transition sound only, nothing visual', async () => {
    const s = sounds(await plan(await stillSoundDeck('transition')))
    expect(s.plan.show.transition as unknown).toEqual({ sound: { start: { mediaId: '@a1', loop: false } } })
    // No sound, no transition at all.
    expect(sounds(await plan(await stillSoundDeck('audio-end'))).plan.show.transition).toBeUndefined()
  })

  it('media off, a still without audio, and a non-still carry no sound-only plan', async () => {
    expect(await plan(await stillSoundDeck('audio-end'), { media: false })).toEqual({ refused: 'still', diagnostics: expect.any(Array) })
    // A video-only still (D22: no record ships).
    const v = await plan(await videoDeck())
    expect('refused' in v && v.refused).toBe('still')
    expect(v).not.toHaveProperty('sounds')
    expect(await plan(await narrationDeck())).not.toHaveProperty('refused')
  })

  it('equivalence: classification and every asset are byte-identical with media on and off', async () => {
    for (const make of [() => stillSoundDeck('audio-end'), () => stillSoundDeck('transition'), () => mixedMediaDeck('video'), () => repeatDeck(2), flyInDeck, narrationDeck]) {
      const bytes = await make()
      // One host per side: media on never touches the media-off host's caches.
      const A = await opened(bytes), B = await opened(bytes)
      const on = await A.h.showPlan(A.docKey, job({ media: true })), off = await B.h.showPlan(B.docKey, job({ media: false }))
      expect('refused' in on).toBe('refused' in off)
      if ('refused' in on && 'refused' in off) {
        const { sounds: _s, ...still } = on
        expect(still).toEqual(off)
      } else if (!('refused' in on) && !('refused' in off)) {
        expect(on.assets).toEqual(off.assets)
        const strip = (p: ShowPlanRecord) => ({ ...p, media: undefined, show: { ...p.show, timing: undefined, transition: undefined } })
        expect(strip(on.plan)).toEqual(strip(off.plan))
      }
      // The still asset (the piece-3 unit raster the client draws) is the same after either build.
      const unit = (x: typeof A) => x.h.render(x.docKey, { unitIndex: 0, scale: 1, projection: 'final-raster', mark: { kind: 'fixed' } })
      expect(unit(A)).toEqual(unit(B))
    }
  }, 60_000)
})

describe.skipIf(!info)('the still sound-only plan on the wire (Task 11b)', () => {
  it('show-plan answers a still with audio with show-plan-start and one plan message; media off and a no-audio still are unchanged', async () => {
    const doc = await stillSoundDeck('audio-end')
    // Real media services (the queue, the cache, the encoder), as main.ts builds them.
    const h = await start({ doc, media: {} }), c = new Client(h.url)
    try {
      const m = await c.hello(); await c.ready()
      // D19: the still's unit is a media unit (the gate covers stills with audio uses).
      expect(m.units.list[0]!.media).toBe(true)
      const r = await c.showPlan(0, 1, 'full')
      expect(r.end).toBeUndefined()
      expect(r.messages).toHaveLength(1)
      expect(r.start!.messages).toBe(1)
      const p = (decodeShowMessage(r.messages[0]!, { unit: 0, maxDecoded: SHOW_PLAN_LIMITS.showDecoded }) as { plan: ShowPlanRecord }).plan
      expect(p.fallback).toBe('still')
      // Declared as the message's whole decoded payload (its tag included): a client's per-message decoded cap is
      // the declared total, and a sound-only plan has no asset to make room (the browser suite found the off-by-one).
      expect(r.start!.decodedBytes).toBe(r.messages[0]!.length)
      const playable = p.media!.find(x => !x.refused)!
      const clicked = p.media!.find(x => x.refused)!
      expect(clicked.refused).toBe('media-unavailable')
      // Queue and delivery are slide-kind agnostic (Tasks 8, 9): the still's narration is served like any plan's.
      expect((await c.mediaShow(0)).type).toBe('media-show-ok')
      const got = await c.mediaRequest(playable.mediaId, { unit: 0, ms: 60_000 })
      expect(got.start?.type).toBe('media-start')
      expect(got.chunks.length).toBe(got.start!.chunks)
    } finally { c.close(); h.stop() }
    // Media off: the piece-4 refusal, byte for byte. A still without audio, media on: the same.
    for (const [d, media] of [[doc, { enabled: false, info }], [await videoDeck(), { enabled: true, info }]] as const) {
      const h2 = await start({ doc: d, pool: poolOn(), media }), c2 = new Client(h2.url)
      try {
        await c2.hello(); await c2.ready()
        const r = await c2.showPlan(0, 1, 'full')
        expect(r.end).toEqual({ reqId: r.reqId, type: 'error', code: 'still' })
        expect(r.start).toBeUndefined()
        expect(r.messages).toHaveLength(0)
      } finally { c2.close(); h2.stop() }
    }
  }, 120_000)
})

/**
 * Task 11b fix round (controller ruling): nothing delivers a use the plan refuses, sound-only still plans included.
 * Each plan bound for a session records its (mediaId, unit) answers; `media-request` for a refused pair answers the
 * plan's reason, and one for a unit no plan was sent for answers `media-unavailable`.
 */
describe.skipIf(!info)('a plan-refused use is never delivered (Task 11b fix round)', () => {
  const planOf = (r: Awaited<ReturnType<Client['showPlan']>>, unit: number) => (decodeShowMessage(r.messages[0]!, { unit, maxDecoded: SHOW_PLAN_LIMITS.showDecoded }) as { plan: ShowPlanRecord }).plan
  const reason = (r: Awaited<ReturnType<Client['mediaRequest']>>) => (r.start ? 'delivered' : r.end?.type === 'media-refused' ? r.end.reason : JSON.stringify(r.end))

  it("a still's click sound requested by id is refused; the animated slide's playable use of the same part is served", async () => {
    const h = await start({ doc: await stillSoundDeck('shared'), media: {} }), c = new Client(h.url)
    try {
      await c.hello(); await c.ready()
      const still = planOf(await c.showPlan(0, 1, 'full'), 0)
      expect(still.fallback).toBe('still')
      const clicked = still.media!.find(x => x.refused)!
      expect(clicked.refused).toBe('media-unavailable')
      const live = planOf(await c.showPlan(1, 1, 'full'), 1)
      expect(live.fallback).toBeUndefined()
      // One part, identity audio|k: the same mediaId on both slides, refused on the still only.
      expect(live.media).toEqual([{ mediaId: clicked.mediaId, kind: 'audio', durationMs: expect.any(Number) }])
      expect((await c.mediaShow(0)).type).toBe('media-show-ok')
      expect(reason(await c.mediaRequest(clicked.mediaId, { unit: 0, ms: 60_000 }))).toBe('media-unavailable')
      expect(reason(await c.mediaRequest(clicked.mediaId, { unit: 1, ms: 60_000 }))).toBe('delivered')
    } finally { c.close(); h.stop() }
  }, 120_000)

  it('a request before show-plan for that unit is refused media-unavailable', async () => {
    const h = await start({ doc: await stillSoundDeck('shared'), media: {} }), c = new Client(h.url)
    try {
      await c.hello(); await c.ready()
      const id = planOf(await c.showPlan(0, 1, 'full'), 0).media!.find(x => x.refused)!.mediaId
      expect((await c.mediaShow(1)).type).toBe('media-show-ok')
      // No plan was sent for unit 1 yet: never served because of ordering.
      expect(reason(await c.mediaRequest(id, { unit: 1, ms: 60_000 }))).toBe('media-unavailable')
      await c.showPlan(1, 1, 'full')
      expect(reason(await c.mediaRequest(id, { unit: 1, ms: 60_000 }))).toBe('delivered')
    } finally { c.close(); h.stop() }
  }, 120_000)

  it('fix round 1 (Important 3): a show-plan reply refused after the build records nothing', async () => {
    const svc = await createMediaServices({ env: process.env, enabled: true, testOnly: { root: mkdtempSync(join(root, 'svc-')) } })
    const real = new Pool({ workers: 1, media: { info: svc.info!, dir: svc.temp.dir } })
    let failMark = false
    const pool = new Proxy(real, { get(t, k) {
      if (k === 'markLayer') return async (...a: Parameters<Pool['markLayer']>) => { if (failMark) throw new Error('mark refused (test)'); return await t.markLayer(...a) }
      const v = Reflect.get(t, k, t); return typeof v === 'function' ? v.bind(t) : v
    } }) as Pool
    const h = await start({ doc: await stillSoundDeck('shared'), pool, policy: { forensicMark: true }, media: { enabled: true, info: svc.info!, services: svc } }), c = new Client(h.url)
    try {
      await c.hello(); await c.ready()
      const id = planOf(await c.showPlan(0, 1, 'full'), 0).media!.find(x => x.refused)!.mediaId
      failMark = true
      // Step 7 (Layer-1) refuses unit 1's reply: no plan was sent for it.
      expect((await c.showPlan(1, 1, 'full')).end?.code).toBe('policy-refused')
      failMark = false
      expect((await c.mediaShow(1)).type).toBe('media-show-ok')
      expect(reason(await c.mediaRequest(id, { unit: 1, ms: 60_000 }))).toBe('media-unavailable')
      expect((await c.showPlan(1, 1, 'full')).end).toBeUndefined()
      expect(reason(await c.mediaRequest(id, { unit: 1, ms: 60_000 }))).toBe('delivered')
    } finally { c.close(); h.stop(); real.dispose(); await svc.close() }
  }, 120_000)

  it('fix round 1 (Important 4): encode-ahead waits for each unit\'s plan and never encodes a use it refuses', async () => {
    const h = await start({ doc: await stillSoundDeck('audio-end'), media: {} }), c = new Client(h.url)
    const entries = () => h.media!.cache.stats().entries + h.media!.cache.stats().inflight
    try {
      await c.hello(); await c.ready()
      // The normal startup order: media-show before show-plan. Nothing is encoded ahead yet.
      const shown = await c.mediaShow(0)
      expect(shown).toMatchObject({ type: 'media-show-ok', queued: 0 })
      await Bun.sleep(500)
      expect(entries()).toBe(0)
      // The still's plan: its narration is encoded ahead, its refused click sound never.
      const p = planOf(await c.showPlan(0, 1, 'full'), 0)
      expect(p.media!.filter(x => x.refused)).toHaveLength(1)
      const end = Date.now() + 20_000
      while (h.media!.cache.stats().entries < 1 && Date.now() < end) await Bun.sleep(50)
      expect(h.media!.cache.stats().entries).toBe(1)
      await Bun.sleep(1500)
      expect(entries()).toBe(1)
    } finally { c.close(); h.stop() }
  }, 120_000)
})
