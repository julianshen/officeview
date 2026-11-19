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
import { flyInDeck, mixedMediaDeck, narrationDeck, repeatDeck, stillSoundDeck, videoDeck } from '../../src/testdata/show-decks'
import { DocumentHost } from '../src/render'
import { Pool } from '../src/pool'
import { createMediaTemp } from '../src/media/temp'
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
