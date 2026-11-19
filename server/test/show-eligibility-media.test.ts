/**
 * Piece 5 Task 5 (spec §5; plan D18, D22, X5): the classifier with media on evaluates its histories with the
 * computed audio ends (`media-ends.ts`, the client driver's own computation); video classifies as in piece 4; with
 * media off (no `media` input) everything is piece 4's.
 */
import { afterAll, beforeAll, describe, expect, it } from 'vitest'
import { installCanvasBackend, type CanvasBackend } from '../../src/core/canvas-backend'
import { ManualClock } from '../../src/slideshow/clock'
import { createDiagnosticSink } from '../../src/slideshow/diagnostics'
import { mediaClockOf, type MediaClockRecord } from '../../src/slideshow/engine/media-ends'
import type { SlideCursor } from '../../src/slideshow/engine/steps'
import { createShowDriver, type DriverStage, type SlideSource } from '../../src/slideshow/player/driver'
import type { PreparedSlide } from '../../src/slideshow/prepare'
import { createRng } from '../../src/slideshow/rng'
import type { ShowSlide } from '../../src/slideshow/types'
import {
  clickPauseVideoDeck, loadShowDeck, MEDIA_END_REVEAL, mediaEndRevealDeck, mixedMediaDeck, morphDeck, togglePauseDeck, type MediaEndKind,
} from '../../src/testdata/show-decks'
import { serverBackend } from '../src/backend'
import { createMediaTemp } from '../src/media/temp'
import { DocumentHost } from '../src/render'
import { drawsAt, slideStateAt } from '../src/show/compose'
import { classifierTimeline, mediaRefused, mediaUnprovable, type Classification } from '../src/show/eligibility'
import type { ShowPlanJob } from '../src/show/build'
import { ffmpegForTests } from './media/ffmpeg-gate'
import { classify, reconstruct, type Run } from './show-classify-helper'
import { mkdtempSync, rmSync } from 'node:fs'
import { tmpdir } from 'node:os'
import { join } from 'node:path'

let backend: CanvasBackend
let previous: CanvasBackend | undefined
beforeAll(() => { backend = serverBackend(); previous = installCanvasBackend(backend) })
afterAll(() => { installCanvasBackend(previous) })

const KINDS: MediaEndKind[] = ['onEnd', 'afterPrevious', 'afterIndefinite', 'unrelated', 'endCondition', 'endSync', 'unresolved', 'cycle', 'effectSound']
const rec = (spid: string, durationMs: number): MediaClockRecord => ({ mediaId: `@a0.${spid}`, kind: 'audio', durationMs, spid })
const AUDIO_3000 = [rec('3', 3000)]
const ink = (run: Run) => run.ps.layers.find(l => l.kind === 'target' && l.key === 'sp:2')!
const drawsInk = (run: Run, cursor: SlideCursor, t: number, tl = run.timeline) =>
  drawsAt(run.ps, slideStateAt(run.ps, tl, cursor, t, createRng(1), createDiagnosticSink())).some(d => d.layer === ink(run) && d.state.visible && d.state.opacity > 0)
/** What a classification decides, comparable across two runs (layers by key). */
const summary = (run: Run, c: Classification = run.c) => {
  const key = (l: object) => (l as { key?: string }).key ?? (l as { kind: string }).kind
  return { fallback: c.fallback, steps: c.steps, drop: [...c.drop].map(key).sort(), masks: [...c.masks.keys()].map(key).sort(), triggers: [...c.triggers], diagnostics: c.diagnostics }
}
const rootsOf = (ps: PreparedSlide) => [ps.meta.timing && !('refused' in ps.meta.timing) ? ps.meta.timing.root : undefined, ps.textTiming?.tree.root]

describe('classification with media on (Task 5)', () => {
  it("mediaEndRevealDeck('onEnd', {media:'audio'}) with a 3000 ms record classifies layers and the reveal fires at 3000", async () => {
    const run = await classify(await mediaEndRevealDeck('onEnd', { media: 'audio' }), { media: { records: AUDIO_3000 } })
    expect(run.c.fallback).toBeUndefined()
    expect(run.c.drop.has(ink(run))).toBe(false)
    const tl = classifierTimeline(run.ps, run.timeline, { records: AUDIO_3000 })
    expect(drawsInk(run, tl.initial(), 2999, tl)).toBe(false)
    expect(drawsInk(run, tl.initial(), 3000, tl)).toBe(true)
    // The classified layers reproduce the revealed state.
    const ended = run.timeline.mediaEnd(run.timeline.initial(), MEDIA_END_REVEAL.videoId, 3000, false)
    expect(reconstruct(run, ended, 3500).diff).toBe(0)
    // Media off: piece 4's still.
    expect((await classify(await mediaEndRevealDeck('onEnd', { media: 'audio' }))).c.fallback).toBe('still')
  })

  it('media on: mediaEndRevealDeck(kind) (video) classifies exactly as piece 4 for every MediaEndKind', async () => {
    for (const kind of KINDS) {
      const bytes = await mediaEndRevealDeck(kind)
      const off = await classify(bytes), on = await classify(bytes, { media: { records: [] } })
      expect(summary(on), kind).toEqual(summary(off))
    }
  })

  it('mixed slide: a video end that gates a build → still; one that gates nothing → the audio ends on the clock', async () => {
    const records = [rec('4', 3000)]
    const gated = await classify(await mixedMediaDeck('video'), { media: { records } })
    expect(gated.c.fallback).toBe('still')
    expect(mediaUnprovable(gated.ps, records)).toBe(true)
    const free = await classify(await mixedMediaDeck('none'), { media: { records } })
    expect(free.c.fallback).toBeUndefined()
    expect(mediaUnprovable(free.ps, records)).toBe(false)
    const tl = classifierTimeline(free.ps, free.timeline, { records })
    expect(drawsInk(free, tl.initial(), 2999, tl)).toBe(false)
    expect(drawsInk(free, tl.initial(), 3000, tl)).toBe(true)
  })

  it("media on: mediaEndRevealDeck('unrelated') and clickPauseVideoDeck classify as with media off (not stills)", async () => {
    for (const bytes of [await mediaEndRevealDeck('unrelated'), await clickPauseVideoDeck()]) {
      const off = await classify(bytes), on = await classify(bytes, { media: { records: [] } })
      expect(off.c.fallback).toBeUndefined()
      expect(summary(on)).toEqual(summary(off))
    }
  })

  it('togglePauseDeck (interactive) → still with reason media', async () => {
    const run = await classify(await togglePauseDeck(), { media: { records: AUDIO_3000 } })
    expect(run.c.fallback).toBe('still')
    expect(mediaUnprovable(run.ps, AUDIO_3000)).toBe(true)
    // Piece 4 (media off) never saw a reason here: nothing drawn waits on the audio.
    expect(mediaRefused(run.ps)).toBe(false)
  })

  it("mediaEndRevealDeck('unresolved', {media:'audio'}) → still", async () => {
    const run = await classify(await mediaEndRevealDeck('unresolved', { media: 'audio' }), { media: { records: AUDIO_3000 } })
    expect(run.c.fallback).toBe('still')
    expect(mediaUnprovable(run.ps, AUDIO_3000)).toBe(true)
  })

  it('every MediaEndKind with media audio: classifier history equals driver history at sampled times', async () => {
    for (const kind of KINDS) {
      const d = await loadShowDeck(await mediaEndRevealDeck(kind, { media: 'audio' }))
      const meta = await d.meta(0)
      const run = await classify(await mediaEndRevealDeck(kind, { media: 'audio' }), { media: { records: AUDIO_3000 } })
      const tl = classifierTimeline(run.ps, run.timeline, { records: AUDIO_3000 })
      // One fake clock drives the client driver over the same slide, with the same records.
      const clock = new ManualClock()
      let origin = 0
      const source: SlideSource = {
        count: 1, settings: { showAnimation: true, useTimings: true, loop: false, type: 'present' }, hidden: () => false,
        acquire: async () => ({ slide: meta as ShowSlide, release() {} }),
        mediaClock: () => mediaClockOf(rootsOf(run.ps), AUDIO_3000),
      }
      const stage: DriverStage = {
        prepare() {}, async enter(_m, t) { origin = t }, frame: () => new Map(), flattenCurrent() {}, leave() {}, hitTest: () => undefined, slideTime: (t) => t - origin,
      }
      const driver = createShowDriver({ source, stage, clock, sink: createDiagnosticSink(), reducedMotion: false, kiosk: false, hooks: { onSlideChange() {}, onVisit() {}, onCursor() {}, runLink() {}, onEnd() {} } })
      await driver.start(0)
      const dtl = driver.visit!.timeline!
      for (const t of [0, 16, 20, 500, 1000, 1020, 2999, 3000, 3001, 3020, 4000, 5000, 6000, 8000, 8001, 9000]) {
        clock.advance(t - clock.now())
        driver.tick()
        const ids = (s: { active: ReadonlyArray<{ node: { ctn: { id: number } } }> }) => s.active.map(e => e.node.ctn.id).sort((a, b) => a - b)
        expect(ids(dtl.stateAt(driver.visit!.cursor!, t)), `${kind} @${t}`).toEqual(ids(tl.stateAt(tl.initial(), t)))
      }
      driver.dispose()
    }
  })

  it('can.media false (no media input): mediaEndRevealDeck classifies exactly as piece 4, still with reason media', async () => {
    for (const kind of ['onEnd', 'afterPrevious', 'afterIndefinite'] as const) {
      const run = await classify(await mediaEndRevealDeck(kind))
      expect(run.c.fallback).toBe('still')
      expect(mediaRefused(run.ps)).toBe(true)
      // No input, no computed end: the classifier's timeline is the plain one.
      expect(classifierTimeline(run.ps, run.timeline)).toBe(run.timeline)
    }
    const audio = await classify(await mediaEndRevealDeck('onEnd', { media: 'audio' }))
    expect(audio.c.fallback).toBe('still')
    expect(mediaRefused(audio.ps)).toBe(true)
  })
})

const info = await ffmpegForTests()

describe('show builds with media on (Task 5)', () => {
  const morphJob = (media?: boolean) => ({ from: 0, to: 1, fromStep: 0, scale: 1 as const, motion: 'full' as const, mark: { kind: 'fixed' as const }, ...(media === undefined ? {} : { media }) })

  it('media on: a Morph pair whose video end gates drawing refuses morph-fallback', async () => {
    const h = new DocumentHost()
    for (const side of ['from', 'to'] as const) {
      const { docKey } = await h.open(`m-${side}`, await morphDeck({ media: side, mediaReveal: true }), 'visible-only')
      expect(await h.showMorph(docKey, morphJob(true))).toEqual({ refused: 'morph-fallback' })
      const free = await h.open(`f-${side}`, await morphDeck({ media: side }), 'visible-only')
      expect('refused' in await h.showMorph(free.docKey, morphJob(true))).toBe(false)
    }
  })

  it('media off: a Morph pair with media refuses morph-fallback as in piece 4', async () => {
    const h = new DocumentHost()
    for (const side of ['from', 'to'] as const) {
      const { docKey } = await h.open(`o-${side}`, await morphDeck({ media: side, mediaReveal: true }), 'visible-only')
      expect(await h.showMorph(docKey, morphJob())).toEqual({ refused: 'morph-fallback' })
      expect(await h.showMorph(docKey, morphJob(false))).toEqual({ refused: 'morph-fallback' })
    }
  })

  describe.skipIf(!info)('with ffmpeg', () => {
    const root = mkdtempSync(join(tmpdir(), 'ov-t5-'))
    let temp: Awaited<ReturnType<typeof createMediaTemp>>
    beforeAll(async () => { temp = await createMediaTemp(root) })
    afterAll(async () => { await temp.remove(); rmSync(root, { recursive: true, force: true }) })
    const job = (media: boolean): ShowPlanJob => ({ unitIndex: 0, scale: 1, motion: 'full', mark: { kind: 'fixed' }, strictMarks: false, media })
    const planOf = async (bytes: Uint8Array, media: boolean) => {
      const h = new DocumentHost()
      h.configureMedia({ info: info!, dir: temp.dir })
      const { docKey } = await h.open(`d${Math.random()}`, bytes, 'visible-only')
      return h.showPlan(docKey, job(media))
    }

    it("mixedMediaDeck('chain') ships the narration refused media-unavailable and animates; 'chain-build' → still (D22)", async () => {
      const chain = await planOf(await mixedMediaDeck('chain'), true)
      if ('refused' in chain) throw new Error(`refused: ${chain.refused}`)
      expect(chain.plan.media).toEqual([{ mediaId: '@a1.4', kind: 'audio', spid: '4', durationMs: 0, refused: 'media-unavailable' }])
      const build = await planOf(await mixedMediaDeck('chain-build'), true)
      expect('refused' in build && build.refused).toBe('still')
      // Not chained: the narration keeps its probed duration.
      const free = await planOf(await mixedMediaDeck('none'), true)
      if ('refused' in free) throw new Error(`refused: ${free.refused}`)
      expect(free.plan.media![0]!.durationMs).toBeGreaterThan(0)
      expect(free.plan.media![0]!.refused).toBeUndefined()
    }, 60_000)
  })
})
