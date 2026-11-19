/**
 * Piece 5 Task 5 (controller ruling, fix round 1): with media on a slide classifies exactly as with media off (every
 * media end seeded, piece 4's `mediaRefused`); the computed audio ends (`media-ends.ts`) only start and stop sound on
 * the client. Two checks replace a proof:
 * - equivalence: every test deck builds the same plan (classification, steps, layers, tracks, marks) with media on
 *   and off; only the media records, the kept commands and the transition sound differ;
 * - playback: on every eligible deck with audio, the client driver with a media clock draws, at sampled times
 *   (just before and after each computed audio end included), exactly what the classified layers reconstruct.
 */
import { afterAll, beforeAll, describe, expect, it } from 'vitest'
import { mkdtempSync, rmSync } from 'node:fs'
import { tmpdir } from 'node:os'
import { join } from 'node:path'
import { installCanvasBackend, type CanvasBackend } from '../../src/core/canvas-backend'
import { ManualClock } from '../../src/slideshow/clock'
import { createDiagnosticSink } from '../../src/slideshow/diagnostics'
import { mediaClockOf, type MediaClockRecord } from '../../src/slideshow/engine/media-ends'
import { createShowDriver, type DriverStage, type SlideSource } from '../../src/slideshow/player/driver'
import type { ShowSlide, TimeNode } from '../../src/slideshow/types'
import type { ShowPlanRecord } from '../../src/stream/show-plan'
import {
  autoplayRevealDeck, clickPauseVideoDeck, cmdHoldDeck, flyInDeck, loadShowDeck, MEDIA_DECKS, mediaEndRevealDeck, mixedMediaDeck, morphDeck, triggerFadeDeck, type MediaEndKind,
} from '../../src/testdata/show-decks'
import { serverBackend } from '../src/backend'
import { createMediaTemp } from '../src/media/temp'
import { DocumentHost } from '../src/render'
import type { ShowPlanJob } from '../src/show/build'
import { ffmpegForTests } from './media/ffmpeg-gate'
import { classify, reconstruct } from './show-classify-helper'

let backend: CanvasBackend
let previous: CanvasBackend | undefined
beforeAll(() => { backend = serverBackend(); previous = installCanvasBackend(backend) })
afterAll(() => { installCanvasBackend(previous) })

const KINDS: MediaEndKind[] = ['onEnd', 'afterPrevious', 'afterIndefinite', 'unrelated', 'endCondition', 'endSync', 'unresolved', 'cycle', 'effectSound']
/** The kinds whose drawing a media end reaches (or cannot be told from): the still, with audio as with video. */
const DEPENDENT = new Set<MediaEndKind>(['onEnd', 'afterPrevious', 'afterIndefinite', 'endCondition', 'endSync', 'unresolved'])

/** Every deck the equivalence covers. */
async function decks(): Promise<Array<[string, Uint8Array]>> {
  const out: Array<[string, Uint8Array]> = []
  for (const [name, f] of Object.entries(MEDIA_DECKS)) out.push([name, await f()])
  for (const kind of KINDS) for (const media of ['video', 'audio'] as const) out.push([`reveal:${kind}:${media}`, await mediaEndRevealDeck(kind, { media })])
  for (const gate of ['video', 'none', 'chain', 'chain-build'] as const) out.push([`mixed:${gate}`, await mixedMediaDeck(gate)])
  out.push(['clickPauseVideo', await clickPauseVideoDeck()], ['flyIn', await flyInDeck()], ['triggerFade', await triggerFadeDeck()], ['cmdHold', await cmdHoldDeck()], ['autoplayReveal', await autoplayRevealDeck()])
  return out
}

/** A plan with what may differ between media on and off taken out: the records, command texts (and `cmd-dropped`), the transition sound. */
function comparable(p: ShowPlanRecord): unknown {
  const strip = (n: TimeNode): TimeNode => ('children' in n ? { ...n, children: n.children.map(strip) } : n.kind === 'cmd' ? { ...n, cmd: '' } : n)
  const timing = p.show.timing && !('refused' in p.show.timing) ? { ...p.show.timing, root: strip(p.show.timing.root) } : p.show.timing
  const transition = p.show.transition ? (({ sound: _s, ...rest }) => rest)(p.show.transition as { sound?: unknown }) : undefined
  const { media: _m, ...rest } = p
  return { ...rest, diagnostics: p.diagnostics.filter(d => d.code !== 'cmd-dropped'), show: { ...p.show, timing, ...(transition ? { transition } : {}) } }
}

const info = await ffmpegForTests()

describe.skipIf(!info)('media on builds exactly what media off builds (controller ruling)', () => {
  const root = mkdtempSync(join(tmpdir(), 'ov-t5e-'))
  let temp: Awaited<ReturnType<typeof createMediaTemp>>
  beforeAll(async () => { temp = await createMediaTemp(root) })
  afterAll(async () => { await temp.remove(); rmSync(root, { recursive: true, force: true }) })
  const job = (unitIndex: number, media: boolean): ShowPlanJob => ({ unitIndex, scale: 1, motion: 'full', mark: { kind: 'fixed' }, strictMarks: false, media })
  const host = () => { const h = new DocumentHost(); h.configureMedia({ info: info!, dir: temp.dir }); return h }

  it('every test deck: identical classification, steps, layers, tracks and marks', async () => {
    let compared = 0
    for (const [name, bytes] of await decks()) {
      const h = host()
      const { docKey, units } = await h.open(`e${Math.random()}`, bytes, 'visible-only')
      for (let u = 0; u < units.length; u++) {
        const off = await h.showPlan(docKey, job(u, false)), on = await h.showPlan(docKey, job(u, true))
        if ('refused' in off || 'refused' in on) {
          expect('refused' in on && on.refused, `${name}#${u}`).toBe('refused' in off && off.refused)
          continue
        }
        expect(comparable(on.plan), `${name}#${u}`).toEqual(comparable(off.plan))
        expect(on.assets.map(a => [a.assetId, a.w, a.h, Buffer.from(a.png).toString('base64')]), `${name}#${u} assets`).toEqual(off.assets.map(a => [a.assetId, a.w, a.h, Buffer.from(a.png).toString('base64')]))
        for (let k = 0; k < on.plan.steps.length; k++) {
          if (on.plan.steps[k]!.eligibility !== 'composited') continue
          const t = (media: boolean) => h.showTrack(docKey, { ...job(u, media), step: k })
          expect((await t(true)).map(f => Buffer.from(f.png).toString('base64')), `${name}#${u} track ${k}`).toEqual((await t(false)).map(f => Buffer.from(f.png).toString('base64')))
        }
        compared++
      }
    }
    expect(compared).toBeGreaterThan(10)
  }, 600_000)

  it("mixedMediaDeck('chain') ships the narration refused media-unavailable and animates; 'chain-build' → still (D22)", async () => {
    const planOf = async (bytes: Uint8Array) => { const h = host(); const { docKey } = await h.open(`c${Math.random()}`, bytes, 'visible-only'); return h.showPlan(docKey, job(0, true)) }
    const chain = await planOf(await mixedMediaDeck('chain'))
    if ('refused' in chain) throw new Error(`refused: ${chain.refused}`)
    expect(chain.plan.media).toEqual([{ mediaId: '@a1.4', kind: 'audio', spid: '4', durationMs: 0, refused: 'media-unavailable' }])
    const build = await planOf(await mixedMediaDeck('chain-build'))
    expect('refused' in build && build.refused).toBe('still')
    // Not chained: the narration keeps its probed duration (and its end reveals the box: the still, as media off).
    const free = await planOf(await mixedMediaDeck('none'))
    expect('refused' in free && free.refused).toBe('still')
  }, 60_000)
})

/** Every audio node of `meta` gets a 2000 ms record, found by its target as the plan's would be. */
function recordsOf(meta: ShowSlide): MediaClockRecord[] {
  const out: MediaClockRecord[] = []
  const walk = (n: TimeNode) => {
    if ('children' in n) { n.children.forEach(walk); return }
    if (n.kind !== 'audio') return
    if ('spid' in n.target) out.push({ mediaId: `@s${n.target.spid}`, kind: 'audio', durationMs: 2000, spid: n.target.spid })
    else if ('snd' in n.target) out.push({ mediaId: n.target.snd, kind: 'audio', durationMs: 2000 })
  }
  if (meta.timing && !('refused' in meta.timing)) walk(meta.timing.root)
  return out
}
const hasAudio = (meta: ShowSlide) => recordsOf(meta).length > 0

/** Plays slide 0 with a media clock on a fake clock: ticks at `times`, Next at `clicks`; returns each tick's cursor. */
async function play(meta: ShowSlide, records: MediaClockRecord[], times: readonly number[], clicks: readonly number[]) {
  const clock = new ManualClock()
  let origin = 0
  const roots = [meta.timing && !('refused' in meta.timing) ? meta.timing.root : undefined]
  const source: SlideSource = {
    count: 1, settings: { showAnimation: true, useTimings: true, loop: false, type: 'present' }, hidden: () => false,
    acquire: async () => ({ slide: meta, release() {} }),
    mediaClock: () => mediaClockOf(roots, records),
  }
  const stage: DriverStage = { prepare() {}, async enter(_m, t) { origin = t }, frame: () => new Map(), flattenCurrent() {}, leave() {}, hitTest: () => undefined, slideTime: (t) => t - origin }
  const driver = createShowDriver({ source, stage, clock, sink: createDiagnosticSink(), reducedMotion: false, kiosk: false, hooks: { onSlideChange() {}, onVisit() {}, onCursor() {}, runLink() {}, onEnd() {} } })
  await driver.start(0)
  const out: Array<{ t: number; cursor: NonNullable<NonNullable<typeof driver.visit>['cursor']> }> = []
  const pending = [...clicks]
  for (const t of times) {
    while (pending.length && pending[0]! <= t) { clock.advance(pending[0]! - clock.now()); driver.next(); pending.shift() }
    clock.advance(t - clock.now())
    driver.tick()
    if (driver.slide !== 0) break
    out.push({ t, cursor: driver.visit!.cursor! })
  }
  driver.dispose()
  return out
}

describe('the client driver with a media clock draws what the classified layers reconstruct', () => {
  it('every MediaEndKind with media audio: the dependent kinds are stills (on as off); the others reconstruct', async () => {
    for (const kind of KINDS) {
      const run = await classify(await mediaEndRevealDeck(kind, { media: 'audio' }))
      if (DEPENDENT.has(kind)) { expect(run.c.fallback, kind).toBe('still'); continue }
      expect(run.c.fallback, kind).toBeUndefined()
    }
  })

  it('every eligible deck with audio: zero pixel difference at sampled times, around each computed end', async () => {
    let played = 0, withEnds = 0
    for (const [name, bytes] of await decks()) {
      const run = await classify(bytes)
      const meta = await (await loadShowDeck(bytes)).meta(0)
      if (run.c.fallback || !hasAudio(meta)) continue
      const records = recordsOf(meta)
      const clicks = run.timeline.steps > 0 ? Array.from({ length: run.timeline.steps }, (_, k) => 300 + 1500 * k) : []
      // Pass 1: where the computed ends fall. Pass 2: sample around them.
      const grid = Array.from({ length: 121 }, (_, i) => i * 50)
      const first = await play(meta, records, grid, clicks)
      const ends = [...new Set(first.at(-1)?.cursor.events.flatMap(e => (e.kind === 'mediaEnd' ? [e.t] : [])) ?? [])]
      // A deck whose audio never begins on the timeline (playFrom's indefinite node) has no end to sample around.
      if (ends.length) withEnds++
      const times = [...new Set([...grid, ...ends.flatMap(e => [e - 1, e, e + 1]).filter(t => t >= 0)])].sort((a, b) => a - b)
      for (const { t, cursor } of await play(meta, records, times, clicks)) {
        expect(reconstruct(run, cursor, t).diff, `${name} @${t}`).toBe(0)
      }
      played++
    }
    expect(played).toBeGreaterThan(3)
    expect(withEnds).toBeGreaterThan(2)
  }, 300_000)
})

describe('Morph pairs with media (piece 4 rule, media on or off)', () => {
  it('a pair whose video end gates drawing refuses morph-fallback; one whose does not plays', async () => {
    const h = new DocumentHost()
    const job = { from: 0, to: 1, fromStep: 0, scale: 1 as const, motion: 'full' as const, mark: { kind: 'fixed' as const } }
    for (const side of ['from', 'to'] as const) {
      const { docKey } = await h.open(`m-${side}`, await morphDeck({ media: side, mediaReveal: true }), 'visible-only')
      expect(await h.showMorph(docKey, job)).toEqual({ refused: 'morph-fallback' })
      const free = await h.open(`f-${side}`, await morphDeck({ media: side }), 'visible-only')
      expect('refused' in await h.showMorph(free.docKey, job)).toBe(false)
    }
  })
})
