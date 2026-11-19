/**
 * Task 11b (owner decision 2026-11-19): a still slide plays its sounds. The server answers its show-plan with a
 * sound-only plan (`fallback: 'still'`); the client draws the piece-3 still exactly as for a still refusal and hands
 * the plan's records and no-input schedule to the media player, its transition sound to the transition-sound path.
 * A click advances (the still has no step); click-started sounds ship refused and are never asked for.
 *
 * (Header of the Task 11 suite, whose rig and mocks this file shares:)
 * Piece 5 Task 11 (spec §4.4, §5; plan D4, D12, D16, D18, D19, D22): the remote show plays audio. The media player
 * is mounted on the audience surface with the remote opener (`openRemoteMediaLease` over the session), the
 * transition sound plays through the same opener, `media-show` / `media-cursor` / `media-end` follow the show, a
 * slide's media is requested when it becomes the next or the current slide (paced by the session's `media` log, 6 per
 * 60 s), and audio ends on the show clock (`begin + durationMs`) whatever the delivery does (D4).
 *
 * Driven over the recorded media-deck bundles (`remote-show-media-rig.ts`) with the piece-4 T5 throw-on-import mocks,
 * plus the package's media lease and part reader (`slideshow/media/lease`, `slideshow/media/read`).
 *
 * Task 5 ruling (binding): a slide whose drawing waits on a sound end is a still, with media on as off, so no
 * animated slide has a build after a narration that the narration's end feeds. The D4 case is therefore asserted on
 * the history itself: a runtime refusal leaves the driver's media-end events exactly those of the delivered run,
 * at `begin + durationMs`, the shared computation (`slideMediaEnds`) the classifier and driver evaluate.
 */
import { afterEach, beforeEach, describe, expect, vi } from 'vitest'

const { boom, calls } = vi.hoisted(() => ({
  boom: (m: string) => () => { throw new Error(`T5: package module reached: ${m}`) },
  calls: (m: string) => () => new Proxy({}, {
    get: (_, k) => (k === 'then' || typeof k === 'symbol' || k === '__esModule' ? undefined : () => { throw new Error(`T5: package code called: ${m}.${String(k)}`) }),
    has: () => true,
  }),
}))
vi.mock('../../src/pptx/types', boom('pptx/types'))
// Fix round 1 (minor 5): every transition PlanStage runs goes through `dispatchTransition` (spied, real behaviour).
const { dispatched } = vi.hoisted(() => ({ dispatched: [] as string[] }))
vi.mock('../../src/slideshow/render/three/framework', async (orig) => {
  const real = await orig<typeof import('../../src/slideshow/render/three/framework')>()
  return { ...real, dispatchTransition: (...a: Parameters<typeof real.dispatchTransition>) => { dispatched.push(a[0].element); return real.dispatchTransition(...a) } }
})
vi.mock('../../src/core/zip', boom('core/zip'))
vi.mock('../../src/core/package-sniff', boom('core/package-sniff'))
vi.mock('../../src/slideshow/read/store', boom('read/store'))
vi.mock('../../src/slideshow/read/deck', boom('read/deck'))
vi.mock('../../src/slideshow/prepare/index', boom('prepare/index'))
vi.mock('../../src/slideshow/render/raster', boom('render/raster'))
vi.mock('../../src/slideshow/render/plan', boom('render/plan'))
vi.mock('../../src/slideshow/render/backend', boom('render/backend'))
vi.mock('../../src/slideshow/morph/match', boom('morph/match'))
vi.mock('../../src/slideshow/player/stage', boom('player/stage'))
vi.mock('../../src/slideshow/player/session', boom('player/session'))
vi.mock('../../src/slideshow/player/text-builds', boom('player/text-builds'))
vi.mock('../../src/slideshow/player/hit-test', boom('player/hit-test'))
vi.mock('../../src/slideshow/media/lease', boom('media/lease'))
vi.mock('../../src/slideshow/media/read', boom('media/read'))
vi.mock('../../src/slideshow/render/direct', calls('render/direct'))
vi.mock('../../src/slideshow/render/scene', calls('render/scene'))
vi.mock('../../src/pptx/parse', calls('pptx/parse'))
vi.mock('../../src/core/package-error', calls('core/package-error'))

import { elements, it, mediaDoc, mediaEnds, mediaFixture, run, setUp, show, tearDown, type MediaFixture } from './remote-show-media-rig'
import { decodeShowMessage, SHOW_PLAN_LIMITS, type ShowPlanRecord } from '../../src/stream/show-plan'

const THROW_ON_IMPORT = [
  '../../src/pptx/types', '../../src/core/zip', '../../src/core/package-sniff',
  '../../src/slideshow/read/store', '../../src/slideshow/read/deck', '../../src/slideshow/prepare/index',
  '../../src/slideshow/render/raster', '../../src/slideshow/render/plan', '../../src/slideshow/render/backend',
  '../../src/slideshow/morph/match', '../../src/slideshow/player/stage', '../../src/slideshow/player/session', '../../src/slideshow/player/text-builds', '../../src/slideshow/player/hit-test',
  '../../src/slideshow/media/lease', '../../src/slideshow/media/read',
]
const THROW_ON_CALL = ['../../src/slideshow/render/direct', '../../src/slideshow/render/scene', '../../src/pptx/parse', '../../src/core/package-error']

const SOUND = mediaFixture('still-sound')
const TRANSITION = mediaFixture('still-transition')
const DRIVES = 15_000

beforeEach(() => { dispatched.length = 0; setUp() })
afterEach(() => tearDown())

const audios = (root: ParentNode) => [...root.querySelectorAll('audio, video')] as HTMLMediaElement[]

describe('a still slide plays its sounds (Task 11b)', () => {
  it('the recorded fixtures are sound-only still plans', () => {
    expect(THROW_ON_IMPORT.length + THROW_ON_CALL.length).toBeGreaterThan(0)
    for (const fx of [SOUND, TRANSITION]) {
      expect(fx.plan(0).fallback).toBe('still')
      expect(fx.plans.get(0)).toHaveLength(1)
      expect(fx.list[0]).toMatchObject({ anim: true, media: true })
    }
  })

  it('audio-end still: the piece-3 still is drawn, the narration plays at its begin and ends on the clock; the click sound is never asked for', async () => {
    const p = SOUND.plan(0)
    const narration = p.media!.find((r) => !r.refused)!, clicked = p.media!.find((r) => r.refused)!
    expect(clicked.refused).toBe('media-unavailable')
    const fd = mediaDoc(SOUND)
    const { s, root, diags } = await show(fd)
    await run(150)
    expect(fd.requests()).toEqual([`media-request:${narration.mediaId}@0`])
    // Begins at 250 ms (the no-input schedule), not at entry.
    expect(elements().playing()).toEqual([])
    await run(250)
    // The still path, as for a refusal: the unit raster is painted and the slide plays as its still.
    expect(fd.log).toContain('paint:0')
    expect(diags.some((d) => d.code === 'still')).toBe(true)
    expect(root.getAttribute('data-officeview-show-frame')).toBe('image')
    const els = audios(root)
    expect(els).toHaveLength(1)
    expect(els[0]!.getAttribute('data-officeview-show-media-spid')).toBe('3')
    expect(elements().playing()).toEqual(els)
    await run(narration.durationMs + 100)
    // Its end on the show clock: begin (250) + durationMs.
    expect(mediaEnds(s)).toEqual([[3, 250 + narration.durationMs]])
    expect(elements().playing()).toEqual([])
    // A click advances the still (no step, no trigger): the effect sound never plays and is never requested.
    s.next(); await run(100)
    expect(s.state.slide).toBe('end')
    expect(root.querySelector('audio[data-officeview-show-effect-sound]')).toBeNull()
    expect(fd.requests()).toEqual([`media-request:${narration.mediaId}@0`])
    s.exit()
    expect(elements().revoked).toEqual(elements().created)
  }, DRIVES)

  it('leaving a still stops its sound', async () => {
    const fd = mediaDoc(SOUND)
    const { s } = await show(fd)
    await run(500)
    expect(elements().playing()).toHaveLength(1)
    s.next(); await run(50)
    expect(s.state.slide).toBe('end')
    expect(elements().playing()).toEqual([])
    s.exit()
  }, DRIVES)

  it('a transition sound plays on entry to a still; no transition is drawn', async () => {
    const id = TRANSITION.plan(0).media![0]!.mediaId
    expect(TRANSITION.plan(0).show.transition?.sound?.start?.mediaId).toBe(id)
    const fd = mediaDoc(TRANSITION)
    const { s, root, diags } = await show(fd)
    await run(300)
    expect(fd.log).toContain('paint:0')
    expect(fd.requests()).toEqual([`media-request:${id}@0`])
    const sound = root.querySelector<HTMLMediaElement>('audio[data-officeview-show-sound]')
    expect(sound).not.toBeNull()
    expect(elements().playing()).toEqual([sound])
    // The still enters as piece 3's still does: PlanStage ran no transition (the fade is never dispatched).
    expect(diags.some((d) => d.code === 'transition-approximate')).toBe(false)
    expect(dispatched).toEqual([])
    s.exit()
    // Control (the spy is live): the same fade with its sound on an animated slide is dispatched.
    const b = await show(mediaDoc(mediaFixture('transition-sound')))
    await run(300)
    expect(dispatched).toEqual(['fade'])
    b.s.exit()
  }, DRIVES)

  it('a pointer click on the still advances it (fix round 1, minor 6)', async () => {
    const { s, root } = await show(mediaDoc(SOUND))
    await run(500)
    expect(elements().playing()).toHaveLength(1)
    root.dispatchEvent(new MouseEvent('click', { clientX: 480, clientY: 270, bubbles: true }))
    await run(50)
    expect(s.state.slide).toBe('end')
    expect(elements().playing()).toEqual([])
    s.exit()
  }, DRIVES)

  it('a hostile still plan (click steps, carried sound) is rejected: a click still advances and nothing sounds after exit (fix round 1, Important 1)', async () => {
    const good = SOUND.plan(0)
    const t = good.show.timing as { root: { children: unknown[] } }
    const hostile = { ...good, show: { ...good.show, timing: { ...t, root: { ...t.root, children: [
      ...t.root.children.map((c) => { const w = c as { children: Array<Record<string, unknown>> }; return { ...w, children: w.children.map((l) => ({ ...l, numSld: 99 })) } }),
      { kind: 'par', ctn: { id: 900, durMs: 'indefinite', accel: 0, decel: 0, autoRev: false, nodeType: 'mainSeq', start: [], end: [] }, children: [] },
    ] } } } }
    const json = new TextEncoder().encode(JSON.stringify(hostile))
    const raw = new Uint8Array(1 + json.length); raw[0] = 1; raw.set(json, 1)
    expect(() => decodeShowMessage(raw, { unit: 0, maxDecoded: SHOW_PLAN_LIMITS.showDecoded })).toThrow()
    const fx: MediaFixture = { list: SOUND.list, plans: new Map([[0, [raw]]]), plan: () => hostile as unknown as ShowPlanRecord }
    const fd = mediaDoc(fx)
    const { s, root } = await show(fd)
    await run(600)
    expect(fd.requests()).toEqual([])
    expect(root.querySelectorAll('audio')).toHaveLength(0)
    root.dispatchEvent(new MouseEvent('click', { clientX: 480, clientY: 270, bubbles: true }))
    await run(50)
    expect(s.state.slide).toBe('end')
    expect(elements().playing()).toEqual([])
    s.exit()
  }, DRIVES)

  it('media off: the still is silent and nothing media is sent', async () => {
    const fd = mediaDoc(SOUND, { media: false })
    const { s, root } = await show(fd)
    await run(600)
    expect(fd.log).toContain('paint:0')
    expect(fd.log.filter((l) => l.startsWith('media-'))).toEqual([])
    expect(audios(root)).toEqual([])
    s.exit()
  }, DRIVES)
})
