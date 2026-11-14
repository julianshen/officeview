/**
 * Piece 4 Task 11: the remote show on the shared driver (`runRemoteShow` over
 * `createPlanSource` and `PlanStage`), driven by a fake session that serves the
 * Task 9 recorded bundles (`tests/fixtures/show-plan/{fly,morph}.bin`) as raw
 * kind-6 messages, so the real decoders (`decodeBundle`, `decodeTrack`,
 * `decodeMorph`) run. Variants (more slides, composited steps with keyframe
 * tracks, link regions, other Morph pairs) are re-encoded from the recorded
 * records with the client codec.
 *
 * The Task 9 throw-on-import mocks apply, except `render/direct`,
 * `render/scene` and `pptx/parse`, which piece 3's show chunk already imports
 * statically (D25), and `core/package-error`, which piece 3 also reaches
 * (`presenter/preview.ts:42`): those are replaced by modules whose every
 * export throws when called.
 *
 * jsdom with node-canvas; PlanStage on Canvas 2D; the show clock, the fetcher
 * pacing and the frame loop run on an injected manual clock.
 */
import { createCanvas, loadImage } from 'canvas'
import { readFileSync } from 'node:fs'
import { resolve } from 'node:path'
import { afterEach, beforeEach, describe, expect, it, vi } from 'vitest'

const { boom, calls } = vi.hoisted(() => ({
  boom: (m: string) => () => { throw new Error(`T5: package module reached: ${m}`) },
  /** A module whose every export throws when called (never on import). */
  calls: (m: string) => () => new Proxy({}, {
    get: (_, k) => (k === 'then' || typeof k === 'symbol' || k === '__esModule' ? undefined : () => { throw new Error(`T5: package code called: ${m}.${String(k)}`) }),
    has: () => true,
  }),
}))
vi.mock('../../src/pptx/types', boom('pptx/types'))
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
vi.mock('../../src/slideshow/render/direct', calls('render/direct'))
vi.mock('../../src/slideshow/render/scene', calls('render/scene'))
vi.mock('../../src/pptx/parse', calls('pptx/parse'))
// Not in D25's list, but piece 3 reaches it too (`presenter/preview.ts:42`, `isPackageError`), never on the remote path.
vi.mock('../../src/core/package-error', calls('core/package-error'))

import { remoteShowDeps, runRemoteShow, type RemoteShowStart } from '../../src/stream/remote/show'
import { StreamSessionError, type StreamErrorCode } from '../../src/stream/remote/errors'
import type { Manifest } from '../../src/stream/manifest'
import type { RemoteOfficeDoc } from '../../src/stream/remote/remote-doc'
import type { ShowReply, ShowRequestOptions } from '../../src/stream/remote/session'
import { decodeBundle, decodeMorph, decodeTrack } from '../../src/stream/remote/show-bundle'
import { decodeShowMessage, encodeShowMessage, SHOW_PLAN_LIMITS, type ShowMessage, type ShowPlanRecord } from '../../src/stream/show-plan'
import { createSlideShowWorkspace } from '../../src/slideshow/workspace'
import type { TimeNode } from '../../src/slideshow/types'

const THROW_ON_IMPORT = [
  '../../src/pptx/types', '../../src/core/zip', '../../src/core/package-sniff',
  '../../src/slideshow/read/store', '../../src/slideshow/read/deck', '../../src/slideshow/prepare/index',
  '../../src/slideshow/render/raster', '../../src/slideshow/render/plan', '../../src/slideshow/render/backend',
  '../../src/slideshow/morph/match', '../../src/slideshow/player/stage', '../../src/slideshow/player/session', '../../src/slideshow/player/text-builds', '../../src/slideshow/player/hit-test',
]
const THROW_ON_CALL = ['../../src/slideshow/render/direct', '../../src/slideshow/render/scene', '../../src/pptx/parse', '../../src/core/package-error']

// ---------------------------------------------------------------------------
// The recorded messages, and variants re-encoded from them
// ---------------------------------------------------------------------------

interface Rec { op: number; unit: number; bytes: Uint8Array }
function records(name: string): Rec[] {
  const buf = new Uint8Array(readFileSync(resolve(__dirname, '../fixtures/show-plan', `${name}.bin`)))
  const v = new DataView(buf.buffer, buf.byteOffset, buf.byteLength)
  const out: Rec[] = []
  for (let at = 0; at < buf.length;) {
    const len = v.getUint32(at + 13, true)
    out.push({ op: v.getUint8(at), unit: v.getUint32(at + 1, true), bytes: buf.subarray(at + 17, at + 17 + len) })
    at += 17 + len
  }
  return out
}
const decodeAll = (rs: Rec[]) => rs.map((r) => decodeShowMessage(r.bytes, { unit: r.unit, maxDecoded: SHOW_PLAN_LIMITS.showDecoded }))
const FLY = decodeAll(records('fly').filter((r) => r.op === 1))
const MORPH_RECS = records('morph')
/** The Morph target slide's bundle (a transition-only slide: plan + entry frame). */
const MORPH_TO = decodeAll(MORPH_RECS.filter((r) => r.op === 1))
/** The recorded layered pair (0 → 1 at 0): the tag-5 record and its four assets. */
const PAIR = decodeAll(MORPH_RECS.filter((r) => r.op === 3))

type Plan = ShowPlanRecord
/** A recorded bundle re-addressed to `unit` (the plan's index and every asset's unit), the plan optionally edited. */
function bundleOf(src: ShowMessage[], unit: number, edit: (p: Plan) => Plan = (p) => p): Uint8Array[] {
  return src.map((m) => (m.tag === 1
    ? encodeShowMessage({ tag: 1, plan: edit({ ...m.plan, slide: { ...m.plan.slide, index: unit } }) })
    : encodeShowMessage({ ...(m as Extract<ShowMessage, { tag: 2 | 3 | 4 }>), unit })))
}
/** The recorded pair re-addressed to `(from → to @ fromStep)`. */
function pairOf(from: number, to: number, fromStep: number): Uint8Array[] {
  return PAIR.map((m) => (m.tag === 5
    ? encodeShowMessage({ tag: 5, morph: { ...m.morph, from, to, fromStep } })
    : encodeShowMessage({ ...(m as Extract<ShowMessage, { tag: 2 | 3 | 4 }>), unit: to })))
}

/** The fly-in timing with its one click step repeated `n` times (fresh node ids). */
function flyTiming(n: number): Plan['show']['timing'] {
  const plan = (FLY[0] as Extract<ShowMessage, { tag: 1 }>).plan
  const timing = plan.show.timing!
  if ('refused' in timing) throw new Error('refused')
  const root = structuredClone(timing.root) as TimeNode & { children: TimeNode[] }
  const main = root.children[0] as TimeNode & { children: TimeNode[] }
  const click = main.children[0]!
  const renumber = (node: TimeNode, k: number): TimeNode => {
    const c = structuredClone(node) as TimeNode
    const walk = (x: TimeNode) => { (x.ctn as { id: number }).id += 1000 * k; if ('children' in x) for (const y of x.children) walk(y) }
    walk(c)
    return c
  }
  main.children = Array.from({ length: n }, (_, k) => (k === 0 ? click : renumber(click, k)))
  return { ...timing, root }
}

const FRAME_W = 96, FRAME_H = 72
const COLORS = ['#ff0000', '#00ff00', '#0000ff']
const png = (color: string, w = FRAME_W, h = FRAME_H) => {
  const c = createCanvas(w, h)
  const ctx = c.getContext('2d')
  ctx.fillStyle = color
  ctx.fillRect(0, 0, w, h)
  return new Uint8Array(c.toBuffer('image/png'))
}
/** A fly slide whose `steps` click steps are all composited, each with a 3-frame keyframe track. */
const compositedFly = (steps: number) => (p: Plan): Plan => ({
  ...p,
  show: { ...p.show, timing: flyTiming(steps) },
  steps: Array.from({ length: steps }, () => ({ trigger: 'click' as const, eligibility: 'composited' as const, track: { fps: 15 as const, count: 3 } })),
})
const trackOf = (unit: number, step: number) => [0, 1, 2].map((i) => encodeShowMessage({ tag: 4, unit, assetId: i, w: FRAME_W, h: FRAME_H, png: png(COLORS[step % 3]!) }))

// ---------------------------------------------------------------------------
// The fake streamed document and session
// ---------------------------------------------------------------------------

type Kind = 'still' | 'fly' | 'morphTo' | { messages: (unit: number) => Uint8Array[] }
type Rule = 'hold' | StreamErrorCode | (() => Uint8Array[])

const flush = async () => { for (let i = 0; i < 6; i++) await new Promise((r) => setTimeout(r, 0)) }

function animDoc(deck: Kind[], o: { tracks?: (unit: number, step: number) => Uint8Array[] } = {}) {
  const list = deck.map((k) => (k === 'still' ? { w: 960, h: 720 } : { w: 960, h: 720, anim: true as const }))
  const manifest = {
    units: { count: deck.length, list },
    can: { search: false, copy: false, print: false, textA11y: false, present: true, notes: false, links: false },
    recorder: { scalesBuckets: [1, 2] },
  } as unknown as Manifest
  const log: string[] = []
  const rules = new Map<string, Rule>()
  const held = new Map<string, Array<() => void>>()
  const signals = new Map<string, AbortSignal>()
  let endIt!: (e: StreamSessionError | undefined) => void
  const ended = new Promise<StreamSessionError | undefined>((r) => { endIt = r })
  const messagesFor = (unit: number): Uint8Array[] => {
    const k = deck[unit]!
    if (k === 'fly') return bundleOf(FLY, unit)
    if (k === 'morphTo') return bundleOf(MORPH_TO, unit)
    if (k === 'still') throw new Error(`unit ${unit} is a still`)
    return k.messages(unit)
  }
  /** One show reply: logged, charged once, each message leased at its bytes in the show workspace (as the session). */
  function reply(key: string, sig: AbortSignal, o: ShowRequestOptions, make: () => Uint8Array[]): Promise<ShowReply> {
    log.push(key)
    signals.set(key, sig)
    o.onSend?.()
    const rule = rules.get(key)
    return new Promise<ShowReply>((res, rej) => {
      const answer = () => {
        if (sig.aborted) { rej(sig.reason ?? new Error('aborted')); return }
        const r = rules.get(key)
        if (typeof r === 'string' && r !== 'hold') { rej(new StreamSessionError(r)); return }
        const messages = typeof r === 'function' ? r() : make()
        const leases = messages.map((m) => o.workspace.lease(m.length, 'remote-show:message'))
        if (leases.some((l) => !l)) { for (const l of leases) l?.release(); rej(new StreamSessionError('workspace-full')); return }
        let freed = false
        res({ messages, release: () => { if (freed) return; freed = true; for (const l of leases) l!.release() } })
      }
      sig.addEventListener('abort', () => rej(sig.reason ?? new Error('aborted')), { once: true })
      if (rule === 'hold') { const q = held.get(key) ?? []; q.push(answer); held.set(key, q) } else queueMicrotask(answer)
    })
  }
  const session = {
    showPlan: vi.fn((unit: number, scale: number, motion: string, sig: AbortSignal, so: ShowRequestOptions) => {
      log.push(`motion:${motion}@${scale}`)
      return reply(`plan:${unit}`, sig, so, () => messagesFor(unit))
    }),
    showTrack: vi.fn((unit: number, step: number, _scale: number, _motion: string, sig: AbortSignal, so: ShowRequestOptions) =>
      reply(`track:${unit}:${step}`, sig, so, () => (o.tracks ?? trackOf)(unit, step))),
    showMorph: vi.fn((from: number, to: number, fromStep: number, _scale: number, _motion: string, sig: AbortSignal, so: ShowRequestOptions) =>
      reply(`morph:${from}>${to}@${fromStep}`, sig, so, () => pairOf(from, to, fromStep))),
  }
  const doc = {
    kind: 'remote',
    get manifest() { return manifest },
    session,
    paint: vi.fn(async (unit: number, scale: number | 'fit', dpr: number, canvas: HTMLCanvasElement, signal?: AbortSignal, po?: { onSend?: () => void }) => {
      log.push(`paint:${unit}`)
      po?.onSend?.()
      await Promise.resolve()
      if (signal?.aborted) return
      const k = (typeof scale === 'number' ? scale : 1) * dpr
      canvas.width = 960 * k; canvas.height = 720 * k
      const ctx = canvas.getContext('2d')!
      ctx.fillStyle = '#123456'
      ctx.fillRect(0, 0, canvas.width, canvas.height)
    }),
    paintNotes: vi.fn(() => Promise.reject(new StreamSessionError('policy-refused'))),
    onManifestChange() { return () => {} },
    ended,
    close: vi.fn(),
  } as unknown as RemoteOfficeDoc
  return {
    doc, session, log, rules, signals,
    release(key: string) { const q = held.get(key) ?? []; held.delete(key); for (const f of q) f() },
    end: async (e?: StreamSessionError) => { endIt(e); await flush() },
  }
}

/** The fetcher, notes and show clock: time moves only through `advance`. */
function manualClock() {
  let t = 0
  let timers: Array<{ at: number; cb: () => void }> = []
  return {
    now: () => t,
    wake(ms: number, cb: () => void) {
      const e = { at: t + ms, cb }
      timers.push(e)
      return () => { timers = timers.filter((x) => x !== e) }
    },
    advance(ms: number) {
      const end = t + ms
      for (;;) {
        const due = timers.filter((x) => x.at <= end).sort((a, b) => a.at - b.at)[0]
        if (!due) break
        timers = timers.filter((x) => x !== due)
        t = Math.max(t, due.at)
        due.cb()
      }
      t = end
    },
  }
}

let clock: ReturnType<typeof manualClock>
let frame: (() => void) | undefined
const bitmapOf = async (bytes: Uint8Array) => {
  const img = await loadImage(Buffer.from(bytes))
  const cv = document.createElement('canvas')
  cv.width = img.width; cv.height = img.height
  cv.getContext('2d')!.drawImage(img as unknown as CanvasImageSource, 0, 0)
  return cv
}
beforeEach(() => {
  clock = manualClock()
  frame = undefined
  vi.spyOn(remoteShowDeps, 'clock').mockImplementation(() => clock)
  vi.spyOn(remoteShowDeps, 'every').mockImplementation(() => () => {})
  vi.spyOn(remoteShowDeps, 'frames').mockImplementation((_w, f) => { frame = f; return () => { frame = undefined } })
  vi.spyOn(remoteShowDeps, 'decode').mockImplementation(bitmapOf)
  remoteShowDeps.backend = 'canvas2d'
  Object.defineProperty(window, 'innerWidth', { configurable: true, value: 960 })
  Object.defineProperty(window, 'innerHeight', { configurable: true, value: 540 })
})
afterEach(() => {
  remoteShowDeps.backend = undefined
  vi.restoreAllMocks()
  document.body.replaceChildren()
})

/** Runs the show `ms` at ~60 fps: the clock, then one frame, letting decodes settle. */
async function run(ms: number) {
  for (let t = 0; t < ms; t += 16) {
    clock.advance(16)
    frame?.()
    await flush()
  }
}

async function show(fd: ReturnType<typeof animDoc>, o: Partial<RemoteShowStart> = {}) {
  const host = document.createElement('div')
  host.tabIndex = -1
  document.body.append(host)
  const ac = new AbortController()
  const log: string[] = []
  const start: RemoteShowStart = {
    doc: fd.doc, host,
    token: { release() {}, released: false, owns: () => false },
    platform: { hasActivation: () => true, requestFullscreen: async () => {}, exitFullscreen: async () => {}, fullscreenElement: () => null, open: () => null },
    fullscreen: Promise.resolve('entered'),
    retryFullscreen: async () => 'entered',
    from: 'beginning', motion: 'full',
    callbacks: {
      onStart: () => log.push('onStart'),
      onEnd: () => log.push('onEnd'),
      onSlideChange: (i, b) => log.push(`slide:${i}:${b}`),
      onDiagnostic: (d) => log.push(`diag:${d.code}`),
      currentSlide: () => 0, restoreFocus() {}, scrollToSlide: () => {},
    },
    signal: ac.signal, generation: 1, cleanup: () => { ac.abort() },
    ...o,
  }
  const s = await runRemoteShow(start)
  await flush()
  const canvas = host.querySelector<HTMLCanvasElement>('canvas[data-officeview-show-canvas]')!
  const root = host.querySelector<HTMLElement>('[data-officeview-show-surface]')!
  root.getBoundingClientRect = () => ({ left: 0, top: 0, width: canvas.width, height: canvas.height, right: canvas.width, bottom: canvas.height, x: 0, y: 0, toJSON: () => ({}) })
  return { s, host, log, canvas, root }
}

const frameKind = (root: HTMLElement) => root.getAttribute('data-officeview-show-frame')
const px = (c: HTMLCanvasElement, x: number, y: number) => [...c.getContext('2d')!.getImageData(x, y, 1, 1).data].slice(0, 3)
const sent = (fd: ReturnType<typeof animDoc>, prefix: string) => fd.log.filter((l) => l.startsWith(prefix))

describe('remote show on the driver (piece 4 Task 11)', () => {
  it('fly-in deck: step count and order equal the plan; clicks to end = steps + slides', async () => {
    const fd = animDoc(['fly', 'fly'])
    const { s, root, log } = await show(fd)
    await run(1_200)
    expect(frameKind(root)).toBe('image')
    const plan = (FLY[0] as Extract<ShowMessage, { tag: 1 }>).plan
    let clicks = 0
    while (s.state.slide !== 'end' && clicks < 10) { s.next(); clicks++; await run(1_200) }
    expect(s.state.slide).toBe('end')
    expect(clicks).toBe(2 * plan.steps.length + 2)
    // Every step of every slide, in order: each slide's build 0, then its click steps.
    expect(log.filter((l) => l.startsWith('slide:'))).toEqual(['slide:0:0', 'slide:0:1', 'slide:1:0', 'slide:1:1'])
    expect(log.filter((l) => l.startsWith('diag:'))).not.toContain('diag:plan-stage-pressure')
    s.next()
    await s.ended
    expect(log).toContain('onEnd')
  })

  it('unanimated slide requests a unit, never a show-plan', async () => {
    const fd = animDoc(['still', 'fly'])
    const { s, root } = await show(fd)
    await run(200)
    expect(sent(fd, 'paint:')).toEqual(['paint:0'])
    expect(frameKind(root)).toBe('image')
    s.next()
    await run(1_200)
    expect(s.state.slide).toBe(1)
    expect(fd.session.showPlan.mock.calls.map((c) => c[0])).not.toContain(0)
    expect(sent(fd, 'plan:')).toEqual(['plan:1'])
    s.exit()
  })

  it('prefetch +2/+1 bundles, target first', async () => {
    const fd = animDoc(['fly', 'fly', 'fly', 'fly', 'fly'])
    const { s } = await show(fd, { from: 2 })
    expect(sent(fd, 'plan:')).toEqual(['plan:2'])
    await run(1_500)
    expect(sent(fd, 'plan:')).toEqual(['plan:2', 'plan:3', 'plan:4', 'plan:1'])
    s.exit()
  })

  it('tracks fetched just in time (D4)', async () => {
    const three = { messages: (u: number) => bundleOf(FLY, u, compositedFly(3)) }
    const fd = animDoc([three, three])
    const { s, canvas } = await show(fd)
    await run(1_500)
    // Entering: the first two composited steps' tracks; the prefetched neighbour fetches its bundle only.
    expect(sent(fd, 'track:')).toEqual(['track:0:0', 'track:0:1'])
    expect(sent(fd, 'plan:')).toEqual(['plan:0', 'plan:1'])
    s.next()
    await run(100)
    // Step 1 plays from its track (a red keyframe), never from the layers.
    expect(px(canvas, canvas.width >> 1, canvas.height >> 1)).toEqual([255, 0, 0])
    expect(sent(fd, 'track:')).toEqual(['track:0:0', 'track:0:1'])
    await run(400)
    // Starting step 2 asks for step 3's track.
    s.next()
    await run(100)
    expect(px(canvas, canvas.width >> 1, canvas.height >> 1)).toEqual([0, 255, 0])
    expect(sent(fd, 'track:')).toEqual(['track:0:0', 'track:0:1', 'track:0:2'])
    s.exit()
  })

  it('a track with no budget room plays its step as a cut (step-deferred)', async () => {
    const three = { messages: (u: number) => bundleOf(FLY, u, compositedFly(1)) }
    const fd = animDoc([three])
    // The bundle takes the last send: the track is never asked for.
    Object.assign(fd.doc, { unitBudget: { room: () => (fd.log.includes('plan:0') ? 0 : 1), waitMs: () => 60_000 } })
    const { s, log } = await show(fd)
    await run(200)
    s.next()
    await run(300)
    expect(sent(fd, 'track:')).toEqual([])
    expect(log).toContain('diag:step-deferred')
    s.exit()
  })

  it('budget at the Morph reserve: the target bundle still goes out, the Morph pair is not asked and fades (Task 12)', async () => {
    const fd = animDoc(['fly', 'morphTo', 'fly'])
    // 3 sends of room: over a track's reserve, under a Morph pair's, over a target's need.
    Object.assign(fd.doc, { unitBudget: { room: () => 3, waitMs: () => 60_000 } })
    const { s, log } = await show(fd, { from: 2 })
    await run(1_500)
    s.previous()
    await run(300)
    expect(s.state.slide).toBe(1)
    expect(sent(fd, 'plan:')).toContain('plan:1')
    expect(sent(fd, 'morph:')).toEqual([])
    await run(600)
    expect(log).toContain('diag:morph-late')
    s.exit()
  })

  it('Previous into a Morph slide fetches its own pair or fades at 500 ms', async () => {
    {
      const fd = animDoc(['fly', 'morphTo', 'fly'])
      const { s, log } = await show(fd, { from: 2 })
      await run(1_500)
      s.previous()
      await run(300)
      expect(s.state.slide).toBe(1)
      expect(sent(fd, 'morph:')).toEqual(['morph:2>1@0'])
      await run(1_500)
      expect(log).not.toContain('diag:morph-late')
      // The Morph played (its approximation note is local mode's too), never the fade.
      expect(s.presenterSource.diagnostics().filter((d) => d.effect === 'morph-fallback')).toEqual([])
      s.exit()
    }
    {
      const fd = animDoc(['fly', 'morphTo', 'fly'])
      fd.rules.set('morph:2>1@0', 'hold')
      const { s, log } = await show(fd, { from: 2 })
      await run(1_500)
      s.previous()
      await run(300)
      expect(log).not.toContain('diag:morph-late')
      await run(300)
      expect(log).toContain('diag:morph-late')
      s.exit()
    }
  })

  it('goTo into a Morph slide sends fromStep = cursor.step of the slide left', async () => {
    const fd = animDoc(['fly', 'morphTo', 'fly'])
    const { s } = await show(fd, { from: 2 })
    await run(1_500)
    s.next()
    await run(800)
    expect(s.state.build).toBe(1)
    s.goTo(1)
    await run(300)
    expect(sent(fd, 'morph:')).toEqual(['morph:2>1@1'])
    s.exit()
  })

  it('the sequential predecessor\'s pair is prefetched with the slide at fromStep = steps(prev); a different left step re-requests', async () => {
    {
      const fd = animDoc(['fly', 'morphTo'])
      const { s, log } = await show(fd)
      await run(1_500)
      // Slide 1's bundle is in: its pair from slide 0 is asked for at slide 0's last step (1).
      expect(sent(fd, 'morph:')).toEqual(['morph:0>1@1'])
      s.next(); await run(800)
      s.next(); await run(1_500)
      expect(s.state.slide).toBe(1)
      // Left at step 1: the prefetched pair plays; nothing more is asked.
      expect(sent(fd, 'morph:')).toEqual(['morph:0>1@1'])
      expect(log).not.toContain('diag:morph-late')
      s.exit()
    }
    {
      const fd = animDoc(['fly', 'morphTo'])
      const { s } = await show(fd)
      await run(1_500)
      expect(sent(fd, 'morph:')).toEqual(['morph:0>1@1'])
      // Left at step 0: another pair; the prefetched one is dropped.
      s.goTo(1)
      await run(300)
      expect(sent(fd, 'morph:')).toEqual(['morph:0>1@1', 'morph:0>1@0'])
      s.exit()
    }
  })

  it('revocation mid-Morph releases every lease and draws nothing after ended', async () => {
    const ws = createSlideShowWorkspace('desktop', { limit: 512 * 1024 * 1024, scratch: 0 })
    vi.spyOn(remoteShowDeps, 'workspace').mockImplementation(() => ws)
    const fd = animDoc(['fly', 'morphTo'])
    // The successor's pair is prefetched at step 1; the Morph runs when slide 1 is entered.
    const { s, host, canvas, log } = await show(fd)
    await run(1_500)
    s.next(); await run(800)
    s.next(); await run(300)
    expect(s.state.slide).toBe(1)
    expect(log).not.toContain('diag:morph-late')
    expect(ws.used()).toBeGreaterThan(0)
    const before = [...canvas.getContext('2d')!.getImageData(0, 0, canvas.width, canvas.height).data]
    await fd.end(new StreamSessionError('revoked'))
    await s.ended
    expect(s.state.exited).toBe(true)
    expect(frame).toBeUndefined()
    expect(host.children.length).toBe(0)
    // Every bundle, pair, decoded bitmap, retained message and stage frame has been given back.
    expect(ws.used()).toBe(0)
    await run(200)
    expect([...canvas.getContext('2d')!.getImageData(0, 0, canvas.width, canvas.height).data]).toEqual(before)
  })

  it('a bucket change re-fetches the bundle at scale 2; the slide on screen keeps drawing the bundle it holds', async () => {
    const fd = animDoc(['fly'])
    const { s, canvas, log } = await show(fd)
    await run(300)
    const mid = () => px(canvas, canvas.width >> 1, canvas.height >> 1)
    const before = mid()
    expect(before[0]).toBeGreaterThan(200)
    Object.defineProperty(window, 'innerWidth', { configurable: true, value: 1920 })
    Object.defineProperty(window, 'innerHeight', { configurable: true, value: 1080 })
    window.dispatchEvent(new Event('resize'))
    await run(600)
    // The new bucket's bundle replaced the cached one; the visit's pinned bundle is still drawn (never blank).
    expect(fd.log.filter((l) => l.startsWith('motion:'))).toEqual(['motion:full@1', 'motion:full@2'])
    expect(canvas.width).toBe(1920)
    expect(mid()).toEqual(before)
    expect(log).not.toContain('diag:plan-stage-pressure')
    s.exit()
  })

  it('a still slide of an animated deck redraws its new bucket\'s bitmap whole after a bucket change', async () => {
    const fd = animDoc(['still', 'fly'])
    const { s, canvas } = await show(fd)
    await run(300)
    // The still fills the letterboxed slide: its bottom-right corner is the still's colour.
    const corner = () => px(canvas, Math.round(canvas.width * 0.6), Math.round(canvas.height * 0.9))
    expect(corner()).toEqual([0x12, 0x34, 0x56])
    Object.defineProperty(window, 'innerWidth', { configurable: true, value: 1920 })
    Object.defineProperty(window, 'innerHeight', { configurable: true, value: 1080 })
    window.dispatchEvent(new Event('resize'))
    await run(600)
    expect(sent(fd, 'paint:')).toEqual(['paint:0', 'paint:0'])
    expect(corner()).toEqual([0x12, 0x34, 0x56])
    s.exit()
  })

  it('a failed bundle retried on a later visit plays once it lands', async () => {
    const fd = animDoc(['fly', 'fly'])
    fd.rules.set('plan:0', 'unit-failed')
    const { s, root } = await show(fd)
    await run(100)
    expect(frameKind(root)).toBe('error')
    fd.rules.set('plan:0', 'hold')
    s.next()
    await run(2_500)
    s.previous()
    await run(100)
    expect(sent(fd, 'plan:').filter((l) => l === 'plan:0')).toHaveLength(2)
    // Until it lands the slide shows loading (it waits for its bundle), not the old error frame.
    expect(frameKind(root)).toBe('loading')
    fd.release('plan:0')
    await run(300)
    expect(frameKind(root)).toBe('image')
    // It plays from its plan: entered by Previous at its last checkpoint, it steps back to build 0.
    expect(s.state.slide).toBe(0)
    expect(s.state.build).toBe(1)
    s.previous()
    await run(100)
    expect(s.state.slide).toBe(0)
    expect(s.state.build).toBe(0)
    s.exit()
  })

  it('a failed bundle shows the typed error frame; next() still advances', async () => {
    const fd = animDoc(['fly', 'fly', 'fly'])
    // A bundle without its entry frame: decodeBundle refuses it (policy-refused).
    fd.rules.set('plan:0', () => bundleOf(FLY, 0).filter((_, i) => i !== 1))
    const { s, root } = await show(fd)
    await run(100)
    expect(frameKind(root)).toBe('error')
    expect(root.querySelector('[data-officeview-show-frame-text]')?.textContent).toBe('Slide 1 could not be shown (policy-refused)')
    s.next()
    expect(s.state.slide).toBe(1)
    await run(1_200)
    expect(frameKind(root)).toBe('image')
    s.exit()
  })

  it('a refused slide plays its piece-3 still through the unit path', async () => {
    const fd = animDoc(['fly', 'fly'])
    fd.rules.set('plan:0', 'still')
    const { s, root, log } = await show(fd)
    await run(600)
    expect(sent(fd, 'paint:')).toEqual(['paint:0'])
    expect(frameKind(root)).toBe('image')
    expect(log).toContain('diag:still')
    // A still slide has no step: Next leaves it.
    s.next()
    await run(100)
    expect(s.state.slide).toBe(1)
    s.exit()
  })

  it('links withheld: click on a URL region does not advance; diag link-withheld', async () => {
    const withLink = { messages: (u: number) => bundleOf(FLY, u, (p) => ({ ...p, linkRegions: [{ rect: { x: 0, y: 0, w: 960, h: 720 }, action: { kind: 'withheld' as const } }] })) }
    const fd = animDoc([withLink, 'fly'])
    const { s, root, log } = await show(fd)
    await run(1_200)
    root.dispatchEvent(new MouseEvent('click', { clientX: 480, clientY: 270, bubbles: true }))
    await run(100)
    expect(s.state.slide).toBe(0)
    expect(s.state.build).toBe(0)
    expect(log).toContain('diag:link-withheld')
    // Next from the keyboard still steps.
    s.next()
    await run(100)
    expect(s.state.build).toBe(1)
    s.exit()
  })

  it('an internal link region jumps to its streamed slide', async () => {
    const toSlide = { messages: (u: number) => bundleOf(FLY, u, (p) => ({ ...p, linkRegions: [{ rect: { x: 0, y: 0, w: 960, h: 720 }, action: { kind: 'slide' as const, index: 2 } }] })) }
    const fd = animDoc([toSlide, 'fly', 'fly'])
    const { s, root } = await show(fd)
    await run(1_200)
    root.dispatchEvent(new MouseEvent('click', { clientX: 480, clientY: 270, bubbles: true }))
    await run(100)
    expect(s.state.slide).toBe(2)
    s.exit()
  })

  it('reduced motion sends motion:\'reduced\'', async () => {
    const fd = animDoc(['fly', 'morphTo'])
    const { s } = await show(fd, { motion: 'reduced' })
    await run(1_500)
    s.next(); await run(800)
    s.next(); await run(1_500)
    expect(s.state.slide).toBe(1)
    expect(fd.log.filter((l) => l.startsWith('motion:'))).toEqual(['motion:reduced@1', 'motion:reduced@1'])
    // A Morph under reduced motion is a fade: no pair is asked for (nothing charged).
    expect(fd.session.showMorph).not.toHaveBeenCalled()
    s.exit()
  })

  it('no package module is imported', async () => {
    for (const m of THROW_ON_IMPORT) await expect(import(/* @vite-ignore */ m), m).rejects.toThrow()
    for (const m of THROW_ON_CALL) {
      const mod = await import(/* @vite-ignore */ m) as Record<string, () => unknown>
      expect(() => mod.anything!(), m).toThrow(/package code called/)
    }
    // And a whole animated show (bundle, layers, transition, Morph) runs with them in place.
    const fd = animDoc(['fly', 'morphTo'])
    const { s, log } = await show(fd)
    await run(1_500)
    s.next(); await run(800)
    s.next(); await run(1_500)
    expect(s.state.slide).toBe(1)
    expect(log.filter((l) => /T5/.test(l))).toEqual([])
    s.exit()
  })
})

describe('Morph and track replies fail closed (Task 11)', () => {
  const ws = () => createSlideShowWorkspace('desktop', { limit: 512 * 1024 * 1024, scratch: 0 })
  const reply = (messages: Uint8Array[]) => { const r = { messages, release: vi.fn() }; return r }
  const PAIR_OK = () => pairOf(0, 1, 0)

  it('a layered pair decodes every side asset once, for the target unit', async () => {
    const w = ws()
    const r = reply(PAIR_OK())
    const m = await decodeMorph({ from: 0, to: 1, fromStep: 0 }, r, w, bitmapOf)
    expect(r.release).toHaveBeenCalledOnce()
    const ids = new Set(m.record.items.flatMap((it) => [it.from?.assetId, it.to?.assetId]).filter((x): x is number => x !== undefined))
    for (const id of ids) expect(m.assets.bitmap(id)).toBeDefined()
    expect(w.used()).toBeGreaterThan(0)
    m.release()
    expect(w.used()).toBe(0)
    for (const id of ids) expect(m.assets.bitmap(id)).toBeUndefined()
  })

  it.each([
    ['a missing asset', () => PAIR_OK().slice(0, -1)],
    ['a duplicated asset', () => { const p = PAIR_OK(); return [...p.slice(0, -1), p[1]!] }],
    ['an asset no side names', () => { const p = PAIR_OK(); const a = PAIR[1] as Extract<ShowMessage, { tag: 2 | 3 | 4 }>; return [...p.slice(0, -1), encodeShowMessage({ ...a, unit: 1, assetId: 999 })] }],
    ['an asset for another unit', () => { const p = PAIR_OK(); const a = PAIR[1] as Extract<ShowMessage, { tag: 2 | 3 | 4 }>; return [p[0]!, encodeShowMessage({ ...a, unit: 0 }), ...p.slice(2)] }],
    ['a keyframe in a layered pair', () => { const p = PAIR_OK(); return [...p.slice(0, -1), trackOf(1, 0)[0]!] }],
    ['the record of another pair', () => pairOf(0, 1, 1)],
  ])('%s is refused, every lease released', async (_, make) => {
    const w = ws()
    const r = reply(make())
    await expect(decodeMorph({ from: 0, to: 1, fromStep: 0 }, r, w, bitmapOf)).rejects.toMatchObject({ code: 'policy-refused' })
    expect(r.release).toHaveBeenCalled()
    expect(w.used()).toBe(0)
  })

  it('a composited pair decodes its frames through the 8-frame window; its encoded frames stay held until release', async () => {
    const w = ws()
    const rec = (PAIR[0] as Extract<ShowMessage, { tag: 5 }>).morph
    const frames = Array.from({ length: 10 }, (_, i) => encodeShowMessage({ tag: 4, unit: 1, assetId: i, w: FRAME_W, h: FRAME_H, png: png(COLORS[i % 3]!) }))
    const r = reply([encodeShowMessage({ tag: 5, morph: { ...rec, items: [], over: [], pairs: [], composited: { fps: 15, count: 10 } } }), ...frames])
    const m = await decodeMorph({ from: 0, to: 1, fromStep: 0 }, r, w, bitmapOf)
    expect(m.track?.count).toBe(10)
    expect(r.release).not.toHaveBeenCalled()
    m.track!.want(0)
    await flush()
    expect(m.track!.frame(0)).toBeDefined()
    expect(m.track!.frame(8)).toBeUndefined()
    m.release()
    expect(r.release).toHaveBeenCalledOnce()
    expect(w.used()).toBe(0)
  })

  it.each([
    ['too few frames', () => trackOf(0, 0).slice(0, 2)],
    ['a frame of another size', () => [...trackOf(0, 0).slice(0, 2), encodeShowMessage({ tag: 4, unit: 0, assetId: 2, w: 10, h: 10, png: png('#fff', 10, 10) })]],
    ['a layer asset in a track', () => [...trackOf(0, 0).slice(0, 2), bundleOf(FLY, 0)[2]!]],
    ['a frame for another unit', () => trackOf(1, 0)],
  ])('a track with %s is refused and its reply released', (_, make) => {
    const w = ws()
    const r = reply(make())
    expect(() => decodeTrack(0, 3, r, w, bitmapOf)).toThrow(StreamSessionError)
    expect(r.release).toHaveBeenCalledOnce()
    expect(w.used()).toBe(0)
  })
})

describe('fix round 1: pause freezes the show clock', () => {
  const fadeIn = (o: { durationMs: number; advTmMs?: number }) => (p: Plan): Plan => ({ ...p, show: { ...p.show, transition: { element: 'fade', ns: 'p', attrs: {}, advClick: true, ...o } as unknown as NonNullable<Plan['show']['transition']> } })
  const shot = (c: HTMLCanvasElement) => c.toDataURL()

  it('a build paused mid-flight keeps its pixels until resumed, then finishes', async () => {
    const fd = animDoc(['fly'])
    const { s, canvas } = await show(fd)
    await run(300)
    s.next()
    await run(80)
    s.dispatch({ kind: 'pause' })
    expect(s.state.paused).toBe(true)
    const before = shot(canvas)
    await run(800)
    expect(shot(canvas)).toBe(before)
    s.dispatch({ kind: 'pause' })
    expect(s.state.paused).toBe(false)
    await run(800)
    expect(shot(canvas)).not.toBe(before)
    s.exit()
  })

  it('a transition paused mid-way holds its progress and completes after resume', async () => {
    const fd = animDoc(['fly', { messages: (u) => bundleOf(FLY, u, fadeIn({ durationMs: 600 })) }])
    const { s, canvas } = await show(fd)
    await run(1_500)
    s.goTo(1)
    await run(200)
    s.dispatch({ kind: 'pause' })
    const mid = shot(canvas)
    await run(1_000)
    expect(shot(canvas)).toBe(mid)
    s.dispatch({ kind: 'pause' })
    await run(300)
    const later = shot(canvas)
    expect(later).not.toBe(mid)
    await run(600)
    // Settled: the transition is over.
    const settled = shot(canvas)
    await run(300)
    expect(shot(canvas)).toBe(settled)
    s.exit()
  })

  it('the after-timer auto-advance waits out a pause', async () => {
    const auto = { messages: (u: number) => bundleOf(MORPH_TO, u, fadeIn({ durationMs: 100, advTmMs: 500 })) }
    const fd = animDoc([auto, 'fly'])
    const { s } = await show(fd)
    await run(50)
    s.dispatch({ kind: 'pause' })
    await run(2_000)
    expect(s.state.slide).toBe(0)
    s.dispatch({ kind: 'pause' })
    await run(800)
    expect(s.state.slide).toBe(1)
    s.exit()
  })
})

describe('fix round 1: decoded track frames are admitted against the show budget', () => {
  const RGB: Record<string, number[]> = { '#ff0000': [255, 0, 0], '#00ff00': [0, 255, 0], '#0000ff': [0, 0, 255] }
  /** One track whose frame i is COLORS[i]. */
  const rainbow = (unit: number) => [0, 1, 2].map((i) => encodeShowMessage({ tag: 4, unit, assetId: i, w: FRAME_W, h: FRAME_H, png: png(COLORS[i]!) }))
  const F = FRAME_W * FRAME_H * 4
  const bundleBytes = (msgs: Uint8Array[], unit: number) => msgs.reduce((n, m, i) => {
    if (i === 0) return n + m.length
    const a = decodeShowMessage(m, { unit, maxDecoded: SHOW_PLAN_LIMITS.assetDecoded }) as Extract<ShowMessage, { w: number }>
    return n + a.w * a.h * 4
  }, 0)
  const encoded = (unit: number) => rainbow(unit).reduce((n, m) => n + m.length, 0)
  const one = { messages: (u: number) => bundleOf(FLY, u, compositedFly(1)) }
  const mid = (c: HTMLCanvasElement) => px(c, c.width >> 1, c.height >> 1)

  it('the review probe: a window of large frames never takes the show past its budget', async () => {
    const ws = createSlideShowWorkspace('desktop', { limit: 512 * 1024 * 1024, scratch: 0 })
    vi.spyOn(remoteShowDeps, 'workspace').mockReturnValue(ws)
    let counted = 0, peak = 0
    const lease = ws.lease.bind(ws)
    vi.spyOn(ws, 'lease').mockImplementation((n, tag) => {
      const l = lease(n, tag)
      if (l && (tag === 'plan:track-frame' || tag === 'remote-show:plan' || tag.startsWith('remote-show:asset:'))) {
        counted += n; peak = Math.max(peak, counted)
        const rel = l.release.bind(l)
        let done = false
        l.release = () => { if (!done) { done = true; counted -= n } rel() }
      }
      return l
    })
    const data = png('#f00', 3344, 3344)
    const fd = animDoc([one], { tracks: (u) => [0, 1, 2].map((i) => encodeShowMessage({ tag: 4, unit: u, assetId: i, w: 3344, h: 3344, png: data })) })
    vi.spyOn(remoteShowDeps, 'decode').mockImplementation(async (bytes) => {
      const v = new DataView(bytes.buffer, bytes.byteOffset, bytes.byteLength)
      return v.getUint32(16) === 3344 ? { width: 3344, height: 3344, close() {} } : bitmapOf(bytes)
    })
    const { s } = await show(fd)
    await run(100)
    s.next(); frame?.(); await flush()
    s.exit()
    expect(ws.used()).toBe(0)
    expect(peak).toBeLessThanOrEqual(SHOW_PLAN_LIMITS.showDecoded)
  })

  it('a frame with no room holds the previous picture (D23)', async () => {
    const B = bundleBytes(bundleOf(FLY, 0, compositedFly(1)), 0)
    vi.spyOn(remoteShowDeps, 'showBytes', 'get').mockReturnValue(B + encoded(0) + F - 1)
    const fd = animDoc([one], { tracks: rainbow })
    const { s, canvas } = await show(fd)
    await run(300)
    const before = mid(canvas)
    s.next()
    await run(600)
    // No keyframe fits: the step holds what was on screen.
    expect(mid(canvas)).toEqual(before)
    s.exit()
  })

  it('room for one frame at a time: the window rolls and the step ends on its last keyframe', async () => {
    const B = bundleBytes(bundleOf(FLY, 0, compositedFly(1)), 0)
    vi.spyOn(remoteShowDeps, 'showBytes', 'get').mockReturnValue(B + encoded(0) + F)
    const fd = animDoc([one], { tracks: rainbow })
    const { s, canvas } = await show(fd)
    await run(300)
    s.next()
    await run(600)
    expect(mid(canvas)).toEqual(RGB['#0000ff'])
    s.exit()
  })

  it('a pinned bundle the fetcher replaced (a bucket change) still counts until it is released', async () => {
    const B = bundleBytes(bundleOf(FLY, 0, compositedFly(1)), 0)
    // Room for both copies of the bundle, and for a frame beside one of them only: counting the pinned old copy,
    // no frame fits.
    vi.spyOn(remoteShowDeps, 'showBytes', 'get').mockReturnValue(2 * B + encoded(0) + F - 1)
    const fd = animDoc([one], { tracks: rainbow })
    const { s, canvas } = await show(fd)
    await run(300)
    Object.defineProperty(window, 'innerWidth', { configurable: true, value: 1920 })
    Object.defineProperty(window, 'innerHeight', { configurable: true, value: 1080 })
    window.dispatchEvent(new Event('resize'))
    await run(600)
    expect(fd.log.filter((l) => l.startsWith('motion:'))).toEqual(['motion:full@1', 'motion:full@2'])
    const before = mid(canvas)
    s.next()
    await run(600)
    expect(mid(canvas)).toEqual(before)
    s.exit()
  })

  it('a neighbour\'s bundle is evicted before a current frame is refused', async () => {
    const ws = createSlideShowWorkspace('desktop', { limit: 512 * 1024 * 1024, scratch: 0 })
    vi.spyOn(remoteShowDeps, 'workspace').mockReturnValue(ws)
    const B0 = bundleBytes(bundleOf(FLY, 0, compositedFly(1)), 0)
    const B1 = bundleBytes(bundleOf(FLY, 1), 1)
    // Beside both bundles not even one frame fits; without the neighbour every frame does, one at a time.
    vi.spyOn(remoteShowDeps, 'showBytes', 'get').mockReturnValue(B0 + B1 + encoded(0) + F - 1)
    const fd = animDoc([one, 'fly'], { tracks: rainbow })
    const { s, canvas } = await show(fd)
    await run(1_500)
    expect(sent(fd, 'plan:')).toEqual(['plan:0', 'plan:1'])
    const assets = () => ws.used()
    const withBoth = assets()
    s.next()
    await run(600)
    expect(mid(canvas)).toEqual(RGB['#0000ff'])
    // The neighbour's decoded bundle went to make room.
    expect(assets()).toBeLessThan(withBoth)
    s.exit()
  })
})

describe('final review finding 2: an in-progress decode has a cancellation owner', () => {
  /** A decode that holds every call until `go()`; each bitmap it yields records whether it was closed. */
  function heldDecode() {
    const queue: Array<() => void> = []
    const made: Array<{ png: Uint8Array; closed: boolean }> = []
    const decode = (bytes: Uint8Array) => new Promise<HTMLCanvasElement>((res, rej) => {
      queue.push(() => { bitmapOf(bytes).then((cv) => { const tag = { png: bytes, closed: false }; made.push(tag); res(Object.assign(cv, { close: () => { tag.closed = true } })) }, rej) })
    })
    return {
      decode, made,
      pending: () => queue.length,
      async go() { for (const f of queue.splice(0)) f(); await flush() },
    }
  }
  const ws = () => createSlideShowWorkspace('desktop', { limit: 512 * 1024 * 1024, scratch: 0 })
  /** The recorded pair's asset PNGs: the bitmaps a Morph decode makes. */
  const pairPngs = PAIR.filter((m): m is Extract<ShowMessage, { tag: 2 | 3 | 4 }> => m.tag !== 5).map((m) => Buffer.from(m.png).toString('hex'))
  const isPair = (b: { png: Uint8Array }) => pairPngs.includes(Buffer.from(b.png).toString('hex'))

  it('revocation while a Morph bitmap decode is pending: every lease released at ended; the late bitmap is closed', async () => {
    const w = ws()
    vi.spyOn(remoteShowDeps, 'workspace').mockImplementation(() => w)
    const fd = animDoc(['fly', 'morphTo'])
    const { s } = await show(fd)
    await run(1_500)
    const d = heldDecode()
    vi.mocked(remoteShowDeps.decode).mockImplementation(d.decode)
    s.goTo(1); await flush()
    expect(sent(fd, 'morph:')).toContain('morph:0>1@0')
    expect(d.pending()).toBeGreaterThan(0)
    await fd.end(new StreamSessionError('revoked'))
    await s.ended
    expect(w.used()).toBe(0)
    await d.go()
    expect(d.made.length).toBeGreaterThan(0)
    expect(d.made.every((b) => b.closed)).toBe(true)
    expect(w.used()).toBe(0)
  })

  it('navigating away while a Morph bitmap decode is pending releases the pair at once; the late bitmap is closed, never charged', async () => {
    const w = ws()
    vi.spyOn(remoteShowDeps, 'workspace').mockImplementation(() => w)
    const fd = animDoc(['fly', 'morphTo', 'fly'])
    const { s } = await show(fd)
    await run(1_500)
    const d = heldDecode()
    vi.mocked(remoteShowDeps.decode).mockImplementation(d.decode)
    s.goTo(1); await flush()
    expect(sent(fd, 'morph:')).toContain('morph:0>1@0')
    expect(d.pending()).toBeGreaterThan(0)
    // Back on slide 0 (its successor's pair is asked again, at its last step, and decodes at once).
    vi.mocked(remoteShowDeps.decode).mockImplementation(bitmapOf)
    s.goTo(0); await flush()
    expect(sent(fd, 'morph:').at(-1)).toBe('morph:0>1@1')
    const after = w.used()
    await d.go()
    // The abandoned pair's one pending decode lands late: closed, and nothing else of it was decoded or charged.
    expect(d.made.map((b) => [isPair(b), b.closed])).toEqual([[true, true]])
    expect(d.pending()).toBe(0)
    expect(w.used()).toBe(after)
    s.exit()
    await s.ended
    expect(w.used()).toBe(0)
  })

  it('revocation while a bundle bitmap decode is pending: every lease released at ended; the late bitmap is closed', async () => {
    const w = ws()
    vi.spyOn(remoteShowDeps, 'workspace').mockImplementation(() => w)
    const fd = animDoc(['fly', 'fly', 'fly', 'fly', 'fly'])
    const { s } = await show(fd)
    await run(1_500)
    const d = heldDecode()
    vi.mocked(remoteShowDeps.decode).mockImplementation(d.decode)
    s.goTo(4); await flush()
    expect(sent(fd, 'plan:')).toContain('plan:4')
    expect(d.pending()).toBeGreaterThan(0)
    await fd.end(new StreamSessionError('revoked'))
    await s.ended
    expect(w.used()).toBe(0)
    await d.go()
    expect(d.made.every((b) => b.closed)).toBe(true)
    expect(w.used()).toBe(0)
  })

  it('re-review: an abort after the image fulfils but before the decoder resumes closes that bitmap and leases nothing more', async () => {
    for (const which of ['bundle', 'layered'] as const) {
      const w = ws(), ac = new AbortController()
      let deliver: ((b: unknown) => void) | undefined
      const decode = vi.fn(() => new Promise<never>((res) => { deliver = res as (b: unknown) => void }))
      const messages = which === 'bundle' ? bundleOf(FLY, 0) : pairOf(0, 1, 0)
      const release = vi.fn()
      const p = which === 'bundle'
        ? decodeBundle(0, { messages, release }, w, decode, ac.signal)
        : decodeMorph({ from: 0, to: 1, fromStep: 0 }, { messages, release }, w, decode, ac.signal)
      const outcome = p.then((x) => { x.release(); return 'resolved' }, (e: unknown) => (e as StreamSessionError).code)
      await flush()
      const a = decodeShowMessage(messages[1]!, { unit: which === 'bundle' ? 0 : 1, maxDecoded: SHOW_PLAN_LIMITS.assetDecoded }) as Extract<ShowMessage, { tag: 2 | 3 | 4 }>
      const close = vi.fn()
      // The image fulfils; the abort runs after the owner's fulfilment handler, before the decoder's continuation.
      deliver!({ width: a.w, height: a.h, close })
      queueMicrotask(() => ac.abort(new StreamSessionError('closed')))
      expect({ which, result: await outcome, used: w.used(), closed: close.mock.calls.length, decodes: decode.mock.calls.length })
        .toEqual({ which, result: 'closed', used: 0, closed: 1, decodes: 1 })
      expect(release).toHaveBeenCalled()
    }
  })

  it('decodeBundle and decodeMorph: an abort mid-decode rejects at once, releases the reply and every lease, closes the late bitmap', async () => {
    for (const which of ['bundle', 'layered', 'composited'] as const) {
      const w = ws(), d = heldDecode(), ac = new AbortController()
      const rec = (PAIR[0] as Extract<ShowMessage, { tag: 5 }>).morph
      const messages = which === 'bundle' ? bundleOf(FLY, 0)
        : which === 'layered' ? pairOf(0, 1, 0)
          : [encodeShowMessage({ tag: 5, morph: { ...rec, items: [], over: [], pairs: [], composited: { fps: 15, count: 3 } } }), ...trackOf(1, 0)]
      const release = vi.fn()
      const p = which === 'bundle'
        ? decodeBundle(0, { messages, release }, w, d.decode, ac.signal)
        : decodeMorph({ from: 0, to: 1, fromStep: 0 }, { messages, release }, w, d.decode, ac.signal)
      const settled = p.then(() => 'resolved', (e: unknown) => (e as StreamSessionError).code ?? String(e))
      await flush()
      if (which === 'composited') {
        // A composited pair awaits nothing: it is handed back, and its owner releases it.
        expect(await settled).toBe('resolved')
        ;(await p).release()
        expect(w.used()).toBe(0)
        continue
      }
      expect(d.pending()).toBe(1)
      expect(w.used()).toBeGreaterThan(0)
      ac.abort(new StreamSessionError('closed'))
      expect({ which, used: w.used() }).toEqual({ which, used: 0 })
      expect(release).toHaveBeenCalled()
      expect(await settled).toBe('closed')
      await d.go()
      expect({ which, closed: d.made.map((b) => b.closed) }).toEqual({ which, closed: [true] })
      expect(d.pending()).toBe(0)
      expect(w.used()).toBe(0)
    }
  })
})
