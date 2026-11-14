/**
 * Piece 4 Task 8: `PlanStage` plays a show plan (OVSP records, layer bitmaps,
 * keyframe tracks and Morph records) on the compositor, the behaviours and
 * `morphFrame`, with no package. Records are hand-built in the `flyInDeck`
 * shape (its real timing tree, read from the fixture) and round-tripped
 * through the OVSP codec; bitmaps are solid canvases at a 0.1 bucket. The
 * output is a 96 × 72 Canvas 2D surface, so one output px is 10 slide px.
 */
import { afterEach, beforeAll, describe, expect, it, vi } from 'vitest'
import { createDiagnosticSink } from '../../src/slideshow/diagnostics'
import { createSlideTimeline, type SlideCursor, type SlideTimeline } from '../../src/slideshow/engine/steps'
import type { StageModel } from '../../src/slideshow/player/stage'
import type { BehaviourNode, CTn, ContainerNode, Lease, ShowSlide, SlideShowWorkspace, TimingTree, TransitionSpec, Compositor, LayerHandle } from '../../src/slideshow/types'
import { decodeShowMessage, encodeShowMessage, SHOW_PLAN_LIMITS, type MorphPlanRecord, type ShowPlanRecord } from '../../src/stream/show-plan'
import type { ImageBitmapLike } from '../../src/stream/remote/show-fetch'
import { flyInDeck, loadShowDeck } from '../../src/testdata/show-decks'

// A WebGL2 stand-in, off unless a test turns it on (jsdom has no WebGL2).
const gl = vi.hoisted(() => ({ on: false, made: [] as Array<{ uploads: number; lost: boolean }> }))
vi.mock('../../src/slideshow/render/webgl2', async (orig) => {
  const real = await orig<typeof import('../../src/slideshow/render/webgl2')>()
  return {
    ...real,
    createWebGL2Compositor: (_canvas: HTMLCanvasElement, _ws: SlideShowWorkspace) => {
      if (!gl.on) return undefined
      const rec = { uploads: 0, lost: false }
      gl.made.push(rec)
      const tex = new Map<number, Lease>()
      let id = 1
      const c: Compositor & { contextLost(): void; contextRestored(): boolean } = {
        backend: 'webgl2' as const, maxTextureSize: 8192,
        upload(layer, lease) {
          if (!lease.grow(layer.canvas.width * layer.canvas.height * 4)) return { refused: 'workspace' }
          lease.shrink(Math.min(layer.bytes, lease.bytes - layer.canvas.width * layer.canvas.height * 4))
          rec.uploads++
          const h = { id: id++ }
          tex.set(h.id, lease)
          return h
        },
        release(h: LayerHandle) { tex.get(h.id)?.release(); tex.delete(h.id) },
        frame() {}, resize() {}, setViewport() { return true }, setSlideIndex() {},
        contextLost() { rec.lost = true; for (const l of tex.values()) l.release(); tex.clear() },
        contextRestored: () => false,
        dispose() { for (const l of tex.values()) l.release(); tex.clear() },
      }
      return c
    },
  }
})

import { createPlanStage, createTrackWindow, MORPH_WAIT_MS, planIterations, type PlanAssets, type PlanSlideData, type PlanStage, type PlanStageDeps, type PlanTrack } from '../../src/stream/remote/plan-stage'

// ---------------------------------------------------------------------------
// Fixtures
// ---------------------------------------------------------------------------

const S = 0.1
const W = 960, H = 720
const OW = 96, OH = 72

type RGB = [number, number, number]
const WHITE: RGB = [255, 255, 255]
const BOX: RGB = [0x33, 0x66, 0x99]
const FOOT: RGB = [0xf4, 0xb1, 0x83]
const SLIDE_B: RGB = [0xdd, 0xeb, 0xf7]
const hex = (c: RGB) => `rgb(${c[0]},${c[1]},${c[2]})`

/** A solid canvas of `w × h` px, optionally with a rect of another colour (in its own px). */
function solid(w: number, h: number, c: RGB, more: Array<{ x: number; y: number; w: number; h: number; c: RGB }> = []): HTMLCanvasElement {
  const cv = document.createElement('canvas')
  cv.width = Math.max(1, Math.ceil(w)); cv.height = Math.max(1, Math.ceil(h))
  const ctx = cv.getContext('2d')!
  ctx.fillStyle = hex(c); ctx.fillRect(0, 0, cv.width, cv.height)
  for (const r of more) { ctx.fillStyle = hex(r.c); ctx.fillRect(r.x, r.y, r.w, r.h) }
  return cv
}

/** The OVSP codec round trip: what the client receives. */
const viaWire = (plan: ShowPlanRecord): ShowPlanRecord =>
  (decodeShowMessage(encodeShowMessage({ tag: 1, plan }), { unit: plan.slide.index, maxDecoded: SHOW_PLAN_LIMITS.showDecoded }) as { plan: ShowPlanRecord }).plan
const morphViaWire = (morph: MorphPlanRecord): MorphPlanRecord =>
  (decodeShowMessage(encodeShowMessage({ tag: 5, morph }), { unit: 0, maxDecoded: SHOW_PLAN_LIMITS.showDecoded }) as { morph: MorphPlanRecord }).morph

let flyTiming: ShowSlide['timing']
beforeAll(async () => {
  const deck = await loadShowDeck(await flyInDeck())
  flyTiming = (await deck.meta(0)).timing
})

const assetsOf = (m: Map<number, ImageBitmapLike>): PlanAssets => ({ bitmap: (id) => m.get(id), bytes: (id) => { const b = m.get(id); return b ? b.width * b.height * 4 : 0 } })

/** The fly-in slide: background (asset 1), the box `sp:3` (asset 2, flies in on click 1), the footer (asset 3); entry frame 0. */
function flyIn(over: Partial<ShowPlanRecord> = {}, index = 0): { plan: ShowPlanRecord; assets: Map<number, ImageBitmapLike> } {
  const assets = new Map<number, ImageBitmapLike>([
    [0, solid(OW, OH, WHITE, [{ x: 8, y: 60, w: 80, h: 6, c: FOOT }])],
    [1, solid(OW, OH, WHITE)],
    [2, solid(20, 20, BOX)],
    [3, solid(80, 6, FOOT)],
  ])
  const plan: ShowPlanRecord = {
    v: 1, slide: { index, w: W, h: H }, options: { showAnimation: true, useTimings: true },
    show: { timing: flyTiming, builds: [] },
    layers: [
      { id: 0, bounds: { x: 0, y: 0, w: W, h: H }, z: 0, assetId: 1, decodedBytes: OW * OH * 4, scale: S, background: true },
      { id: 1, key: 'sp:3', bounds: { x: 380, y: 260, w: 200, h: 200 }, z: 1, assetId: 2, decodedBytes: 20 * 20 * 4, scale: S },
      { id: 2, bounds: { x: 80, y: 600, w: 800, h: 60 }, z: 2, assetId: 3, decodedBytes: 80 * 6 * 4, scale: S },
    ],
    base: [['sp:3', { x: 380, y: 260, w: 200, h: 200, rot: 0 }]],
    steps: [{ trigger: 'click', eligibility: 'layers' }],
    triggers: [], linkRegions: [], iterate: { keys: [], groups: [] }, entry: 0, diagnostics: [],
    ...over,
  }
  return { plan: viaWire(plan), assets }
}

/** A transition-only slide: no layers, its entry frame (asset 0) is solid `SLIDE_B`. */
function plainSlide(index: number, transition?: TransitionSpec): { plan: ShowPlanRecord; assets: Map<number, ImageBitmapLike> } {
  const plan: ShowPlanRecord = {
    v: 1, slide: { index, w: W, h: H }, options: { showAnimation: true, useTimings: true },
    show: { timing: undefined, builds: [], ...(transition ? { transition } : {}) },
    layers: [], base: [], steps: [], triggers: [], linkRegions: [], iterate: { keys: [], groups: [] }, entry: 0, diagnostics: [],
  }
  return { plan: viaWire(plan), assets: new Map([[0, solid(OW, OH, SLIDE_B)]]) }
}

const fade = (durationMs = 500): TransitionSpec => ({ element: 'fade', ns: 'p', attrs: {}, durationMs, advClick: true })
const morphSpec = (durationMs = 1200): TransitionSpec => ({ element: 'morph', ns: 'p159', attrs: { option: 'byObject' }, durationMs, advClick: true })

/** A track of `count` full-slide frames; frame `i` is the colour `colours(i)`; `missing` frames are not decoded. */
function fakeTrack(count: number, colours: (i: number) => RGB, missing: (i: number) => boolean = () => false): PlanTrack & { wanted: number[] } {
  const frames = Array.from({ length: count }, (_, i) => solid(OW, OH, colours(i)))
  const wanted: number[] = []
  return { fps: 15, count, frame: (i) => (missing(i) ? undefined : frames[i]), want: (i) => { wanted.push(i) }, wanted }
}

// ---------------------------------------------------------------------------
// Rig
// ---------------------------------------------------------------------------

interface Pool { used: number; limit: number }
function workspaceOf(limit = Number.POSITIVE_INFINITY): SlideShowWorkspace & { pool: Pool } {
  const pool: Pool = { used: 0, limit }
  const lease = (bytes: number, tag: string): Lease | undefined => {
    if (bytes < 0 || pool.used + bytes > pool.limit) return undefined
    pool.used += bytes
    let held = bytes, live = true
    const l: Lease = {
      get tag() { return tag }, get bytes() { return live ? held : 0 },
      grow(extra) { if (!live || pool.used + extra > pool.limit) return false; pool.used += extra; held += extra; return true },
      shrink(less) { const n = Math.min(less, held); held -= n; pool.used -= n },
      transfer() { return l },
      release() { if (!live) return; live = false; pool.used -= held },
    }
    return l
  }
  return { pool, limit, lease, used: () => pool.used, reserveTerminal: () => undefined }
}

interface Rig {
  stage: PlanStage
  output: HTMLCanvasElement
  root: HTMLElement
  sink: ReturnType<typeof createDiagnosticSink>
  ws: ReturnType<typeof workspaceOf>
  now: { t: number }
  slides: Map<number, PlanSlideData>
  tracks: Map<string, PlanTrack | 'pending'>
  timelines: Map<number, SlideTimeline>
  model(index: number, cursor?: SlideCursor): StageModel
  /** Enters slide `index` at show time `now.t`. */
  enter(index: number, cursor?: SlideCursor, neighbours?: number[]): Promise<void>
  /** Leaves the shown slide, as the driver does. */
  leave(): void
  pixel(x: number, y: number): RGB
}

const rigs: Rig[] = []
afterEach(() => {
  for (const r of rigs.splice(0)) { r.stage.dispose(); r.root.remove() }
  gl.on = false
  gl.made.length = 0
})

function rig(opts: { slides?: Array<{ plan: ShowPlanRecord; assets: Map<number, ImageBitmapLike> }>; ws?: ReturnType<typeof workspaceOf>; reduced?: boolean; morph?: PlanStageDeps['morph']; force?: 'canvas2d' | 'none'; stills?: Map<number, ImageBitmapLike> } = {}): Rig {
  const root = document.createElement('div')
  document.body.appendChild(root)
  root.getBoundingClientRect = () => ({ left: 0, top: 0, width: OW, height: OH, right: OW, bottom: OH, x: 0, y: 0, toJSON: () => ({}) })
  const output = document.createElement('canvas')
  output.width = OW; output.height = OH
  root.appendChild(output)
  const sink = createDiagnosticSink()
  const ws = opts.ws ?? workspaceOf()
  const now = { t: 0 }
  const slides = new Map<number, PlanSlideData>()
  const tracks = new Map<string, PlanTrack | 'pending'>()
  for (const s of opts.slides ?? []) {
    const index = s.plan.slide.index
    slides.set(index, { plan: s.plan, assets: assetsOf(s.assets), track: (step) => tracks.get(`${index}:${step}`) })
  }
  const stage = createPlanStage({
    root, output, workspace: ws, sink, reduced: opts.reduced ?? false,
    showNow: () => now.t,
    data: (i) => slides.get(i),
    still: (i) => opts.stills?.get(i),
    morph: opts.morph ?? (async () => undefined),
    ...(opts.force === 'none' ? {} : { force: 'canvas2d' as const }),
  })
  const timelines = new Map<number, SlideTimeline>()
  const showSlide = (i: number): ShowSlide => {
    const plan = slides.get(i)?.plan
    return { index: i, hidden: false, timing: plan?.show.timing, ...(plan?.show.transition ? { transition: plan.show.transition } : {}), builds: plan?.show.builds ?? [], notes: [], links: [], media: [] }
  }
  const metas = new Map<number, ShowSlide>()
  const r: Rig = {
    stage, output, root, sink, ws, now, slides, tracks, timelines,
    model(index, cursor) {
      let meta = metas.get(index)
      if (!meta) metas.set(index, meta = showSlide(index))
      const text = stage.textTiming?.(index, meta, () => true)
      let tl = timelines.get(index)
      if (!tl) timelines.set(index, tl = createSlideTimeline(text?.tree ?? meta.timing, { showAnimation: true, reducedMotion: opts.reduced ?? false, ...(text ? { orderOf: text.orderOf } : {}) }))
      return { index, slide: meta, timeline: tl, cursor: cursor ?? tl.initial() }
    },
    enter(index, cursor, neighbours = []) { return stage.enter(r.model(index, cursor), now.t, neighbours) },
    leave() { stage.flattenCurrent(); stage.leave() },
    pixel(x, y) { const d = output.getContext('2d')!.getImageData(x, y, 1, 1).data; return [d[0]!, d[1]!, d[2]!] },
  }
  rigs.push(r)
  return r
}

const near = (a: RGB, b: RGB, tol = 3) => a.every((v, i) => Math.abs(v - b[i]!) <= tol)
const expectPx = (got: RGB, want: RGB, tol = 3) => expect({ got, near: near(got, want, tol) }).toEqual({ got, near: true })
const tick = () => new Promise((res) => setTimeout(res, 0))
const pixels = (c: HTMLCanvasElement) => Array.from(c.getContext('2d')!.getImageData(0, 0, c.width, c.height).data)

// ---------------------------------------------------------------------------
// Tests
// ---------------------------------------------------------------------------

describe('PlanStage', () => {
  it('enter draws the entry frame, then layers at their timeline states', async () => {
    const r = rig({ slides: [plainSlide(0), { ...flyIn({ show: { timing: flyTiming, builds: [], transition: fade(500) } }, 1) }] })
    await r.enter(0)
    r.stage.frame(r.model(0).cursor, 0)
    r.leave()
    r.now.t = 1000
    await r.enter(1)
    const tl = r.timelines.get(1)!
    // The transition's incoming frame is the entry frame (footer, no box).
    r.stage.frame(tl.initial(), 1000 + 499)
    expect(r.stage.inspect().lastDraws).toEqual(['transition'])
    expect(r.stage.inspect().transition).toMatchObject({ effect: 'fade', duration: 500, from: 'flat', to: 'entry' })
    // After it: the layers at the initial state (the box is hidden until its entrance).
    const at = (ms: number) => 1500 + ms
    r.stage.frame(tl.initial(), at(0))
    expect(r.stage.inspect().lastDraws).toEqual(['layer:0', 'layer:2'])
    expectPx(r.pixel(48, 36), WHITE)
    expectPx(r.pixel(48, 63), FOOT)
    // Click 1 at slide time 100: half way the box is left of its rest; settled it is at its rest.
    const c1 = tl.next(tl.initial(), 100)
    r.stage.frame(c1, at(100 + 250))
    expect(r.stage.inspect().lastDraws).toEqual(['layer:0', 'layer:1', 'layer:2'])
    expectPx(r.pixel(48, 36), WHITE)
    r.stage.frame(c1, at(100 + 2000))
    expectPx(r.pixel(48, 36), BOX)
    expectPx(r.pixel(48, 63), FOOT)
  })

  it('transition from-frame comes from flattenCurrent (pixels equal the last drawn frame)', async () => {
    const r = rig({ slides: [flyIn(), plainSlide(1, fade(500))] })
    await r.enter(0)
    const tl = r.timelines.get(0)!
    r.stage.frame(tl.checkpoint(1), 0)
    const last = pixels(r.output)
    expectPx(r.pixel(48, 36), BOX)
    r.leave()
    r.now.t = 10
    await r.enter(1)
    // Progress 0 of the fade: exactly the outgoing frame.
    r.stage.frame(r.timelines.get(1)!.initial(), 10)
    expect(r.stage.inspect().lastDraws).toEqual(['transition'])
    expect(r.stage.inspect().transition).toMatchObject({ from: 'flat', to: 'entry' })
    const got = pixels(r.output)
    let diff = 0
    for (let i = 0; i < got.length; i++) if (Math.abs(got[i]! - last[i]!) > 1) diff++
    expect(diff).toBe(0)
    // The end of the fade is the incoming slide.
    r.stage.frame(r.timelines.get(1)!.initial(), 600)
    expectPx(r.pixel(48, 36), SLIDE_B)
  })

  it('composited step plays its frame at floor(t*15)', async () => {
    const steps: ShowPlanRecord['steps'] = [{ trigger: 'click', eligibility: 'composited', reason: 'colour', track: { fps: 15, count: 8 } }]
    const r = rig({ slides: [flyIn({ steps, diagnostics: [{ code: 'step-composited', step: 0 }] })] })
    const colour = (i: number): RGB => [10 * i, 200, 30]
    const track = fakeTrack(8, colour)
    r.tracks.set('0:0', track)
    await r.enter(0)
    const tl = r.timelines.get(0)!
    const c1 = tl.next(tl.initial(), 0)
    r.stage.frame(c1, 200)
    expect(r.stage.inspect().lastDraws).toEqual(['track:0:3'])
    expectPx(r.pixel(48, 36), colour(3))
    expect(track.wanted).toContain(3)
    r.stage.frame(c1, 466)
    expect(r.stage.inspect().lastDraws).toEqual(['track:0:6'])
    // Past the step's end the track's end frame holds; no layer is drawn for a composited step.
    r.stage.frame(c1, 5000)
    expect(r.stage.inspect().lastDraws).toEqual(['track:0:7'])
    expectPx(r.pixel(48, 36), colour(7))
    // A settled checkpoint after it shows the end frame too.
    r.stage.frame(tl.checkpoint(1), 6000)
    expect(r.stage.inspect().lastDraws).toEqual(['track:0:7'])
  })

  it('track decode window holds ≤ 8 decoded frames; an undecoded frame holds the previous', async () => {
    // The window: 20 frames, decoded asynchronously; at most 8 are held, each charged w·h·4.
    const ws = workspaceOf()
    const decoded = new Set<number>()
    let closed = 0
    const win = createTrackWindow({
      count: 20, frameBytes: OW * OH * 4, workspace: ws,
      decode: async (i) => { decoded.add(i); const c = solid(OW, OH, [i, 0, 0]) as HTMLCanvasElement & { close?(): void }; c.close = () => { closed++; decoded.delete(i) }; return c },
    })
    win.want(0)
    await tick()
    expect([...decoded].sort((a, b) => a - b)).toEqual([0, 1, 2, 3, 4, 5, 6, 7])
    expect(ws.used()).toBe(8 * OW * OH * 4)
    expect(win.frame(3)).toBeDefined()
    expect(win.frame(9)).toBeUndefined()
    win.want(6)
    await tick()
    expect(decoded.size).toBeLessThanOrEqual(8)
    expect([...decoded].sort((a, b) => a - b)).toEqual([6, 7, 8, 9, 10, 11, 12, 13])
    expect(win.frame(2)).toBeUndefined()
    expect(closed).toBe(6)
    expect(ws.used()).toBe(8 * OW * OH * 4)
    win.dispose()
    expect(ws.used()).toBe(0)

    // The stage: frames 4 and up are not decoded yet, so frame 3 holds.
    const steps: ShowPlanRecord['steps'] = [{ trigger: 'click', eligibility: 'composited', track: { fps: 15, count: 8 } }]
    const r = rig({ slides: [flyIn({ steps })] })
    const colour = (i: number): RGB => [20 * i, 100, 200]
    r.tracks.set('0:0', fakeTrack(8, colour, (i) => i >= 4))
    await r.enter(0)
    const tl = r.timelines.get(0)!
    const c1 = tl.next(tl.initial(), 0)
    r.stage.frame(c1, 210)
    expectPx(r.pixel(48, 36), colour(3))
    r.stage.frame(c1, 400)
    expect(r.stage.inspect().lastDraws).toEqual(['track:0:3'])
    expectPx(r.pixel(48, 36), colour(3))
  })

  it('text by word: expandIterations over plan.iterate gives the local fan', async () => {
    const { buildPptxDeck } = await import('../../src/testdata/ooxml-builders')
    const { OfficePackage } = await import('../../src/core/zip')
    const { parsePptx } = await import('../../src/pptx/parse')
    const { textTimingFor } = await import('../../src/slideshow/player/text-builds')
    const bytes = await buildPptxDeck({ slides: [{ shapes: [{ id: 2, off: ['0', '0'], ext: ['4572000', '914400'], paragraphs: [{ runs: [{ text: 'Ab cd' }] }] }] }] })
    const slide = (await parsePptx(await OfficePackage.load(bytes))).slides[0]!
    const c = (id: number, o: Partial<CTn> = {}): CTn => ({ id, durMs: undefined, accel: 0, decel: 0, autoRev: false, start: [], end: [], fill: 'hold', ...o })
    const set = (id: number, to: string): BehaviourNode => ({ kind: 'set', ctn: c(id, { durMs: 1, start: [{ delayMs: 0 }] }), target: { spid: '2' }, attrNames: ['style.visibility'], additive: 'base', accumulate: 'none', to } as BehaviourNode)
    const effect = (id: number, cls: string, unit: 'wd' | 'lt', b: BehaviourNode): ContainerNode =>
      ({ kind: 'par', ctn: c(id, { presetClass: cls, start: [{ delayMs: 0 }], iterate: { type: unit, backwards: false, tmAbsMs: 500 } } as Partial<CTn>), children: [b] })
    const tree: TimingTree = { root: { kind: 'par', ctn: c(1, { durMs: 'indefinite' }), children: [effect(3, 'entr', 'wd', set(4, 'visible')), effect(5, 'exit', 'lt', set(6, 'hidden'))] }, build: [] }
    const tt = textTimingFor(slide, tree, { charge: () => true, sink: createDiagnosticSink() })!
    // What the server sends (build.ts): the items and the word groups, never the fan itself.
    const iterate: ShowPlanRecord['iterate'] = {
      keys: [...tt.items].map(([k, items]) => [k, [...items]]),
      groups: [...(tt.groups ?? [])].map(([k, g]) => [k, { ...(g.wd ? { wd: g.wd.map((x) => [...x]) } : {}), ...(g.lt ? { lt: g.lt.map((x) => [...x]) } : {}) }]),
    }
    const plan = viaWire({ ...flyIn().plan, show: { timing: tree, builds: [] }, iterate })
    const fanOf = (f: ReadonlyMap<BehaviourNode, { items: ReadonlyArray<{ key: string; delayMs: number }> }>) =>
      [...f].map(([n, x]) => [n.ctn.id, x.items.map((i) => [i.key, i.delayMs])])
    const mine = planIterations(plan.show.timing as TimingTree, plan.iterate, { charge: () => true })
    expect(fanOf(mine.fan)).toEqual(fanOf(tt.fan))
    expect(fanOf(mine.fan).length).toBe(2)
    // The stage hands the driver the same working copy.
    const r = rig({ slides: [{ plan, assets: flyIn().assets }] })
    const meta: ShowSlide = { index: 0, hidden: false, timing: plan.show.timing, builds: [], notes: [], links: [], media: [] }
    const text = r.stage.textTiming!(0, meta, () => true)!
    expect(text.tree).not.toBe(tree)
    expect((text.tree.root as ContainerNode).children.length).toBe((tree.root as ContainerNode).children.length)
  })

  it('pending track → cut with step-deferred', async () => {
    const steps: ShowPlanRecord['steps'] = [{ trigger: 'click', eligibility: 'composited', track: { fps: 15, count: 8 } }]
    const r = rig({ slides: [flyIn({ steps })] })
    r.tracks.set('0:0', 'pending')
    await r.enter(0)
    const tl = r.timelines.get(0)!
    r.stage.frame(tl.initial(), 0)
    expectPx(r.pixel(48, 36), WHITE)
    const c1 = tl.next(tl.initial(), 0)
    r.stage.frame(c1, 100)
    // Never the layers of a composited step: the picture holds, diagnosed.
    expect(r.stage.inspect().lastDraws).toEqual(['layer:0', 'layer:2'])
    expectPx(r.pixel(48, 36), WHITE)
    expect(r.sink.list().map((d) => d.code)).toContain('step-deferred')
    // The track arrives mid-step: the step is a cut, so its end frame shows at once.
    const colour = (i: number): RGB => [30 * i, 50, 50]
    r.tracks.set('0:0', fakeTrack(8, colour))
    r.stage.frame(c1, 150)
    expect(r.stage.inspect().lastDraws).toEqual(['track:0:7'])
    expectPx(r.pixel(48, 36), colour(7))
  })

  it('Morph waits ≤ 500 ms for its record or track, then fades with morph-late', async () => {
    let resolve!: (v: Awaited<ReturnType<PlanStageDeps['morph']>>) => void
    const asked: Array<[number, number, number]> = []
    const r = rig({
      slides: [flyIn(), plainSlide(1, morphSpec(1200))],
      morph: (from, to, fromStep) => { asked.push([from, to, fromStep]); return new Promise((res) => { resolve = res }) },
    })
    await r.enter(0)
    const tl0 = r.timelines.get(0)!
    r.stage.frame(tl0.checkpoint(1), 0)
    r.leave()
    r.now.t = 100
    await r.enter(1)
    expect(asked).toEqual([[0, 1, 1]])
    const tl = r.timelines.get(1)!
    // Waiting: the outgoing frame holds, and the slide's timeline has not begun.
    r.now.t = 400
    r.stage.frame(tl.initial(), 400)
    expect(r.stage.inspect().lastDraws).toEqual(['from'])
    expect(r.stage.slideTime(400)).toBe(Number.NEGATIVE_INFINITY)
    // 500 ms after the request: the fade plays from now, diagnosed.
    r.now.t = 100 + MORPH_WAIT_MS
    r.stage.frame(tl.initial(), r.now.t)
    expect(r.sink.list().map((d) => d.code)).toContain('morph-late')
    expect(r.stage.inspect().lastDraws).toEqual(['transition'])
    expect(r.stage.slideTime(r.now.t + 1200)).toBe(0)
    // A record arriving late is unused.
    resolve(undefined)
    await tick()
    r.stage.frame(tl.initial(), r.now.t + 600)
    expect(r.stage.inspect().lastDraws).toEqual(['transition'])
  })

  it('a layered Morph record arriving in time plays morphFrame', async () => {
    const record: MorphPlanRecord = morphViaWire({
      v: 1, from: 0, to: 1, fromStep: 1, over: [], pairs: [{ fromKey: 'sp:3', toKey: 'sp:3', rule: 'name-kind' }],
      items: [{ from: { assetId: 10, geometry: { x: 380, y: 260, w: 200, h: 200, rot: 0 }, bounds: { x: 380, y: 260, w: 200, h: 200 }, key: 'sp:3' }, to: { assetId: 11, geometry: { x: 0, y: 0, w: 200, h: 200, rot: 0 }, bounds: { x: 0, y: 0, w: 200, h: 200 }, key: 'sp:3' } }],
    })
    const assets = assetsOf(new Map([[10, solid(20, 20, BOX)], [11, solid(20, 20, BOX)]]))
    const r = rig({ slides: [flyIn(), plainSlide(1, morphSpec(1000))], morph: async () => ({ record, assets }) })
    await r.enter(0)
    r.stage.frame(r.timelines.get(0)!.checkpoint(1), 0)
    r.leave()
    await r.enter(1)
    await tick()
    const tl = r.timelines.get(1)!
    r.stage.frame(tl.initial(), 0)
    expect(r.stage.inspect().lastDraws).toEqual(['morph'])
    // Half way: the box is between (380, 260) and (0, 0), so at (190, 130)..(390, 330).
    r.stage.frame(tl.initial(), 500)
    expectPx(r.pixel(29, 23), BOX)
    expect(r.stage.slideTime(1000)).toBe(0)
  })

  it('composited Morph plays its frames at floor(t*15) and draws no layers', async () => {
    const record: MorphPlanRecord = morphViaWire({ v: 1, from: 0, to: 1, fromStep: 0, items: [], over: [], pairs: [], composited: { fps: 15, count: 18 } })
    const colour = (i: number): RGB => [10 * i, 10, 240]
    const track = fakeTrack(18, colour)
    const r = rig({ slides: [plainSlide(0), flyIn({ show: { timing: flyTiming, builds: [], transition: morphSpec(1200) } }, 1)], morph: async () => ({ record, assets: assetsOf(new Map()), track }) })
    await r.enter(0)
    r.stage.frame(r.timelines.get(0)!.initial(), 0)
    r.leave()
    await r.enter(1)
    await tick()
    const tl = r.timelines.get(1)!
    r.stage.frame(tl.initial(), 0)
    r.stage.frame(tl.initial(), 300)
    expect(r.stage.inspect().lastDraws).toEqual(['morph-track:4'])
    expectPx(r.pixel(48, 36), colour(4))
    expect(track.wanted).toContain(4)
    r.stage.frame(tl.initial(), 1199)
    expect(r.stage.inspect().lastDraws).toEqual(['morph-track:17'])
    // After the Morph: the incoming slide's own layers.
    r.stage.frame(tl.initial(), 1300)
    expect(r.stage.inspect().lastDraws).toEqual(['layer:0', 'layer:2'])
  })

  it('hitTest moves the trigger rect with the target\'s transform', async () => {
    const r = rig({ slides: [flyIn({ triggers: [{ spid: '3', rects: [{ x: 380, y: 260, w: 200, h: 200 }] }] })] })
    await r.enter(0)
    const tl = r.timelines.get(0)!
    const on = { triggers: ['3'], links: [] }
    // Hidden before its entrance: no hit.
    expect(r.stage.hitTest(48, 36, tl.initial(), 0, on)).toBeUndefined()
    // Settled at its rest.
    expect(r.stage.hitTest(48, 36, tl.checkpoint(1), 0, on)).toEqual({ kind: 'trigger', spid: '3' })
    expect(r.stage.hitTest(10, 36, tl.checkpoint(1), 0, on)).toBeUndefined()
    // Mid-flight: the rect has moved left with the box.
    const c1 = tl.next(tl.initial(), 0)
    r.stage.frame(c1, 250)
    const state = r.stage.frame(c1, 250)!
    const s3 = state.get('sp:3')!
    const dx = s3.transform[4] + (s3.motion?.dx ?? 0)
    expect(dx).toBeLessThan(-100)
    const x = Math.round((480 + dx) / 10)
    expect(r.stage.hitTest(x, 36, c1, 250, on)).toEqual({ kind: 'trigger', spid: '3' })
    expect(r.stage.hitTest(48, 36, c1, 250, on)).toBeUndefined()
    // Not a trigger the timeline knows: passes.
    expect(r.stage.hitTest(48, 36, tl.checkpoint(1), 0, { triggers: [], links: [] })).toBeUndefined()
  })

  it('withheld link region returns {kind:\'withheld\'}', async () => {
    const r = rig({ slides: [flyIn({ linkRegions: [{ rect: { x: 0, y: 0, w: 200, h: 100 }, action: { kind: 'withheld' } }, { rect: { x: 700, y: 0, w: 200, h: 100 }, action: { kind: 'slide', index: 1 } }] })] })
    await r.enter(0)
    const tl = r.timelines.get(0)!
    expect(r.stage.hitTest(5, 5, tl.initial(), 0, { triggers: [], links: [] })).toEqual({ kind: 'withheld' })
    expect(r.stage.hitTest(75, 5, tl.initial(), 0, { triggers: [], links: [] })).toEqual({ kind: 'link', link: { spid: '', action: { kind: 'slide', index: 1 } } })
    expect(r.stage.hitTest(48, 50, tl.initial(), 0, { triggers: [], links: [] })).toBeUndefined()
  })

  it('reduced motion applies reducedTransition', async () => {
    const push: TransitionSpec = { element: 'push', ns: 'p', attrs: { dir: 'u' }, durationMs: 1000, advClick: true }
    const r = rig({ slides: [plainSlide(0), plainSlide(1, push)], reduced: true })
    await r.enter(0)
    r.stage.frame(r.timelines.get(0)!.initial(), 0)
    r.leave()
    await r.enter(1)
    r.stage.frame(r.timelines.get(1)!.initial(), 0)
    expect(r.stage.inspect().transition).toMatchObject({ effect: 'fade', duration: 200 })
    expect(r.stage.slideTime(200)).toBe(0)
  })

  it('reduced motion never asks for a Morph', async () => {
    const morph = vi.fn(async () => undefined)
    const r = rig({ slides: [plainSlide(0), plainSlide(1, morphSpec(1200))], reduced: true, morph })
    await r.enter(0)
    r.stage.frame(r.timelines.get(0)!.initial(), 0)
    r.leave()
    await r.enter(1)
    r.stage.frame(r.timelines.get(1)!.initial(), 0)
    expect(morph).not.toHaveBeenCalled()
    expect(r.stage.inspect().transition).toMatchObject({ effect: 'fade', duration: 200 })
  })

  it('lost WebGL2 → Canvas 2D and re-upload', async () => {
    gl.on = true
    const r = rig({ slides: [flyIn()], force: 'none' })
    await r.enter(0)
    const tl = r.timelines.get(0)!
    r.stage.frame(tl.checkpoint(1), 0)
    expect(gl.made).toHaveLength(1)
    expect(gl.made[0]!.uploads).toBe(4)
    expect(r.stage.inspect().backend).toBe('webgl2')
    const glCanvas = r.root.querySelector('canvas[data-slideshow-backend="webgl2"]')!
    const ev = new Event('webglcontextlost', { cancelable: true })
    glCanvas.dispatchEvent(ev)
    expect(ev.defaultPrevented).toBe(true)
    expect(gl.made[0]!.lost).toBe(true)
    expect(r.stage.inspect().backend).toBe('canvas2d')
    expect(r.sink.list().some((d) => d.code === 'backend-fallback' && d.effect === 'context-lost')).toBe(true)
    // The retained bitmaps were re-uploaded: the next frame draws the layers on the 2-D output.
    r.stage.frame(tl.checkpoint(1), 10)
    expect(r.stage.inspect().lastDraws).toEqual(['layer:0', 'layer:1', 'layer:2'])
    expectPx(r.pixel(48, 36), BOX)
    expectPx(r.pixel(48, 63), FOOT)
  })

  it('workspace refusal holds the entry frame with plan-stage-pressure', async () => {
    // Room for the entry frame's texture and a little more: the layers do not fit.
    const ws = workspaceOf(OW * OH * 4 + 1000)
    const r = rig({ slides: [flyIn()], ws })
    await r.enter(0)
    const tl = r.timelines.get(0)!
    r.stage.frame(tl.checkpoint(1), 0)
    expect(r.stage.inspect().lastDraws).toEqual(['entry'])
    expectPx(r.pixel(48, 36), WHITE)
    expectPx(r.pixel(48, 63), FOOT)
    expect(r.sink.list().map((d) => d.code)).toContain('plan-stage-pressure')
    // Static: a click step changes nothing.
    r.stage.frame(tl.next(tl.initial(), 0), 250)
    expect(r.stage.inspect().lastDraws).toEqual(['entry'])
    expect(ws.used()).toBeLessThanOrEqual(OW * OH * 4 + 1000)
  })

  it('dispose releases every lease', async () => {
    const ws = workspaceOf()
    const base = ws.used()
    const steps: ShowPlanRecord['steps'] = [{ trigger: 'click', eligibility: 'composited', track: { fps: 15, count: 8 } }]
    const r = rig({ slides: [flyIn({ steps }), plainSlide(1, fade(500))], ws })
    r.tracks.set('0:0', fakeTrack(8, (i) => [i, i, i]))
    await r.enter(0)
    const tl = r.timelines.get(0)!
    r.stage.frame(tl.initial(), 0)
    r.stage.frame(tl.next(tl.initial(), 0), 200)
    r.leave()
    await r.enter(1)
    r.stage.frame(r.timelines.get(1)!.initial(), 100)
    const preview = document.createElement('canvas')
    preview.width = 48; preview.height = 36
    expect(r.stage.composeAt(0, tl.checkpoint(0), preview)).toBe(true)
    expect(ws.used()).toBeGreaterThan(base)
    r.stage.dispose()
    expect(ws.used()).toBe(base)
  })

  it('composeAt composes the next checkpoint from the streamed layers', async () => {
    const r = rig({ slides: [flyIn()] })
    await r.enter(0)
    const preview = document.createElement('canvas')
    preview.width = OW; preview.height = OH
    expect(r.stage.composeAt(0, r.timelines.get(0)!.checkpoint(1), preview)).toBe(true)
    const d = preview.getContext('2d')!.getImageData(48, 36, 1, 1).data
    expectPx([d[0]!, d[1]!, d[2]!], BOX)
    expect(r.stage.composeAt(7, r.timelines.get(0)!.checkpoint(1), preview)).toBe(false)
  })

  it('a still slide draws its still, and its still is its from-frame', async () => {
    const still = solid(OW, OH, [200, 10, 10])
    const r = rig({ slides: [plainSlide(1, fade(500))], stills: new Map([[0, still]]) })
    await r.enter(0)
    r.stage.frame(undefined, 0)
    expect(r.stage.inspect().lastDraws).toEqual(['still'])
    expectPx(r.pixel(10, 10), [200, 10, 10])
    r.leave()
    await r.enter(1)
    r.stage.frame(r.timelines.get(1)!.initial(), 0)
    expectPx(r.pixel(10, 10), [200, 10, 10])
  })
})
