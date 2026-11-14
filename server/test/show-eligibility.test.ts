/**
 * Piece 4 Task 5: the show-plan eligibility classifier (§6). No transmitted
 * layer may carry ink the composed show never displays; a step that cannot
 * be proven plays `composited`.
 */
import { afterAll, beforeAll, describe, expect, it } from 'bun:test'
import { installCanvasBackend, type CanvasBackend } from '../../src/core/canvas-backend'
import { createSlideTimeline } from '../../src/slideshow/engine/steps'
import { prepareSlide, rasterPreparedLayer, type PreparedSlide } from '../../src/slideshow/prepare'
import { createDiagnosticSink } from '../../src/slideshow/diagnostics'
import type { PlanLayer } from '../../src/slideshow/render/plan'
import type { CanvasFactory, LayerBitmap, Lease, SlideShowWorkspace } from '../../src/slideshow/types'
import {
  MOVER_REVEAL, adversaryDeck, flashColourDeck, flyInDeck, loadShowDeck, longStepDeck, manyLayersDeck, moveTwiceDeck, spinStepsDeck, triggerColourDeck, triggerDartDeck, triggerDeck,
} from '../../src/testdata/show-decks'
import { serverBackend } from '../src/backend'
import {
  CLASSIFY_BUDGET, SAMPLES_PER_INTERVAL, TRACK_FPS, TRACK_MAX_FRAMES, allocateFrames, classifySlide, sampleTimes, sweptHull, type Classification,
} from '../src/show/eligibility'

let backend: CanvasBackend
let previous: CanvasBackend | undefined
beforeAll(() => { backend = serverBackend(); previous = installCanvasBackend(backend) })
afterAll(() => { installCanvasBackend(previous) })

function roomyWorkspace(): SlideShowWorkspace {
  return {
    limit: Number.POSITIVE_INFINITY, used: () => 0, reserveTerminal: () => undefined,
    lease: (bytes, tag) => {
      const lease: Lease = { tag, bytes, grow: () => true, shrink() {}, transfer: () => lease, release() {} }
      return lease
    },
  }
}
const factory = (): CanvasFactory => (w, h) => backend.createCanvas(Math.max(1, Math.ceil(w)), Math.max(1, Math.ceil(h))) as never

interface Run { ps: PreparedSlide; c: Classification; rasters: Map<PlanLayer, LayerBitmap>; ms: number }

async function classify(bytes: Uint8Array, opts: { scale?: number; strictMarks?: boolean } = {}): Promise<Run> {
  const scale = opts.scale ?? 0.25
  const d = await loadShowDeck(bytes)
  const meta = await d.meta(0)
  const ps = prepareSlide({ index: 0, slide: d.doc.slides[0], meta, watermark: false })
  const sink = createDiagnosticSink()
  const rasters = new Map<PlanLayer, LayerBitmap>()
  for (const layer of ps.layers) {
    const r = rasterPreparedLayer(ps, layer, { scale, images: [], resolveFont: f => f, canvas: factory(), workspace: roomyWorkspace(), maxTextureSize: 8192, sink })
    if (!('refused' in r)) rasters.set(layer, r)
  }
  const timeline = createSlideTimeline(meta.timing, { showAnimation: true, reducedMotion: false })
  const t0 = performance.now()
  const c = classifySlide({ ps, timeline, scale, rasters, strictMarks: opts.strictMarks ?? false, triggers: timeline.triggers() })
  return { ps, c, rasters, ms: performance.now() - t0 }
}

const targetLayer = (ps: PreparedSlide, key: string) => ps.layers.find(l => l.kind === 'target' && l.key === key)!

describe('show eligibility', () => {
  it('exports the §6 constants', () => {
    expect([SAMPLES_PER_INTERVAL, CLASSIFY_BUDGET, TRACK_FPS, TRACK_MAX_FRAMES]).toEqual([8, 4096, 15, 90])
  })

  it('sampleTimes: 8 evenly spaced times plus every keyframe and behaviour begin/end time', async () => {
    const d = await loadShowDeck(await adversaryDeck('moverReveal'))
    const meta = await d.meta(0)
    const timeline = createSlideTimeline(meta.timing, { showAnimation: true, reducedMotion: false })
    const ts = sampleTimes(timeline, 0)
    for (let i = 0; i < 8; i++) expect(ts.some(t => Math.abs(t - (i * MOVER_REVEAL.stepMs) / 7) < 1e-6)).toBe(true)
    for (const t of [0, MOVER_REVEAL.dartAt, MOVER_REVEAL.dartAt + MOVER_REVEAL.dartMs, MOVER_REVEAL.stepMs]) expect(ts).toContain(t)
    // No sample lands inside the reveal window.
    expect(ts.every(t => t <= MOVER_REVEAL.reveal[0] || t >= MOVER_REVEAL.reveal[1])).toBe(true)
    expect([...ts].sort((a, b) => a - b)).toEqual(ts)
  })

  it('sweptHull covers every sampled bound and is widened by the largest per-sample displacement', () => {
    const id = [1, 0, 0, 1, 0, 0] as const
    const states = [0, 100, 300].map(dx => ({ visible: true, opacity: 1, transform: id, motion: { dx, dy: 0 } }))
    const [hull] = sweptHull(states, { x: 10, y: 20, w: 50, h: 40 })
    // Bounds sweep x 10..360; the largest step is 200 px.
    expect(hull).toEqual({ x: 10 - 200, y: 20 - 200, w: 350 + 400, h: 40 + 400 })
  })

  it('a fly-in over an unmasked background is layers', async () => {
    const { c } = await classify(await flyInDeck())
    expect(c.steps).toEqual([{ eligibility: 'layers' }])
    expect(c.drop.size).toBe(0)
  })

  it('never-observable layer is dropped (cover)', async () => {
    const { ps, c } = await classify(await adversaryDeck('cover'))
    expect(c.drop.has(targetLayer(ps, 'sp:2'))).toBe(true)
  })

  it('partly covered layer is masked to its envelope', async () => {
    const scale = 0.5
    const { ps, c, rasters } = await classify(await adversaryDeck('cover'), { scale })
    const bg = ps.layers[0]
    const mask = c.masks.get(bg)!
    expect(mask).toBeDefined()
    const bmp = rasters.get(bg)!
    const w = bmp.canvas.width, h = bmp.canvas.height
    let x0 = Infinity, y0 = Infinity, x1 = -Infinity, y1 = -Infinity
    for (let y = 0; y < h; y++) for (let x = 0; x < w; x++) {
      if (mask[y * w + x] !== 0) continue
      x0 = Math.min(x0, x); y0 = Math.min(y0, y); x1 = Math.max(x1, x + 1); y1 = Math.max(y1, y + 1)
    }
    // The cover: x 360–600, y 260–460 slide px.
    for (const [got, want] of [[x0, 360], [y0, 260], [x1, 600], [y1, 460]] as const) expect(Math.abs(got - want * scale)).toBeLessThanOrEqual(1)
    // Every pixel inside (by more than 1 px) is removed; everything outside is kept.
    let wrong = 0
    for (let y = 0; y < h; y++) for (let x = 0; x < w; x++) {
      const inside = x > 360 * scale + 1 && x < 600 * scale - 2 && y > 260 * scale + 1 && y < 460 * scale - 2
      const outside = x < 360 * scale - 1 || x > 600 * scale || y < 260 * scale - 1 || y > 460 * scale
      if ((inside && mask[y * w + x] !== 0) || (outside && mask[y * w + x] === 0)) wrong++
    }
    expect(wrong).toBe(0)
  })

  it('recolour-to-background step is composited (colour)', async () => {
    const { c } = await classify(await adversaryDeck('recolourBg'))
    expect(c.steps[0]).toMatchObject({ eligibility: 'composited', reason: 'colour' })
  })

  it('full mask is dropped', async () => {
    const { ps, c } = await classify(await adversaryDeck('fullMask'))
    expect(c.drop.has(targetLayer(ps, 'sp:2'))).toBe(true)
  })

  it('alpha-0 layer is dropped', async () => {
    const { ps, c } = await classify(await adversaryDeck('alpha0'))
    expect(c.drop.has(targetLayer(ps, 'sp:2'))).toBe(true)
  })

  it('exit before entrance is dropped', async () => {
    const { ps, c } = await classify(await adversaryDeck('exitBeforeEntrance'))
    expect(c.drop.has(targetLayer(ps, 'sp:2'))).toBe(true)
  })

  it('a fast mover crossing a masked layer between samples is composited by the swept hull', async () => {
    const { ps, c } = await classify(await adversaryDeck('moverReveal'))
    // The ink under the cover is never observed at a sample: the background layer is masked there.
    const bg = ps.layers[0]
    expect(c.masks.get(bg)?.some(a => a === 0)).toBe(true)
    expect(c.steps[0]).toMatchObject({ eligibility: 'composited', reason: 'swept' })
    expect(c.diagnostics).toContainEqual({ code: 'step-composited', step: 0 })
  })

  it('over budget: remaining steps composited (budget)', async () => {
    // 127 boxes plus the background: 128 layers, at the per-slide bound (128 boxes would be 129 layers: `limits`).
    // Spread over 5 steps they cost about 3,800 compositions (hidden and opacity-0 layers compose nothing), so
    // the same boxes are spread over 10 steps, which needs more than the budget.
    const { ps, c, ms } = await classify(await manyLayersDeck(127, { steps: 10 }), { scale: 0.125 })
    expect(ps.layers.length).toBe(128)
    expect(c.compositions).toBeLessThanOrEqual(CLASSIFY_BUDGET)
    expect(c.steps).toHaveLength(10)
    expect(c.steps[0].eligibility).toBe('layers')
    expect(c.steps.some(s => s.eligibility === 'composited' && s.reason === 'budget')).toBe(true)
    expect(c.diagnostics.some(d => d.code === 'show-budget')).toBe(true)
    console.log(`budget case: ${c.compositions} compositions in ${ms.toFixed(0)} ms`)
  }, 60_000)

  it('over 128 layers: every step composited (limits)', async () => {
    const { c } = await classify(await manyLayersDeck(128, { steps: 5 }), { scale: 0.125 })
    expect(c.steps).toHaveLength(5)
    expect(c.steps.every(s => s.eligibility === 'composited' && s.reason === 'limits')).toBe(true)
    expect(c.compositions).toBe(0)
  }, 60_000)

  it('over 64 steps: every step composited (limits) with show-limits', async () => {
    const { c } = await classify(await spinStepsDeck(65, 100))
    expect(c.steps).toHaveLength(65)
    expect(c.steps.every(s => s.eligibility === 'composited' && s.reason === 'limits')).toBe(true)
    expect(c.diagnostics).toContainEqual({ code: 'show-limits' })
  })

  it('strictMarks: every moving step composited', async () => {
    const { c } = await classify(await flyInDeck(), { strictMarks: true })
    expect(c.steps[0]).toMatchObject({ eligibility: 'composited', reason: 'strict-marks' })
    const fade = await classify(await triggerDeck(), { strictMarks: true })
    expect(fade.c.steps).toEqual([])
    expect(fade.c.triggers.get('3')).toBe('layers')
  })

  it('two rest positions → later move composited (mark-rest)', async () => {
    const { c } = await classify(await moveTwiceDeck())
    expect(c.steps[0]).toEqual({ eligibility: 'layers' })
    expect(c.steps[1]).toMatchObject({ eligibility: 'composited', reason: 'mark-rest' })
    expect(c.diagnostics).toContainEqual({ code: 'mark-rest', step: 1 })
  })

  it('long step gets 90 frames and step-truncated', async () => {
    const { c } = await classify(await longStepDeck())
    expect(c.steps[0]).toMatchObject({ eligibility: 'composited', frames: 90 })
    expect(c.diagnostics).toContainEqual({ code: 'step-truncated', step: 0 })
  })

  it('512 frames are allocated end-frames first', async () => {
    const frames = allocateFrames(Array.from({ length: 64 }, () => ({ composited: true, durMs: 2000 })))
    expect(frames.reduce((a, f) => a + f.frames, 0)).toBeLessThanOrEqual(512)
    expect(frames.every(f => f.frames >= 1)).toBe(true)
    expect(frames[0].frames).toBe(30)
    expect(frames.some(f => f.truncated)).toBe(true)
    // Through the classifier: 64 composited 2 s steps.
    const { c } = await classify(await spinStepsDeck(64, 2000), { strictMarks: true })
    expect(c.steps).toHaveLength(64)
    expect(c.steps.every(s => s.eligibility === 'composited' && (s.frames ?? 0) >= 1)).toBe(true)
    expect(c.steps.reduce((a, s) => a + (s.frames ?? 0), 0)).toBeLessThanOrEqual(512)
  }, 60_000)

  it('a visibility-only trigger plays from layers', async () => {
    const { c } = await classify(await triggerDeck())
    expect(c.triggers.get('3')).toBe('layers')
  })

  it('a trigger whose motion happens only between samples is cut', async () => {
    const { c } = await classify(await triggerDartDeck())
    expect(c.triggers.get('4')).toBe('cut')
    expect(c.diagnostics.some(d => d.code === 'trigger-cut')).toBe(true)
  })

  it('a recolour seen at no sample still makes its step composited (colour)', async () => {
    const { c } = await classify(await flashColourDeck())
    expect(c.steps[0]).toMatchObject({ eligibility: 'composited', reason: 'colour' })
  })

  it('trigger with a colour change is cut (trigger-cut)', async () => {
    const { c } = await classify(await triggerColourDeck())
    expect(c.triggers.get('3')).toBe('cut')
    expect(c.diagnostics.some(d => d.code === 'trigger-cut')).toBe(true)
  })
})
