/**
 * Piece 4 Task 5: the show-plan eligibility classifier (§6). No transmitted
 * layer may carry ink the composed show never displays; a step that cannot
 * be proven plays `composited`.
 */
import { afterAll, beforeAll, describe, expect, it } from 'bun:test'
import { installCanvasBackend, type CanvasBackend } from '../../src/core/canvas-backend'
import { createSlideTimeline } from '../../src/slideshow/engine/steps'
import type { PreparedSlide } from '../../src/slideshow/prepare'
import type { TimeNode, TimingTree } from '../../src/slideshow/types'
import {
  MEDIA_END_REVEAL, MOVER_REVEAL, adversaryDeck, mediaEndRevealDeck, autoplayRevealDeck, triggerFadeDeck, triggerPulseDeck, triggerPulseMainDeck, triggerRevealDeck, twinDeck, flashColourDeck, flyInDeck, loadShowDeck, longStepDeck, manyLayersDeck, moveTwiceDeck, spinStepsDeck, triggerColourDeck, triggerDartDeck, triggerDeck,
} from '../../src/testdata/show-decks'
import { serverBackend } from '../src/backend'
import {
  CLASSIFY_BUDGET, SAMPLES_PER_INTERVAL, TRACK_FPS, TRACK_MAX_FRAMES, allocateFrames, mediaEndAffectsDrawing, sampleTimes, sweptHull,
} from '../src/show/eligibility'
import { classifiedRasters, drawsAt, slideStateAt } from '../src/show/compose'
import { createDiagnosticSink } from '../../src/slideshow/diagnostics'
import { createRng } from '../../src/slideshow/rng'
import { classify, inkPixels, reconstruct } from './show-classify-helper'

let backend: CanvasBackend
let previous: CanvasBackend | undefined
beforeAll(() => { backend = serverBackend(); previous = installCanvasBackend(backend) })
afterAll(() => { installCanvasBackend(previous) })

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

  it('over 64 steps: the slide falls back to its still (show-limits)', async () => {
    const { c } = await classify(await spinStepsDeck(65, 100))
    expect(c.fallback).toBe('still')
    expect(c.steps).toHaveLength(65)
    expect(c.steps.every(s => s.eligibility === 'composited' && s.reason === 'limits')).toBe(true)
    expect(c.diagnostics).toContainEqual({ code: 'show-limits' })
    expect(c.drop.size + c.masks.size).toBe(0)
  })

  it('strictMarks: every moving step composited', async () => {
    const { c } = await classify(await flyInDeck(), { strictMarks: true })
    expect(c.steps[0]).toMatchObject({ eligibility: 'composited', reason: 'strict-marks' })
    const fade = await classify(await triggerFadeDeck(), { strictMarks: true })
    expect(fade.c.steps).toEqual([])
    expect(fade.c.triggers.get('3')).toBe('layers')
  })

  it('two rest positions → every step from the first move composited (mark-rest; D8 strengthened, Task 6 ruling)', async () => {
    const { c } = await classify(await moveTwiceDeck())
    expect(c.steps[0]).toMatchObject({ eligibility: 'composited', reason: 'mark-rest' })
    expect(c.steps[1]).toMatchObject({ eligibility: 'composited', reason: 'mark-rest' })
    expect(c.diagnostics).toContainEqual({ code: 'mark-rest', step: 0 })
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
    const { c } = await classify(await triggerFadeDeck())
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

  describe('the classification reconstructs the show (fix round 1)', () => {
    // PR #78 round 12: the client ends every audio/video node at once (`syntheticEnds`); the server's histories
    // carry no media end. Round 12b: a slide whose observations differ under the media-end assumptions plays its still.
    for (const kind of ['onEnd', 'afterPrevious', 'afterIndefinite'] as const) {
      it(`a media end reveals ink (${kind}): the slide plays its still`, async () => {
        const run = await classify(await mediaEndRevealDeck(kind))
        const { ps, timeline: tl } = run
        const ink = targetLayer(ps, 'sp:2')
        const drawsInk = (cursor: ReturnType<typeof tl.initial>, t: number) =>
          drawsAt(ps, slideStateAt(ps, tl, cursor, t, createRng(1), createDiagnosticSink())).some(d => d.layer === ink && d.state.visible && d.state.opacity > 0)
        // The fixture: no media end, no ink; the client's synthetic end (at the first frame) shows it.
        expect(drawsInk(tl.initial(), 100)).toBe(false)
        expect(drawsInk(tl.mediaEnd(tl.initial(), MEDIA_END_REVEAL.videoId, 16, true), 100)).toBe(true)
        expect(run.c.fallback).toBe('still')
        expect(tl.steps).toBe(0)
        expect(run.c.diagnostics).toContainEqual({ code: 'step-composited' })
        expect(run.c.drop.size + run.c.masks.size).toBe(0)
        expect(() => reconstruct(run, tl.mediaEnd(tl.initial(), MEDIA_END_REVEAL.videoId, 16, true), 100)).toThrow()
      })
    }

    it('an entrance whose END condition is the media end (codex counterexample): the slide plays its still (round 12c)', async () => {
      const run = await classify(await mediaEndRevealDeck('endCondition'))
      const { ps, timeline: tl } = run
      const ink = targetLayer(ps, 'sp:2')
      const drawsInk = (cursor: ReturnType<typeof tl.initial>, t: number) =>
        drawsAt(ps, slideStateAt(ps, tl, cursor, t, createRng(1), createDiagnosticSink())).some(d => d.layer === ink && d.state.visible && d.state.opacity > 0)
      // An early end (16 ms, before the entrance begins at 20 ms) is ignored; a delayed first tick (500 ms) cuts it.
      expect(drawsInk(tl.mediaEnd(tl.initial(), MEDIA_END_REVEAL.videoId, 16, true), 600)).toBe(true)
      expect(tl.stateAt(tl.initial(), 500).active.some(e => e.node.ctn.id === MEDIA_END_REVEAL.videoId)).toBe(true)
      expect(drawsInk(tl.mediaEnd(tl.initial(), MEDIA_END_REVEAL.videoId, 500, true), 600)).toBe(false)
      expect(run.c.fallback).toBe('still')
    })

    it('a fill-remove container that ends with its children (endSync) holding media and the ink: the still (round 12c)', async () => {
      const run = await classify(await mediaEndRevealDeck('endSync'))
      const { ps, timeline: tl } = run
      const ink = targetLayer(ps, 'sp:2')
      const drawsInk = (cursor: ReturnType<typeof tl.initial>, t: number) =>
        drawsAt(ps, slideStateAt(ps, tl, cursor, t, createRng(1), createDiagnosticSink())).some(d => d.layer === ink && d.state.visible && d.state.opacity > 0)
      expect(drawsInk(tl.initial(), 100)).toBe(true)
      expect(drawsInk(tl.mediaEnd(tl.initial(), MEDIA_END_REVEAL.videoId, 16, true), 100)).toBe(false)
      expect(run.c.fallback).toBe('still')
    })

    it('media with an unresolved reference: fails closed to the still (round 12c)', async () => {
      const run = await classify(await mediaEndRevealDeck('unresolved'))
      expect(run.c.fallback).toBe('still')
    })

    it('the dependency analysis fails closed on a reference cycle and past its bounds (round 12c)', async () => {
      // The reader breaks a deck's `tn` cycles (`timing-cycle`): the fixture's B no longer names A.
      const meta = await (await loadShowDeck(await mediaEndRevealDeck('cycle'))).meta(0)
      const tree = meta.timing as TimingTree
      const b = tree.root.kind === 'par' ? tree.root.children.find(n => n.ctn.id === 5)! : undefined
      expect(b?.ctn.start.some(c => c.tn === 4)).toBe(false)
      // A tree that reaches the analysis with the cycle intact fails closed, as does one past its bounds.
      const ctn = (id: number, start: TimeNode['ctn']['start'] = []) => ({ id, durMs: undefined, accel: 0, decel: 0, autoRev: false, fill: 'hold' as const, start, end: [] })
      const leaf = (kind: string, id: number, start: TimeNode['ctn']['start'] = []) => ({ kind, ctn: ctn(id, start), target: { spid: '2' }, attrNames: [], additive: 'base', accumulate: 'none' }) as unknown as TimeNode
      const par = (id: number, children: TimeNode[], start: TimeNode['ctn']['start'] = []): TimeNode => ({ kind: 'par', ctn: ctn(id, start), children })
      const onEnd = (tn: number) => [{ evt: 'onEnd', tn, delayMs: 0 }]
      const cyclic = par(1, [leaf('video', 2), par(4, [leaf('set', 6)], onEnd(5)), par(5, [leaf('set', 7)], onEnd(4))])
      expect(mediaEndAffectsDrawing(cyclic)).toBe(true)
      let deep: TimeNode = leaf('set', 1000)
      for (let i = 0; i < 300; i++) deep = par(2000 + i, [deep])
      expect(mediaEndAffectsDrawing(par(1, [leaf('video', 2), deep]))).toBe(true)
      // The same shapes without the cycle or the depth: nothing drawn waits on the video.
      expect(mediaEndAffectsDrawing(par(1, [leaf('video', 2), par(4, [leaf('set', 6)])]))).toBe(false)
      // A container whose end follows the video, holding a child that begins late (not at a fixed offset): its
      // begin may come after the client's end of the container (parent activity).
      expect(mediaEndAffectsDrawing(par(1, [par(3, [leaf('video', 2), par(4, [leaf('set', 6)], [{ delayMs: 'indefinite' }])])]))).toBe(true)
      expect(mediaEndAffectsDrawing(par(1, [par(3, [leaf('video', 2), par(4, [leaf('set', 6)], [{ delayMs: 0 }])])]))).toBe(false)
    })

    it('an effect sound inside click effect 1, then click 2 (the lo-tdf44223 shape): layers that reconstruct the client history (round 12c)', async () => {
      const run = await classify(await mediaEndRevealDeck('effectSound'))
      const { ps, timeline: tl, c } = run
      expect(tl.steps).toBe(2)
      expect(c.fallback).toBeUndefined()
      expect(c.steps.map(s => s.eligibility)).toEqual(['layers', 'layers'])
      expect(c.drop.has(targetLayer(ps, 'sp:2'))).toBe(false)
      // The client: click 1 at 100 ms, the sound ends on the next frame, click 2 at 1000 ms.
      const one = tl.next(tl.initial(), 100)
      expect(tl.stateAt(one, 110).active.some(e => e.node.ctn.id === MEDIA_END_REVEAL.soundId)).toBe(true)
      const ended = tl.mediaEnd(one, MEDIA_END_REVEAL.soundId, 116, true)
      const two = tl.next(ended, 1000)
      expect(two.step).toBe(2)
      for (const [cursor, t] of [[ended, 200], [ended, 400], [ended, 900], [two, 1100], [two, 1300], [two, 2000]] as const) {
        const r = reconstruct(run, cursor, t)
        expect({ t, diff: r.diff }).toEqual({ t, diff: 0 })
      }
      expect(reconstruct(run, two, 2000).inkOriginal).toBeGreaterThan(0)
    })

    it('an effect sound inside a click effect (the lo-tdf44223 shape): the slide stays eligible (round 12c)', () => {
      const ctn = (id: number, o: Partial<TimeNode['ctn']> = {}) => ({ id, durMs: undefined, accel: 0, decel: 0, autoRev: false, fill: 'hold' as const, start: [{ delayMs: 0 }], end: [], ...o }) as TimeNode['ctn']
      const leaf = (kind: string, id: number, o: Partial<TimeNode['ctn']> = {}) => ({ kind, ctn: ctn(id, o), target: { spid: '4' }, attrNames: [], additive: 'base', accumulate: 'none' }) as unknown as TimeNode
      const effect: TimeNode = { kind: 'par', ctn: ctn(5), children: [leaf('set', 6, { durMs: 1 }), leaf('anim', 7, { durMs: 500 }), leaf('audio', -1, { start: [{ evt: 'begin', tn: 5, delayMs: 0 }], end: [{ evt: 'onStopAudio', delayMs: 0 }] })] }
      const click: TimeNode = { kind: 'par', ctn: ctn(3, { start: [{ delayMs: 'indefinite' }] }), children: [{ kind: 'par', ctn: ctn(4), children: [effect] }] }
      const main: TimeNode = { kind: 'seq', ctn: ctn(2, { durMs: 'indefinite', fill: undefined, start: [] }), children: [click], seq: { concurrent: true, nextAc: 'seek', prevAc: 'none' } }
      const root: TimeNode = { kind: 'par', ctn: ctn(1, { durMs: 'indefinite', fill: undefined, start: [] }), children: [main] }
      expect(mediaEndAffectsDrawing(root)).toBe(false)
      // The same click effect with fill="remove" lets go of the fly-in when the sound ends: the still.
      effect.ctn.fill = 'remove'
      expect(mediaEndAffectsDrawing(root)).toBe(true)
    })

    it('media no drawing depends on, beside a click animation: the slide stays eligible and plays its layers (round 12b)', async () => {
      const run = await classify(await mediaEndRevealDeck('unrelated'))
      const { ps, timeline: tl, c } = run
      expect(tl.steps).toBe(1)
      expect(c.fallback).toBeUndefined()
      expect(c.steps.map(s => s.eligibility)).toEqual(['layers'])
      expect(c.drop.has(targetLayer(ps, 'sp:2'))).toBe(false)
      // The client's history (the video ended at once, then the click) reconstructs exactly.
      const ended = tl.mediaEnd(tl.initial(), MEDIA_END_REVEAL.videoId, 16, true)
      const clicked = tl.next(ended, 200)
      for (const t of [200, 450, 700, 1000]) {
        const r = reconstruct(run, clicked, t)
        expect({ t, diff: r.diff }).toEqual({ t, diff: 0 })
      }
      expect(reconstruct(run, clicked, 1000).inkOriginal).toBeGreaterThan(0)
    })

    it('unproven autoplay before the first click falls back to the still', async () => {
      const run = await classify(await autoplayRevealDeck())
      expect(run.timeline.steps).toBe(0)
      expect(run.c.fallback).toBe('still')
      expect(run.c.diagnostics).toContainEqual({ code: 'step-composited' })
      expect(run.c.drop.size + run.c.masks.size).toBe(0)
      expect(() => reconstruct(run, run.timeline.initial(), MOVER_REVEAL.dartAt + MOVER_REVEAL.dartMs / 2)).toThrow()
    })

    it('a cover-hiding trigger after a main step: the ink is kept and reconstructs', async () => {
      const run = await classify(await triggerRevealDeck())
      const { timeline: tl, ps, c } = run
      expect(c.fallback).toBeUndefined()
      expect(c.drop.has(targetLayer(ps, 'sp:2'))).toBe(false)
      const after = tl.trigger(tl.checkpoint(1), '3', 100)!
      const r = reconstruct(run, after, 110)
      expect(r.inkOriginal).toBeGreaterThan(0)
      expect(r).toEqual({ diff: 0, inkOriginal: r.inkOriginal, inkClassified: r.inkOriginal })
      // Every checkpoint, with and without the trigger.
      for (let k = 0; k <= tl.steps; k++) {
        expect(reconstruct(run, tl.checkpoint(k), 0).diff).toBe(0)
        expect(reconstruct(run, tl.trigger(tl.checkpoint(k), '3', 100)!, 110).diff).toBe(0)
      }
    })

    it('two identical overlapping layers: their joint removal never erases the shown box', async () => {
      const run = await classify(await twinDeck())
      const { timeline: tl, c } = run
      expect(c.steps[0].eligibility).toBe('layers')
      const r = reconstruct(run, tl.checkpoint(1), 0)
      expect(r.inkOriginal).toBeGreaterThan(1000)
      expect(r.diff).toBe(0)
      expect(r.inkClassified).toBe(r.inkOriginal)
      for (const t of sampleTimes(tl, 0)) expect(reconstruct(run, tl.next(tl.checkpoint(0), 0), t).diff).toBe(0)
    })

    it('every fixture classification reconstructs at every boundary', async () => {
      for (const bytes of [await flyInDeck(), await adversaryDeck('cover'), await adversaryDeck('moverReveal'), await moveTwiceDeck(), await triggerFadeDeck()]) {
        const run = await classify(bytes)
        for (let k = 0; k <= run.timeline.steps; k++) expect(reconstruct(run, run.timeline.checkpoint(k), 0).diff).toBe(0)
      }
    })

    it('crop: a decoded picture whose cropped-away half is the ink ships without it', async () => {
      const run = await classify(await adversaryDeck('crop'), { images: true })
      const { ps, c, timeline: tl } = run
      const pic = targetLayer(ps, 'sp:2')
      expect(c.drop.has(pic)).toBe(false)
      const shipped = classifiedRasters(run.rasters, c.drop, c.masks).get(pic)!
      const px = (shipped.canvas.getContext('2d') as unknown as CanvasRenderingContext2D).getImageData(0, 0, shipped.canvas.width, shipped.canvas.height).data
      expect(inkPixels(px)).toBe(0)
      // The shown half (#3366CC) is there.
      let blue = 0
      for (let i = 0; i < px.length; i += 4) if (px[i + 3] > 0 && Math.abs(px[i] - 0x33) <= 6 && Math.abs(px[i + 1] - 0x66) <= 6 && Math.abs(px[i + 2] - 0xcc) <= 6) blue++
      expect(blue).toBeGreaterThan(0)
      expect(reconstruct(run, tl.checkpoint(1), 0).diff).toBe(0)
    })
  
  })

  describe('trigger histories (fix round 2)', () => {
    /** Each history must either not play from layers at all (the still) or reconstruct exactly. */
    const safe = (run: Awaited<ReturnType<typeof classify>>, cursor: Parameters<typeof reconstruct>[1], t: number) => {
      if (run.c.fallback === 'still') return
      expect(reconstruct(run, cursor, t).diff).toBe(0)
    }

    it('a trigger mid-step, while the main pulse is up', async () => {
      const run = await classify(await triggerPulseMainDeck())
      const tl = run.timeline
      const mid = tl.trigger(tl.next(tl.checkpoint(0), 0), '3', 500)!
      expect(reconstruct({ ...run, c: { ...run.c, fallback: undefined, drop: new Set(), masks: new Map() } }, mid, 510).inkOriginal).toBeGreaterThan(1000)
      safe(run, mid, 510)
    })

    it('the main step played after the trigger', async () => {
      const run = await classify(await triggerPulseMainDeck())
      const tl = run.timeline
      const after = tl.next(tl.trigger(tl.checkpoint(0), '3', 0)!, 100)
      safe(run, after, 610)
    })

    it('an opacity-pulse trigger, first activation', async () => {
      const run = await classify(await triggerPulseDeck())
      const tl = run.timeline
      const first = tl.trigger(tl.checkpoint(0), '3', 0)!
      expect(reconstruct({ ...run, c: { ...run.c, fallback: undefined, drop: new Set(), masks: new Map() } }, first, 500).inkOriginal).toBeGreaterThan(1000)
      safe(run, first, 500)
    })

    it('an opacity-pulse trigger fired twice', async () => {
      const run = await classify(await triggerPulseDeck())
      const tl = run.timeline
      const twice = tl.trigger(tl.trigger(tl.checkpoint(0), '3', 0)!, '3', 1100)!
      safe(run, twice, 1600)
    })

    it('a trigger slide whose layers need no drop or mask stays layers', async () => {
      const run = await classify(await triggerFadeDeck())
      expect(run.c.fallback).toBeUndefined()
      expect(run.c.drop.size + run.c.masks.size).toBe(0)
      expect(run.c.triggers.get('3')).toBe('layers')
      const tl = run.timeline
      for (const t of [0, 100, 250, 500, 900]) expect(reconstruct(run, tl.trigger(tl.checkpoint(0), '3', 0)!, t).diff).toBe(0)
    })

    it('a trigger slide needing a drop or a mask plays the still', async () => {
      // The pulses' ink is never observed; triggerDeck's background is covered by its button.
      for (const bytes of [await triggerPulseDeck(), await triggerPulseMainDeck(), await triggerDeck()]) {
        const { c } = await classify(bytes)
        expect(c.drop.size + c.masks.size).toBe(0)
        expect(c.fallback).toBe('still')
        expect(c.diagnostics).toContainEqual({ code: 'step-composited' })
        expect([...c.triggers.values()].every(v => v === 'cut')).toBe(true)
      }
    })
  })
})
