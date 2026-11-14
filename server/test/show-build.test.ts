/** Piece 4 Task 6: the show-plan bundle build (plan record, entry frame, layers), its refusals, links, growth and caches. */
import { describe, expect, it } from 'bun:test'
import { decodeShowMessage, encodeShowMessage, isRasterLayer, SHOW_PLAN_LIMITS, type ShowPlanRecord } from '../../src/stream/show-plan'
const R = (p: ShowPlanRecord) => p.layers.filter(isRasterLayer)
import { ADVERSARY_KINDS, adversaryDeck, triggerFadeDeck, flyInDeck, manyLayersDeck, spinStepsDeck } from '../../src/testdata/show-decks'
import { Pool } from '../src/pool'
import { moveTwiceDeck } from '../../src/testdata/show-decks'
import { stateMatrix } from '../../src/slideshow/affine'
import { createDiagnosticSink } from '../../src/slideshow/diagnostics'
import { createRng } from '../../src/slideshow/rng'
import { trackInputs } from '../src/show/build'
import { drawsAt, slideStateAt } from '../src/show/compose'
import type { DocumentHost } from '../src/render'
import { appearMoveDeck, autoplayFadeDeck, revealThenMoveDeck, withFilter, withoutEntrance, withTriggeredMove } from './show-fixtures'
import { triggersMoveGeometry } from '../src/show/build'
import { createSlideComposer } from '../src/show/compose'
import { canvasBackend } from '../../src/core/canvas-backend'
import type { LayerBitmap } from '../../src/slideshow/types'
import { inkPixels } from './show-classify-helper'
import { decodePng } from '../src/mark/png'
import { SHOW_BUNDLE_BUDGET, type BuiltShowPlan, type ShowPlanJob } from '../src/show/build'
import { showPlanL1Key, showPlanL2Key } from '../src/cache'
import { cmdDeck, groupDeck, growthDeck, hiddenLinkDeck, masterTargetDeck, openDeck, transitionOnlyDeck } from './show-fixtures'

const job = (over: Partial<ShowPlanJob> = {}): ShowPlanJob => ({ unitIndex: 0, scale: 1, motion: 'full', mark: { kind: 'fixed' }, strictMarks: false, ...over })
const built = (r: Awaited<ReturnType<import('../src/render').DocumentHost['showPlan']>>): BuiltShowPlan => {
  if ('refused' in r) throw new Error(`refused: ${r.refused}`)
  return r
}
const planBytes = (plan: ShowPlanRecord) => new TextEncoder().encode(JSON.stringify(plan)).length

/** Every message of the bundle framed and decoded as the client would. */
function roundTrip(b: BuiltShowPlan) {
  const unit = b.plan.slide.index
  const expect_ = { unit, maxDecoded: SHOW_PLAN_LIMITS.showDecoded }
  const plan = decodeShowMessage(encodeShowMessage({ tag: 1, plan: b.plan }), expect_)
  const assets = b.assets.map(a => decodeShowMessage(encodeShowMessage({ tag: a.tag, unit, assetId: a.assetId, w: a.w, h: a.h, png: a.png }), expect_))
  return { plan, assets }
}

describe('buildShowPlan', () => {
  it('flyInDeck bundle: plan, entry frame, layers; ids unique; decodedBytes summed', async () => {
    const { host, open, docKey } = await openDeck(await flyInDeck())
    expect(open.anim).toEqual([0])
    const b = built(await host.showPlan(docKey, job()))
    const { plan } = b
    expect(plan.v).toBe(1)
    expect(plan.slide).toEqual({ index: 0, w: 960, h: 720 })
    expect(plan.options).toEqual({ showAnimation: true, useTimings: true })
    expect(plan.steps).toEqual([{ trigger: 'click', eligibility: 'layers' }])
    expect(R(plan).length).toBeGreaterThanOrEqual(2)
    expect(R(plan).some(l => l.key === 'sp:3')).toBe(true)
    expect(R(plan).filter(l => l.background).length).toBe(1)
    expect(plan.base.map(([k]) => k)).toContain('sp:3')
    // Order: the entry frame first, then the layers in z order.
    expect(b.assets[0]!.tag).toBe(3)
    expect(b.assets[0]!.assetId).toBe(plan.entry)
    expect(b.assets.slice(1).map(a => a.assetId)).toEqual([...R(plan)].sort((x, y) => x.z - y.z).map(l => l.assetId))
    expect(b.assets.slice(1).every(a => a.tag === 2)).toBe(true)
    const ids = b.assets.map(a => a.assetId)
    expect(new Set(ids).size).toBe(ids.length)
    // Each asset is what its record says.
    for (const l of R(plan)) {
      const a = b.assets.find(x => x.assetId === l.assetId)!
      expect(l.decodedBytes).toBe(a.w * a.h * 4)
      const d = decodePng(a.png)
      expect([d.w, d.h]).toEqual([a.w, a.h])
    }
    const entry = b.assets[0]!
    expect([entry.w, entry.h]).toEqual([960, 720])
    expect(b.decodedBytes).toBe(R(plan).reduce((s, l) => s + l.decodedBytes, 0) + entry.w * entry.h * 4 + planBytes(plan))
    // The client accepts every message.
    const rt = roundTrip(b)
    expect(rt.plan.tag).toBe(1)
    expect(rt.assets.map(m => m.tag)).toEqual(b.assets.map(a => a.tag))
  })

  it('transition-only slide ships layers: [] and an entry frame', async () => {
    const { host, open, docKey } = await openDeck(await transitionOnlyDeck())
    expect(open.anim).toEqual([0])
    const b = built(await host.showPlan(docKey, job()))
    expect(R(b.plan)).toEqual([])
    expect(b.plan.steps).toEqual([])
    expect(b.plan.show.transition?.element).toBe('fade')
    expect(b.assets.map(a => [a.tag, a.assetId])).toEqual([[3, b.plan.entry]])
    roundTrip(b)
  })

  it('129-layer slide plays every step composited', async () => {
    const { host, docKey } = await openDeck(await manyLayersDeck(129))
    const b = built(await host.showPlan(docKey, job()))
    expect(R(b.plan)).toEqual([])
    expect(b.plan.steps.length).toBe(5)
    for (const s of b.plan.steps) {
      expect(s.eligibility).toBe('composited')
      expect(s.track).toEqual({ fps: 15, count: Math.ceil(300 * 15 / 1000) })
    }
    expect(b.plan.diagnostics.some(d => d.code === 'step-composited')).toBe(true)
    expect(b.assets.map(a => a.tag)).toEqual([3])
    roundTrip(b)
    // The keyframe track of a composited step: count full frames, ids unique and distinct from the bundle's.
    const frames = await host.showTrack(docKey, { ...job(), step: 0 })
    expect(frames.length).toBe(b.plan.steps[0]!.track!.count)
    expect(frames.every(f => f.w === 960 && f.h === 720)).toBe(true)
    const ids = new Set([...frames.map(f => f.assetId), b.plan.entry])
    expect(ids.size).toBe(frames.length + 1)
    for (const f of frames) decodeShowMessage(encodeShowMessage({ tag: 4, unit: 0, assetId: f.assetId, w: f.w, h: f.h, png: f.png }), { unit: 0, maxDecoded: SHOW_PLAN_LIMITS.assetDecoded })
  }, 60_000)

  it('65-step slide is refused show-limits', async () => {
    const { host, docKey } = await openDeck(await spinStepsDeck(65, 100))
    const r = await host.showPlan(docKey, job())
    expect('refused' in r && r.refused).toBe('show-limits')
    expect('refused' in r && r.diagnostics.map(d => d.code)).toContain('show-limits')
  })

  it('bundle over 128 MiB is refused show-budget', async () => {
    expect(SHOW_BUNDLE_BUDGET as number).toBe(128 * 1024 * 1024)
    expect(SHOW_BUNDLE_BUDGET).toBe(SHOW_PLAN_LIMITS.showDecoded)
    const { host, docKey } = await openDeck(await flyInDeck())
    // Seam: the same check against a budget the bundle exceeds (a real 128 MiB slide costs minutes to classify).
    host.showOptions = { budget: 2 * 960 * 720 * 4 }
    const r = await host.showPlan(docKey, job())
    expect('refused' in r && r.refused).toBe('show-budget')
    expect('refused' in r && r.diagnostics.map(d => d.code)).toContain('show-budget')
    host.showOptions = {}
    expect('refused' in await host.showPlan(docKey, job())).toBe(false)
  })

  it('a build past the deadline is refused show-plan-timeout, and a repeat request plays the still without building', async () => {
    const { host, docKey } = await openDeck(await flyInDeck())
    host.showOptions = { deadlineMs: 0, now: (() => { let t = 0; return () => (t += 1) })() }
    const r = await host.showPlan(docKey, job())
    expect('refused' in r && r.refused).toBe('show-plan-timeout')
    expect('refused' in r && r.diagnostics.map(d => d.code)).toEqual(['show-plan-timeout'])
    const before = host.showCacheStats()
    host.showOptions = {}
    // Cached under the L1 key (controller ruling): any mark, no classification, no assembly.
    const again = await host.showPlan(docKey, job({ mark: { kind: 'viewer', text: 'someone else' } }))
    expect('refused' in again && again.refused).toBe('show-plan-timeout')
    const after = host.showCacheStats()
    expect(after.l2.misses).toBe(before.l2.misses)
    expect(after.l1.misses).toBe(before.l1.misses)
    // Another bucket is another slide preparation, and builds.
    expect('refused' in await host.showPlan(docKey, job({ scale: 2 }))).toBe(false)
  }, 60_000)

  it('timing targets resolve with the layout and master artwork, as in local mode', async () => {
    const { host, docKey } = await openDeck(await masterTargetDeck())
    built(await host.showPlan(docKey, job()))
    const show = (host as unknown as { docs: Map<string, { show: { sink: { list(): Array<{ code: string; effect?: string }> } } }> }).docs.get(docKey)!.show
    const codes = show.sink.list().map(d => d.code)
    expect(codes).toContain('master-target')
    expect(codes).not.toContain('unresolved-target')
  })

  it('a group with animated sub-layers ships a container record before its children', async () => {
    const { host, docKey } = await openDeck(await groupDeck())
    const b = built(await host.showPlan(docKey, job()))
    const containers = b.plan.layers.filter(l => !isRasterLayer(l))
    expect(containers.map(c => c.key)).toEqual(['sp:10'])
    const group = containers[0]!
    const children = R(b.plan).filter(l => l.parent === group.id)
    expect(children.map(c => c.key).sort()).toContain('sp:12')
    // Listed before every child; no asset of its own.
    const at = b.plan.layers.indexOf(group)
    for (const c of children) expect(b.plan.layers.indexOf(c)).toBeGreaterThan(at)
    expect(b.assets.some(a => a.assetId === (group as { assetId?: number }).assetId)).toBe(false)
    roundTrip(b)
  })

  /**
   * Every boundary the layers display (the end of a `layers` step, or the start of one): each shipped layer drawn
   * there sits exactly at the rest its mark was burned at (`asset.rest`).
   */
  async function restsHold(host: DocumentHost, docKey: string, b: BuiltShowPlan) {
    const show = (host as unknown as { docs: Map<string, { show: Parameters<typeof trackInputs>[0] }> }).docs.get(docKey)!.show
    const { ps, timeline } = trackInputs(show, job())
    const steps = b.plan.steps, n = steps.length
    let checked = 0
    for (let k = 0; k <= n; k++) {
      const layered = (k > 0 && steps[k - 1]!.eligibility === 'layers') || (k < n && steps[k]!.eligibility === 'layers')
      if (!layered) continue
      const draws = drawsAt(ps, slideStateAt(ps, timeline, timeline.checkpoint(k), 0, createRng(1), createDiagnosticSink()))
      for (const rec of R(b.plan)) {
        const d = draws.find(y => y.layer === ps.layers[rec.id])
        if (!d || !d.state.visible || !(d.state.opacity > 0)) continue
        const m = stateMatrix(d.state), a = b.assets.find(x => x.assetId === rec.assetId)!
        expect({ k, key: rec.key, x: Math.round((rec.bounds.x + m[4]) * job().scale), y: Math.round((rec.bounds.y + m[5]) * job().scale) }).toEqual({ k, key: rec.key, x: a.rest.x, y: a.rest.y })
        checked++
      }
    }
    return checked
  }

  it('D8 strengthened: a layer with two rests plays every step from its first move composited; no boundary shows a mark off its rest', async () => {
    const { host, docKey } = await openDeck(await moveTwiceDeck())
    const b = built(await host.showPlan(docKey, job()))
    expect(b.plan.steps.map(s => [s.eligibility, s.reason])).toEqual([['composited', 'mark-rest'], ['composited', 'mark-rest']])
    expect(b.plan.diagnostics.filter(d => d.code === 'mark-rest').map(d => d.step)).toEqual([0, 1])
    await restsHold(host, docKey, b)
    // Slides whose layers keep one rest still play from layers, and every displayed boundary holds it.
    for (const deck of [await flyInDeck(), await appearMoveDeck(0.25)]) {
      const o = await openDeck(deck)
      const ok = built(await o.host.showPlan(o.docKey, job()))
      expect(ok.plan.steps.map(s => s.eligibility)).toEqual(['layers'])
      expect(await restsHold(o.host, o.docKey, ok)).toBeGreaterThan(2)
    }
  }, 60_000)

  describe('fix round 1: triggers and autoplay keep the layers that draw them', () => {
    const showOf = (host: DocumentHost, docKey: string) => (host as unknown as { docs: Map<string, { show: Parameters<typeof trackInputs>[0] }> }).docs.get(docKey)!.show

    it('a trigger record carries cut: true exactly when the classifier cut it', async () => {
      const fade = await openDeck(await triggerFadeDeck())
      const layered = built(await fade.host.showPlan(fade.docKey, job({ mark: { kind: 'off' } })))
      expect(layered.plan.triggers).toEqual([expect.objectContaining({ spid: '3' })])
      expect(layered.plan.triggers[0]!.cut).toBeUndefined()
      expect(layered.plan.diagnostics.map(d => d.code)).not.toContain('trigger-cut')
      const wipe = await openDeck(await withFilter(await triggerFadeDeck(), 'fade', 'wipe(right)'))
      const cut = built(await wipe.host.showPlan(wipe.docKey, job({ mark: { kind: 'off' } })))
      expect(cut.plan.triggers.map(t => [t.spid, t.cut])).toEqual([['3', true]])
    }, 60_000)

    it('a cut trigger ships the layers that draw its end state, reconstructable from the shipped assets', async () => {
      const deck = await withFilter(await triggerFadeDeck(), 'fade', 'wipe(right)')
      const { host, docKey } = await openDeck(deck)
      const b = built(await host.showPlan(docKey, job({ mark: { kind: 'off' } })))
      expect(b.plan.steps).toEqual([])
      expect(b.plan.triggers.map(t => t.spid)).toEqual(['3'])
      expect(b.plan.diagnostics.map(d => d.code)).toContain('trigger-cut')
      expect(R(b.plan).length).toBeGreaterThan(0)
      // The trigger's end state, composed from the original rasters and from the decoded shipped assets.
      const { ps, timeline, rasters } = trackInputs(showOf(host, docKey), job({ mark: { kind: 'off' } }))
      const cur = timeline.trigger(timeline.checkpoint(0), '3', 0)!
      const keys = timeline.keyTimes(cur, 0, Number.POSITIVE_INFINITY)
      const draws = drawsAt(ps, slideStateAt(ps, timeline, cur, keys[keys.length - 1]! + 1, createRng(1), createDiagnosticSink()))
      const shipped = new Map<(typeof ps.layers)[number], LayerBitmap>()
      for (const rec of R(b.plan)) {
        const a = b.assets.find(x => x.assetId === rec.assetId)!, d = decodePng(a.png)
        const c = canvasBackend().createCanvas(d.w, d.h), x = c.getContext('2d') as unknown as CanvasRenderingContext2D
        const img = x.createImageData(d.w, d.h); img.data.set(d.rgba); x.putImageData(img, 0, 0)
        shipped.set(ps.layers[rec.id]!, { canvas: c as never, bounds: rec.bounds, scale: rec.scale, bytes: d.w * d.h * 4 })
      }
      const o = createSlideComposer(ps, rasters, 1), s2 = createSlideComposer(ps, shipped, 1)
      try {
        const want = o.compose(draws), got = s2.compose(draws)
        let box = 0, diff = 0
        for (let i = 0; i < want.length; i += 4) {
          if (want[i] === 0x33 && want[i + 1] === 0x66 && want[i + 2] === 0x99) box++
          if (want[i] !== got[i] || want[i + 1] !== got[i + 1] || want[i + 2] !== got[i + 2] || want[i + 3] !== got[i + 3]) diff++
        }
        expect(box).toBeGreaterThan(30_000)
        expect(diff).toBe(0)
      } finally { o.dispose(); s2.dispose() }
    }, 60_000)

    it('fix round 3: a trigger that moves a layer refuses the slide still (mark-rest)', async () => {
      // Round 2's single-trigger motion entrance: its rest is no longer enumerated, it is refused by construction.
      const { host, docKey } = await openDeck(await withTriggeredMove(await triggerFadeDeck()))
      const r = await host.showPlan(docKey, job())
      expect('refused' in r && r.refused).toBe('still')
      expect('refused' in r && r.diagnostics.map(d => d.code)).toContain('mark-rest')
    }, 60_000)

    it('fix round 3: reveal-then-move over two triggers is refused still (mark-rest)', async () => {
      const { host, docKey } = await openDeck(await revealThenMoveDeck())
      const r = await host.showPlan(docKey, job())
      expect('refused' in r && r.refused).toBe('still')
      expect('refused' in r && r.diagnostics.map(d => d.code)).toContain('mark-rest')
    }, 60_000)

    it('fix round 3: geometry in trigger sequences is detected; visibility, opacity, colour and masks are not', () => {
      const leaf = (kind: string, attrNames: string[] = []) => ({ kind, attrNames, ctn: { start: [] } })
      const seq = (child: unknown) => ({ root: { kind: 'par', ctn: { start: [] }, children: [{ kind: 'seq', ctn: { nodeType: 'interactiveSeq', start: [] }, children: [child] }] }, build: [] }) as never
      for (const k of ['animMotion', 'animRot', 'animScale']) expect(triggersMoveGeometry(seq(leaf(k)))).toBe(true)
      expect(triggersMoveGeometry(seq(leaf('anim', ['ppt_x'])))).toBe(true)
      expect(triggersMoveGeometry(seq(leaf('set', ['ppt_w'])))).toBe(true)
      for (const n of [leaf('set', ['style.visibility']), leaf('anim', ['style.opacity']), leaf('animClr', ['fillcolor']), leaf('animEffect')]) expect(triggersMoveGeometry(seq(n))).toBe(false)
      // A main-sequence motion is not a trigger's.
      expect(triggersMoveGeometry({ root: { kind: 'par', ctn: { start: [] }, children: [leaf('animMotion')] }, build: [] } as never)).toBe(false)
    })

    it('fix round 3: a trigger moving a shown layer to a second rest is refused still (mark-rest)', async () => {
      const deck = await withoutEntrance(await withTriggeredMove(await triggerFadeDeck()))
      const { host, docKey } = await openDeck(deck)
      const r = await host.showPlan(docKey, job())
      expect('refused' in r && r.refused).toBe('still')
      expect('refused' in r && r.diagnostics.map(d => d.code)).toContain('mark-rest')
    }, 60_000)

    const failing = { compose: (_k: number, img: ImageData) => ({ width: img.width, height: img.height, data: new Uint8ClampedArray(img.data.length) }) as unknown as ImageData }

    it('a trigger slide whose mark validation fails is refused still, never a plan without its trigger content', async () => {
      const { host, docKey } = await openDeck(await triggerFadeDeck())
      const control = built(await host.showPlan(docKey, job()))
      expect(R(control.plan).length).toBeGreaterThan(0)
      host.showOptions = failing
      const r = await host.showPlan(docKey, job())
      expect('refused' in r && r.refused).toBe('still')
      expect('refused' in r && r.diagnostics.map(d => d.code)).toContain('step-composited')
    }, 60_000)

    it('an autoplay-only slide whose mark validation fails is refused still', async () => {
      const { host, docKey } = await openDeck(await autoplayFadeDeck())
      const control = built(await host.showPlan(docKey, job()))
      expect(control.plan.steps).toEqual([])
      expect(R(control.plan).length).toBeGreaterThan(0)
      host.showOptions = failing
      const r = await host.showPlan(docKey, job())
      expect('refused' in r && r.refused).toBe('still')
      expect('refused' in r && r.diagnostics.length).toBeGreaterThan(0)
    }, 60_000)
  })

  it('cmd nodes kept inert (timing intact, command stripped) with cmd-dropped', async () => {
    // Controller ruling (Task 14), replacing D16's drop: the node keeps its place and timing, so the step and
    // click structure equals local; only the command payload is withheld.
    const { host, docKey } = await openDeck(await cmdDeck())
    const b = built(await host.showPlan(docKey, job()))
    const cmds: Array<{ kind: string; cmd?: string; cmdType?: string; ctn: { durMs?: unknown } }> = []
    const kinds: string[] = []
    const walk = (n: { kind: string; children?: unknown[] }) => { kinds.push(n.kind); if (n.kind === 'cmd') cmds.push(n as never); for (const c of n.children ?? []) walk(c as never) }
    const timing = b.plan.show.timing
    if (!timing || 'refused' in timing) throw new Error('no timing')
    walk(timing.root as never)
    expect(kinds).toContain('set')
    expect(cmds.length).toBe(1)
    expect(cmds[0]!.cmd).toBe('')
    expect(cmds[0]!.cmdType).toBe('verb')
    expect(cmds[0]!.ctn.durMs).toBe(1)
    expect(b.plan.steps.length).toBe(1)
    expect(b.plan.diagnostics.map(d => d.code)).toContain('cmd-dropped')
    roundTrip(b)
  })

  it('links: internal jump carries a streamed index; URL and hidden target are withheld', async () => {
    const { host, open, docKey } = await openDeck(await hiddenLinkDeck())
    expect(open.hidden).toEqual([1])
    // Slide 2 has neither timing nor a transition: no plan is requested for it (D5); slide 1 is hidden.
    expect(open.anim).toEqual([0])
    const b = built(await host.showPlan(docKey, job()))
    const byX = [...b.plan.linkRegions].sort((p, q) => p.rect.y - q.rect.y || p.rect.x - q.rect.x)
    expect(byX.map(l => l.action)).toEqual([{ kind: 'slide', index: 1 }, { kind: 'withheld' }, { kind: 'withheld' }])
    // Hit regions are the target's painted bounds (§5.1): they hold the shape's box, with at most a few px of margin.
    const r = byX[0]!.rect
    expect(r.x).toBeLessThanOrEqual(80); expect(r.y).toBeLessThanOrEqual(80)
    expect(r.x + r.w).toBeGreaterThanOrEqual(320); expect(r.y + r.h).toBeGreaterThanOrEqual(200)
    expect(r.w).toBeLessThanOrEqual(240 + 24); expect(r.h).toBeLessThanOrEqual(120 + 24)
    expect(b.plan.diagnostics.map(d => d.code)).toContain('link-withheld')
    roundTrip(b)
  })

  it('growth: animScale 2 layer raster is 2× the bucket, capped at 8192', async () => {
    {
      const { host, docKey } = await openDeck(await growthDeck(2))
      const b = built(await host.showPlan(docKey, job({ scale: 1 })))
      const grow = R(b.plan).find(l => l.key === 'sp:2')!
      expect(grow.scale).toBeCloseTo(2, 6)
      const a = b.assets.find(x => x.assetId === grow.assetId)!
      expect(Math.abs(a.w - grow.bounds.w * 2)).toBeLessThanOrEqual(2)
      expect(b.plan.steps.map(s => s.eligibility)).toEqual(['layers'])
      expect(R(b.plan).find(l => l.background)!.scale).toBe(1)
      roundTrip(b)
    }
    {
      // A full-width strip grown 10× at bucket 2 would be 19,200 px wide: capped at maxUnitPx.
      const { host, docKey } = await openDeck(await growthDeck(10, { x: 0, y: 340, w: 960, h: 24 }))
      const b = built(await host.showPlan(docKey, job({ scale: 2 })))
      const grow = R(b.plan).find(l => l.key === 'sp:2')!
      const a = b.assets.find(x => x.assetId === grow.assetId)!
      expect(Math.max(a.w, a.h)).toBeLessThanOrEqual(8192)
      expect(Math.max(a.w, a.h)).toBeGreaterThan(8000)
      expect(grow.scale).toBeGreaterThan(8)
      expect(grow.scale).toBeLessThan(10 * 2)
      roundTrip(b)
    }
  }, 60_000)

  it('L1 cache hit across two mark texts; L2 differs', async () => {
    const { host, docKey } = await openDeck(await flyInDeck())
    const alice = built(await host.showPlan(docKey, job({ mark: { kind: 'viewer', text: 'alice@example.com' } })))
    expect(host.showCacheStats()).toEqual({ l1: { hits: 0, misses: 1 }, l2: { hits: 0, misses: 1 } })
    const bob = built(await host.showPlan(docKey, job({ mark: { kind: 'viewer', text: 'bob@example.com' } })))
    expect(host.showCacheStats()).toEqual({ l1: { hits: 1, misses: 1 }, l2: { hits: 0, misses: 2 } })
    // The same plan; differently marked pixels.
    expect(bob.plan).toEqual(alice.plan)
    const differ = alice.assets.some((a, i) => Buffer.compare(Buffer.from(a.png), Buffer.from(bob.assets[i]!.png)) !== 0)
    expect(differ).toBe(true)
    await host.showPlan(docKey, job({ mark: { kind: 'viewer', text: 'alice@example.com' } }))
    expect(host.showCacheStats()).toEqual({ l1: { hits: 1, misses: 1 }, l2: { hits: 1, misses: 2 } })
    // Another bucket or motion is another L1 entry.
    await host.showPlan(docKey, job({ motion: 'reduced' }))
    expect(host.showCacheStats().l1.misses).toBe(2)
    // The keys (D6): L1 has no mark; L2 adds the mark settings and strictMarks.
    const f = { docRef: 'd', contentPolicyId: 'full@1', layoutFingerprint: 'x', scaleBucket: 1, motion: 'full' as const, unitId: 0 }
    expect(showPlanL1Key(f)).toBe(showPlanL1Key({ ...f }))
    expect(showPlanL1Key(f)).not.toBe(showPlanL1Key({ ...f, motion: 'reduced' }))
    expect(showPlanL1Key(f)).not.toBe(showPlanL1Key({ ...f, scaleBucket: 2 }))
    const l1 = showPlanL1Key(f)
    const mark = { visible: 'viewer' as const, text: 'alice', forensic: false as const }
    expect(showPlanL2Key(l1, mark, false)).not.toBe(showPlanL2Key(l1, { ...mark, text: 'bob' }, false))
    expect(showPlanL2Key(l1, mark, false)).not.toBe(showPlanL2Key(l1, mark, true))
  })

  it('adversary decks: no transmitted asset or keyframe carries the hidden ink', async () => {
    for (const kind of ADVERSARY_KINDS) {
      const { host, docKey } = await openDeck(await adversaryDeck(kind), kind)
      const r = await host.showPlan(docKey, job())
      if ('refused' in r) continue
      const pngs = [...r.assets.map(a => a.png)]
      for (let k = 0; k < r.plan.steps.length; k++) for (const f of await host.showTrack(docKey, { ...job(), step: k })) pngs.push(f.png)
      const ink = pngs.reduce((n, p) => n + inkPixels(decodePng(p).rgba), 0)
      expect({ kind, ink }).toEqual({ kind, ink: 0 })
    }
  }, 120_000)
})

describe('pool show jobs', () => {
  it('showPlan, showTrack and markLayer run on the pool under the show deadline', async () => {
    const pool = new Pool({ workers: 1 })
    try {
      await pool.open('fly', await flyInDeck(), 'full')
      const r = await pool.showPlan('s1', { docRef: 'fly', contentPolicy: 'full', ...job() }, 64 * 1024 * 1024)
      if ('refused' in r) throw new Error(r.refused)
      expect(r.plan.steps.map(s => s.eligibility)).toEqual(['layers'])
      expect(await pool.showTrack('s1', { docRef: 'fly', contentPolicy: 'full', ...job(), step: 0 }, 1024)).toEqual([])
      const layer = r.assets.find(a => a.tag === 2 && a.w * a.h > 10_000)!
      const m = await pool.markLayer('s1', { png: layer.png, w: layer.w, h: layer.h, kBasis: new Uint8Array(32).fill(7), code: Int8Array.from({ length: 512 }, (_, i) => (i % 3 ? 1 : -1)), crop: { ...layer.rest, pageW: 960, pageH: 720 } }, 1024)
      expect('carriers' in m && m.carriers).toBeGreaterThan(0)
      expect(pool.stats()).toMatchObject({ reserved: 0, running: 0 })
    } finally {
      pool.dispose()
    }
  }, 60_000)

  it('a timed-out slide is answered from the pool refusal cache: no admission, no worker, no second retirement', async () => {
    const pool = new Pool({ workers: 1, testHooks: true, showDeadlineMs: 400 })
    try {
      await pool.open('fly', await flyInDeck(), 'full')
      const first = await pool.showPlan('s1', { docRef: 'fly', contentPolicy: 'full', ...job(), hangMs: 1500 }, 1024)
      expect('refused' in first && first.refused).toBe('show-plan-timeout')
      expect(pool.stats().retired).toBe(1)
      const t0 = performance.now()
      const again = await pool.showPlan('s2', { docRef: 'fly', contentPolicy: 'full', ...job({ mark: { kind: 'off' } }) }, 1024)
      expect('refused' in again && again.refused).toBe('show-plan-timeout')
      expect(performance.now() - t0).toBeLessThan(100)
      expect(pool.stats()).toMatchObject({ retired: 1, reserved: 0, running: 0 })
      // Re-registering the document forgets it: the next request runs on a worker again (and, at 400 ms, times out again).
      await pool.open('fly', await flyInDeck(), 'full')
      await pool.showPlan('s1', { docRef: 'fly', contentPolicy: 'full', ...job(), hangMs: 1500 }, 1024)
      expect(pool.stats().retired).toBe(2)
    } finally {
      pool.dispose()
    }
  }, 60_000)

  it('show jobs run under the 10 s show deadline, not the 2 s render deadline', async () => {
    const pool = new Pool({ workers: 1, testHooks: true })
    try {
      await pool.open('fly', await flyInDeck(), 'full')
      const t0 = performance.now()
      const r = await pool.showPlan('s1', { docRef: 'fly', contentPolicy: 'full', ...job(), hangMs: 3000 }, 1024)
      expect(performance.now() - t0).toBeGreaterThan(3000)
      expect('refused' in r).toBe(false)
      expect(pool.stats().retired).toBe(0)
    } finally {
      pool.dispose()
    }
  }, 60_000)
})
