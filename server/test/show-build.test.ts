/** Piece 4 Task 6: the show-plan bundle build (plan record, entry frame, layers), its refusals, links, growth and caches. */
import { describe, expect, it } from 'bun:test'
import { decodeShowMessage, encodeShowMessage, SHOW_PLAN_LIMITS, type ShowPlanRecord } from '../../src/stream/show-plan'
import { ADVERSARY_KINDS, adversaryDeck, flyInDeck, manyLayersDeck, spinStepsDeck } from '../../src/testdata/show-decks'
import { Pool } from '../src/pool'
import { inkPixels } from './show-classify-helper'
import { decodePng } from '../src/mark/png'
import { SHOW_BUNDLE_BUDGET, type BuiltShowPlan, type ShowPlanJob } from '../src/show/build'
import { showPlanL1Key, showPlanL2Key } from '../src/cache'
import { cmdDeck, growthDeck, hiddenLinkDeck, openDeck, transitionOnlyDeck } from './show-fixtures'

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
    expect(plan.layers.length).toBeGreaterThanOrEqual(2)
    expect(plan.layers.some(l => l.key === 'sp:3')).toBe(true)
    expect(plan.layers.filter(l => l.background).length).toBe(1)
    expect(plan.base.map(([k]) => k)).toContain('sp:3')
    // Order: the entry frame first, then the layers in z order.
    expect(b.assets[0]!.tag).toBe(3)
    expect(b.assets[0]!.assetId).toBe(plan.entry)
    expect(b.assets.slice(1).map(a => a.assetId)).toEqual([...plan.layers].sort((x, y) => x.z - y.z).map(l => l.assetId))
    expect(b.assets.slice(1).every(a => a.tag === 2)).toBe(true)
    const ids = b.assets.map(a => a.assetId)
    expect(new Set(ids).size).toBe(ids.length)
    // Each asset is what its record says.
    for (const l of plan.layers) {
      const a = b.assets.find(x => x.assetId === l.assetId)!
      expect(l.decodedBytes).toBe(a.w * a.h * 4)
      const d = decodePng(a.png)
      expect([d.w, d.h]).toEqual([a.w, a.h])
    }
    const entry = b.assets[0]!
    expect([entry.w, entry.h]).toEqual([960, 720])
    expect(b.decodedBytes).toBe(plan.layers.reduce((s, l) => s + l.decodedBytes, 0) + entry.w * entry.h * 4 + planBytes(plan))
    // The client accepts every message.
    const rt = roundTrip(b)
    expect(rt.plan.tag).toBe(1)
    expect(rt.assets.map(m => m.tag)).toEqual(b.assets.map(a => a.tag))
  })

  it('transition-only slide ships layers: [] and an entry frame', async () => {
    const { host, open, docKey } = await openDeck(await transitionOnlyDeck())
    expect(open.anim).toEqual([0])
    const b = built(await host.showPlan(docKey, job()))
    expect(b.plan.layers).toEqual([])
    expect(b.plan.steps).toEqual([])
    expect(b.plan.show.transition?.element).toBe('fade')
    expect(b.assets.map(a => [a.tag, a.assetId])).toEqual([[3, b.plan.entry]])
    roundTrip(b)
  })

  it('129-layer slide plays every step composited', async () => {
    const { host, docKey } = await openDeck(await manyLayersDeck(129))
    const b = built(await host.showPlan(docKey, job()))
    expect(b.plan.layers).toEqual([])
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

  it('a build past the deadline is refused show-plan-timeout, and the refusal is not cached', async () => {
    const { host, docKey } = await openDeck(await flyInDeck())
    host.showOptions = { deadlineMs: 0, now: (() => { let t = 0; return () => (t += 1) })() }
    const r = await host.showPlan(docKey, job())
    expect('refused' in r && r.refused).toBe('show-plan-timeout')
    expect('refused' in r && r.diagnostics.map(d => d.code)).toEqual(['show-plan-timeout'])
    host.showOptions = {}
    expect('refused' in await host.showPlan(docKey, job())).toBe(false)
    expect(host.showCacheStats().l2.hits).toBe(0)
  })

  it('cmd nodes dropped with cmd-dropped', async () => {
    const { host, docKey } = await openDeck(await cmdDeck())
    const b = built(await host.showPlan(docKey, job()))
    const kinds: string[] = []
    const walk = (n: { kind: string; children?: unknown[] }) => { kinds.push(n.kind); for (const c of n.children ?? []) walk(c as never) }
    const timing = b.plan.show.timing
    if (!timing || 'refused' in timing) throw new Error('no timing')
    walk(timing.root as never)
    expect(kinds).not.toContain('cmd')
    expect(kinds).toContain('set')
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
      const grow = b.plan.layers.find(l => l.key === 'sp:2')!
      expect(grow.scale).toBeCloseTo(2, 6)
      const a = b.assets.find(x => x.assetId === grow.assetId)!
      expect(Math.abs(a.w - grow.bounds.w * 2)).toBeLessThanOrEqual(2)
      expect(b.plan.steps.map(s => s.eligibility)).toEqual(['layers'])
      expect(b.plan.layers.find(l => l.background)!.scale).toBe(1)
      roundTrip(b)
    }
    {
      // A full-width strip grown 10× at bucket 2 would be 19,200 px wide: capped at maxUnitPx.
      const { host, docKey } = await openDeck(await growthDeck(10, { x: 0, y: 340, w: 960, h: 24 }))
      const b = built(await host.showPlan(docKey, job({ scale: 2 })))
      const grow = b.plan.layers.find(l => l.key === 'sp:2')!
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
