/**
 * Piece 4 Task 9 — the T5 no-package prototype (spec §9 item 2, the parent's
 * "Stage-0/3c blocker"; plan D19, D25). A streamed show plays end to end
 * with the package, `PptxDocument` and `PptxShape` unreachable: every module
 * that reads or holds them throws on import (`vi.mock` factories that
 * throw), and the show is driven by the shared `createShowDriver` over
 * `createPlanStage`, from wire bytes only.
 *
 * The wire bytes are the recorded harness sessions in
 * `tests/fixtures/show-plan/{fly,morph}.bin` (`cd server && bun
 * ../scripts/record-show-fixtures.ts`; the format is in that script): the
 * kind-6 show-plan, show-track and show-morph messages, and the piece-3
 * unit message of every still slide. They are decoded here by the client
 * codecs (`decodeShowMessage`, `decodeUnit`, `decodeRasterBody`), and the
 * PNGs by node-canvas (jsdom's canvas), never by server code.
 *
 * Mocked to throw (the brief's list; `src/core/package` does not exist —
 * the package is `OfficePackage` in `core/zip`, and `core/package-sniff` /
 * `core/package-error` are the `core/package*` modules — and `render/backend`
 * and `player/hit-test` are added from the import gate's pattern).
 */
import { loadImage } from 'canvas'
import { readFileSync } from 'node:fs'
import { resolve } from 'node:path'
import { describe, expect, it, vi } from 'vitest'

const { boom } = vi.hoisted(() => ({ boom: (m: string) => () => { throw new Error(`T5: package module reached: ${m}`) } }))
vi.mock('../../src/pptx/parse', boom('pptx/parse'))
vi.mock('../../src/pptx/types', boom('pptx/types'))
vi.mock('../../src/core/zip', boom('core/zip'))
vi.mock('../../src/core/package-sniff', boom('core/package-sniff'))
vi.mock('../../src/core/package-error', boom('core/package-error'))
vi.mock('../../src/slideshow/read/store', boom('read/store'))
vi.mock('../../src/slideshow/read/deck', boom('read/deck'))
vi.mock('../../src/slideshow/prepare/index', boom('prepare/index'))
vi.mock('../../src/slideshow/render/raster', boom('render/raster'))
vi.mock('../../src/slideshow/render/plan', boom('render/plan'))
vi.mock('../../src/slideshow/render/direct', boom('render/direct'))
vi.mock('../../src/slideshow/render/scene', boom('render/scene'))
vi.mock('../../src/slideshow/render/backend', boom('render/backend'))
vi.mock('../../src/slideshow/morph/match', boom('morph/match'))
vi.mock('../../src/slideshow/player/stage', boom('player/stage'))
vi.mock('../../src/slideshow/player/session', boom('player/session'))
vi.mock('../../src/slideshow/player/text-builds', boom('player/text-builds'))
vi.mock('../../src/slideshow/player/hit-test', boom('player/hit-test'))

import { ManualClock } from '../../src/slideshow/clock'
import { createDiagnosticSink } from '../../src/slideshow/diagnostics'
import { createShowDriver, type DriverHooks, type SlideSource } from '../../src/slideshow/player/driver'
import type { ShowSlide } from '../../src/slideshow/types'
import { createSlideShowWorkspace } from '../../src/slideshow/workspace'
import { decodeRasterBody } from '../../src/stream/raster'
import { createPlanStage, createTrackWindow, type PlanAssets, type PlanSlideData, type PlanStageDeps, type PlanTrack } from '../../src/stream/remote/plan-stage'
import type { ImageBitmapLike } from '../../src/stream/remote/show-fetch'
import { decodeShowMessage, isRasterLayer, SHOW_PLAN_LIMITS, type MorphPlanRecord, type ShowMessage, type ShowPlanRecord } from '../../src/stream/show-plan'
import { decodeUnit } from '../../src/stream/unit'

const MOCKED = [
  '../../src/pptx/parse', '../../src/pptx/types', '../../src/core/zip', '../../src/core/package-sniff', '../../src/core/package-error',
  '../../src/slideshow/read/store', '../../src/slideshow/read/deck', '../../src/slideshow/prepare/index',
  '../../src/slideshow/render/raster', '../../src/slideshow/render/plan', '../../src/slideshow/render/direct', '../../src/slideshow/render/scene', '../../src/slideshow/render/backend',
  '../../src/slideshow/morph/match', '../../src/slideshow/player/stage', '../../src/slideshow/player/session', '../../src/slideshow/player/text-builds', '../../src/slideshow/player/hit-test',
]

// ---------------------------------------------------------------------------
// The recorded wire session
// ---------------------------------------------------------------------------

interface Rec { op: number; unit: number; a: number; b: number; bytes: Uint8Array }
function records(name: string): Rec[] {
  const buf = new Uint8Array(readFileSync(resolve(__dirname, '../fixtures/show-plan', `${name}.bin`)))
  const v = new DataView(buf.buffer, buf.byteOffset, buf.byteLength)
  const out: Rec[] = []
  for (let at = 0; at < buf.length;) {
    const len = v.getUint32(at + 13, true)
    out.push({ op: v.getUint8(at), unit: v.getUint32(at + 1, true), a: v.getUint32(at + 5, true), b: v.getUint32(at + 9, true), bytes: buf.subarray(at + 17, at + 17 + len) })
    at += 17 + len
  }
  return out
}

/** A PNG → a jsdom canvas (node-canvas decodes it). */
async function bitmapOf(png: Uint8Array): Promise<HTMLCanvasElement> {
  const img = await loadImage(Buffer.from(png))
  const cv = document.createElement('canvas')
  cv.width = img.width; cv.height = img.height
  cv.getContext('2d')!.drawImage(img as unknown as CanvasImageSource, 0, 0)
  return cv
}

const assetsOf = (m: Map<number, ImageBitmapLike>): PlanAssets => ({ bitmap: (id) => m.get(id), bytes: (id) => { const b = m.get(id); return b ? b.width * b.height * 4 : 0 } })

interface Session {
  count: number
  visible: number[]
  plans: Map<number, { plan: ShowPlanRecord; assets: Map<number, ImageBitmapLike>; tracks: Map<number, Array<Extract<ShowMessage, { png: Uint8Array }>>> }>
  stills: Map<number, ImageBitmapLike>
  morphs: Map<number, { from: number; fromStep: number; record: MorphPlanRecord; assets: Map<number, ImageBitmapLike> }>
}

async function session(name: string): Promise<Session> {
  const s: Session = { count: 0, visible: [], plans: new Map(), stills: new Map(), morphs: new Map() }
  for (const r of records(name)) {
    if (r.op === 0) { const h = JSON.parse(new TextDecoder().decode(r.bytes)) as { units: number; visible: number[] }; s.count = h.units; s.visible = h.visible; continue }
    if (r.op === 4) {
      const u = decodeUnit(r.bytes)
      expect(u.discriminator).toBe(2)
      s.stills.set(r.unit, await bitmapOf(decodeRasterBody(u.body).png))
      continue
    }
    const m = decodeShowMessage(r.bytes, { unit: r.unit, maxDecoded: r.op === 1 && r.bytes[0] === 1 ? SHOW_PLAN_LIMITS.showDecoded : SHOW_PLAN_LIMITS.assetDecoded })
    if (r.op === 1) {
      if (m.tag === 1) { s.plans.set(r.unit, { plan: m.plan, assets: new Map(), tracks: new Map() }); continue }
      if (m.tag !== 2 && m.tag !== 3) throw new Error(`plan message tag ${m.tag}`)
      s.plans.get(r.unit)!.assets.set(m.assetId, await bitmapOf(m.png))
    } else if (r.op === 2) {
      if (m.tag !== 4) throw new Error(`track message tag ${m.tag}`)
      const p = s.plans.get(r.unit)!
      if (!p.tracks.has(r.a)) p.tracks.set(r.a, [])
      p.tracks.get(r.a)!.push(m)
    } else if (r.op === 3) {
      if (m.tag === 5) { s.morphs.set(r.unit, { from: r.a, fromStep: r.b, record: m.morph, assets: new Map() }); continue }
      if (m.tag !== 2) throw new Error(`morph message tag ${m.tag}`)
      s.morphs.get(r.unit)!.assets.set(m.assetId, await bitmapOf(m.png))
    }
  }
  return s
}

// ---------------------------------------------------------------------------
// The prototype: driver + PlanStage over the recorded session
// ---------------------------------------------------------------------------

const OW = 96, OH = 72
const tick = () => new Promise((res) => setTimeout(res, 0))

async function play(s: Session) {
  const root = document.createElement('div')
  document.body.appendChild(root)
  root.getBoundingClientRect = () => ({ left: 0, top: 0, width: OW, height: OH, right: OW, bottom: OH, x: 0, y: 0, toJSON: () => ({}) })
  const output = document.createElement('canvas')
  output.width = OW; output.height = OH
  root.appendChild(output)
  const workspace = createSlideShowWorkspace('desktop')
  const sink = createDiagnosticSink()
  const clock = new ManualClock()
  const windows: Array<{ dispose(): void }> = []
  const data = new Map<number, PlanSlideData>()
  for (const [u, p] of s.plans) {
    const tracks = new Map<number, PlanTrack>()
    for (const [k, frames] of p.tracks) {
      const f = frames[0]!
      const w = createTrackWindow({ count: frames.length, frameBytes: f.w * f.h * 4, workspace, decode: (i) => bitmapOf(frames[i]!.png) })
      windows.push(w)
      tracks.set(k, w)
    }
    data.set(u, { plan: p.plan, assets: assetsOf(p.assets), track: (k) => tracks.get(k) })
  }
  const morphAsked: string[] = []
  const morph: PlanStageDeps['morph'] = async (from, to, fromStep) => {
    morphAsked.push(`${from}→${to}@${fromStep}`)
    const m = s.morphs.get(to)
    if (!m || m.from !== from || m.fromStep !== fromStep) return undefined
    return { record: m.record, assets: assetsOf(m.assets) }
  }
  const stage = createPlanStage({ root, output, workspace, sink, reduced: false, showNow: () => clock.now(), data: (i) => data.get(i), still: (i) => s.stills.get(i), morph, force: 'canvas2d' })
  // The slide source a streamed show reads: the plan's show metadata, or a static still slide.
  const slideOf = (i: number): ShowSlide => {
    const plan = s.plans.get(i)?.plan
    return { index: i, hidden: false, timing: plan?.show.timing, ...(plan?.show.transition ? { transition: plan.show.transition } : {}), builds: plan?.show.builds ?? [], notes: [], links: [], media: [] }
  }
  const source: SlideSource = {
    count: s.count,
    settings: { showAnimation: true, useTimings: true, loop: false, type: 'present' },
    hidden: (i) => !s.visible.includes(i),
    acquire: async (i) => ({ slide: slideOf(i), charge: () => true, release() {} }),
  }
  let ended = false
  const hooks: DriverHooks = { onSlideChange() {}, onVisit() {}, onCursor() {}, runLink() {}, onEnd() { ended = true } }
  const driver = createShowDriver({ source, stage, clock, sink, reducedMotion: false, kiosk: false, hooks })
  const draws = new Set<string>()
  /** Runs the show for `ms` at 60 fps, collecting what each frame drew. */
  const run = async (ms: number) => {
    for (let t = 0; t < ms; t += 16) { clock.advance(16); driver.tick(); for (const d of stage.inspect().lastDraws) draws.add(d); await tick() }
  }
  const pixel = (x: number, y: number) => Array.from(output.getContext('2d')!.getImageData(x, y, 1, 1).data.subarray(0, 3))
  return {
    driver, stage, sink, draws, run, pixel, morphAsked, ended: () => ended,
    async start() { await driver.start(s.visible[0]!); await tick(); await run(32) },
    async click() { driver.next(); await tick(); await tick() },
    dispose() { driver.dispose(); stage.dispose(); for (const w of windows) w.dispose(); root.remove() },
  }
}

const near = (a: number[], hex: string, tol = 12) => [0, 2, 4].every((k, i) => Math.abs(a[i]! - parseInt(hex.slice(k, k + 2), 16)) <= tol)

describe('T5 no-package prototype', () => {
  it('every package module is mocked to throw on import', async () => {
    for (const m of MOCKED) await expect(import(/* @vite-ignore */ m), m).rejects.toThrow()
  })

  it('a plan plays end to end with the package unreachable', async () => {
    // --- The fly-in deck: one slide, a box flying in on click 1, played from its layers. ---
    const fly = await session('fly')
    const flyPlan = fly.plans.get(0)?.plan
    expect(flyPlan, 'the server refused the fly-in plan: the prototype would play a still').toBeDefined()
    expect(flyPlan!.steps.some((s) => s.eligibility === 'layers')).toBe(true)
    const box = flyPlan!.layers.find((l) => isRasterLayer(l) && l.key === 'sp:3')
    expect(box).toBeDefined()
    expect(fly.plans.get(0)!.assets.has((box as { assetId: number }).assetId)).toBe(true)
    {
      const p = await play(fly)
      await p.start()
      expect(p.driver.slide).toBe(0)
      // Before click 1 the box is hidden: the centre is the white slide.
      expect(near(p.pixel(48, 36), 'FFFFFF')).toBe(true)
      let clicks = 0
      p.draws.clear()
      await p.click(); clicks++
      await p.run(150)
      // Mid-flight the box layer is drawn.
      expect([...p.draws]).toContain(`layer:${box!.id}`)
      await p.run(3000)
      expect(near(p.pixel(48, 36), '336699')).toBe(true)
      while (p.driver.slide !== 'end' && clicks < 10) { await p.click(); clicks++; await p.run(32) }
      expect(p.driver.slide).toBe('end')
      expect(clicks).toBe(flyPlan!.steps.length + 1)
      expect(p.sink.list().map((d) => d.code)).not.toContain('plan-stage-pressure')
      p.dispose()
    }

    // --- The Morph deck: a still slide, then a slide entered by a layered Morph. ---
    const mo = await session('morph')
    expect(mo.visible).toEqual([0, 1])
    expect(mo.stills.has(0)).toBe(true)
    const toPlan = mo.plans.get(1)?.plan
    expect(toPlan?.show.transition?.element).toBe('morph')
    const pair = mo.morphs.get(1)
    expect(pair).toBeDefined()
    expect(pair!.record.composited).toBeUndefined()
    expect(pair!.record.items.length).toBeGreaterThan(0)
    {
      const p = await play(mo)
      await p.start()
      expect(p.driver.slide).toBe(0)
      expect([...p.draws]).toContain('still')
      const perSlide = new Map<number, number>()
      let guard = 0
      while (p.driver.slide !== 'end' && guard++ < 20) {
        const at = p.driver.slide as number
        p.draws.clear()
        await p.click()
        perSlide.set(at, (perSlide.get(at) ?? 0) + 1)
        // The Morph runs on the show clock (well inside MORPH_WAIT_MS once its record resolves).
        await p.run(at === 0 ? 400 : 32)
        if (at === 0) {
          expect(p.driver.slide).toBe(1)
          expect([...p.draws]).toContain('morph')
          await p.run(3000)
        }
      }
      expect(p.driver.slide).toBe('end')
      expect(p.morphAsked).toEqual([`0→1@${pair!.fromStep}`])
      // The clicks per slide equal its record's steps + 1 (a still slide has no step).
      expect(perSlide.get(0)).toBe(0 + 1)
      expect(perSlide.get(1)).toBe(toPlan!.steps.length + 1)
      const codes = p.sink.list().map((d) => d.code as string)
      expect(codes).not.toContain('morph-late')
      expect(p.sink.list().filter((d) => (d as { effect?: string }).effect === 'morph-fallback')).toEqual([])
      p.dispose()
    }
  }, 60_000)
})
