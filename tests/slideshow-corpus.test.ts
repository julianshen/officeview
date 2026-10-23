/**
 * Slideshow plan S1-22: corpus conformance (spec 13 Stage 1 "corpus decks
 * present with correct build order and timing"; 6.4 diagnostics; the owner
 * rule that every effect except the "never" rows renders as its closest
 * equivalent).
 *
 * Real decks from LibreOffice core `sd/qa/unit/data/pptx` (MPL-2.0), fetched
 * by `bun scripts/fetch-corpus.ts` and never committed. The suite skips when
 * the corpus is absent; CI's check job fetches it first, so it runs there.
 *
 * The real lazy engine runs in jsdom with node-canvas (the Canvas 2D
 * compositor; jsdom has no WebGL2, diagnosed as `backend-fallback`). Every
 * clock is driven (tests/helpers/slideshow-show-rig.ts). The render-state
 * oracles are computed here from the file's own values (durations, tav
 * points, formulas, accel/decel), not read back from the engine.
 */
import { existsSync, readdirSync, readFileSync } from 'node:fs'
import { dirname, join, relative } from 'node:path'
import { fileURLToPath } from 'node:url'
import JSZip from 'jszip'
import { afterEach, beforeEach, describe, expect, it } from 'vitest'
import { rowForTransition } from '../src/slideshow/catalogue'
import { createDiagnosticSink, DIAGNOSTIC_CODES } from '../src/slideshow/diagnostics'
import { baseGeometry } from '../src/slideshow/player/stage'
import { readTransition } from '../src/slideshow/read/transition'
import { resolveTransition2D, transitionDraw } from '../src/slideshow/render/transitions2d'
import { dispatchTransition } from '../src/slideshow/render/three/framework'
import { createRng } from '../src/slideshow/rng'
import type { Affine, Compositor, TargetState, TransitionSpec } from '../src/slideshow/types'
import { coverageMarkdown, deckCoverage, playThrough, slideXmls } from './helpers/slideshow-coverage'
import { installShowSeams, loadDeck, startShow, type ShowRig, type ShowSeams } from './helpers/slideshow-show-rig'

const ROOT = join(dirname(fileURLToPath(import.meta.url)), '..')
const CORPUS = join(ROOT, 'corpus')
const deckPath = (name: string) => join(CORPUS, `lo-${name}.pptx`)
const HAVE = existsSync(deckPath('tdf104788'))
const bytesOf = (name: string) => new Uint8Array(readFileSync(deckPath(name)))

let seams: ShowSeams
beforeEach(() => { seams = installShowSeams() })
afterEach(() => { seams.restore(); document.body.replaceChildren() })

async function show(name: string): Promise<{ rig: ShowRig; doc: Awaited<ReturnType<typeof loadDeck>> }> {
  const doc = await loadDeck(bytesOf(name))
  return { rig: await startShow(doc, { seams }), doc }
}
const codes = (r: ShowRig, slide?: number) => r.diagnostics.filter(d => slide === undefined || d.slideIndex === slide).map(d => `${d.code}:${d.effect ?? ''}`)

/** The transition the slide's XML declares (1-based, show order), read with the engine's reader. */
async function transitionOf(name: string, slide: number): Promise<TransitionSpec> {
  const xml = (await slideXmls(bytesOf(name)))[slide - 1]
  const spec = xml && readTransition(xml, slide - 1, createDiagnosticSink())
  if (!spec) throw new Error(`${name} slide ${slide} has no transition`)
  return spec
}

const near = (a: Affine, b: Affine, digits = 6) => a.forEach((v, i) => expect(v).toBeCloseTo(b[i], digits))
const IDENTITY: Affine = [1, 0, 0, 1, 0, 0]
/** Settled: the entrance is complete and no offset remains. */
function expectSettled(s: TargetState | undefined) {
  expect(s?.visible).toBe(true)
  expect(s?.opacity).toBeCloseTo(1, 9)
  near(s!.transform, IDENTITY, 9)
  expect(s?.motion?.dx ?? 0).toBeCloseTo(0, 9)
  expect(s?.motion?.dy ?? 0).toBeCloseTo(0, 9)
}

/** SMIL time filter for accel `a` / decel `d` (fractions of the duration), as the spec's §6.2 semantics table cites it. */
function eased(p: number, a: number, d: number): number {
  const r = 1 / (1 - a / 2 - d / 2)
  if (p < a) return (r * p * p) / (2 * a)
  if (p <= 1 - d) return r * (p - a / 2)
  const t = 1 - p
  return 1 - (r * t * t) / (2 * d)
}

describe.skipIf(!HAVE)('slideshow corpus conformance (LibreOffice sd/qa/unit/data/pptx)', () => {
  it('tdf104788 plays six steps in file order with file durations', async () => {
    const { rig: r, doc } = await show('tdf104788')
    r.s.goTo(5)
    await r.tick(0)
    expect([r.s.slideIndex, r.s.buildIndex]).toEqual([5, 0])
    // 4:3, 9,144,000 × 6,858,000 EMU = 960 × 720 slide px.
    const W = 960, H = 720
    const g = baseGeometry(doc.slides[5])
    const geo = (k: string) => g.get(k)!
    // ppt_x / ppt_y are the target's centre as a fraction of the slide (ECMA-376 §19.5.4; LibreOffice reads X/Y as the
    // shape centre), so a "from (-#ppt_w/2)" fly-in starts with the shape just off the left edge.
    const xc = (k: string) => (geo(k).x + geo(k).w / 2) / W
    const yc = (k: string) => (geo(k).y + geo(k).h / 2) / H
    const steps: Array<{ key: string; span: number; check(s: TargetState): void }> = [
      // entr 37 rise up: fade 1000; ppt_y #ppt_y+1 → #ppt_y-.03 over 900 ms with decel 100 %. At 500 ms: p = 5/9.
      { key: 'sp:2', span: 1000, check(s) {
        expect(s.opacity).toBeCloseTo(0.5, 9)
        const y = 1 + (-0.03 - 1) * eased(5 / 9, 0, 1)
        expect(s.motion!.dy).toBeCloseTo(y * H, 6)
        expect(s.motion!.dx).toBeCloseTo(0, 9)
      } },
      // entr 34 breaks: ppt_x (-#ppt_w/2) → #ppt_x linear over 600 ms; the shear and squash start at 600 ms. At 500 ms: p = 5/6.
      { key: 'sp:34', span: 1000, check(s) {
        expect(s.opacity).toBe(1)
        const from = -geo('sp:34').w / W / 2, to = xc('sp:34')
        expect(s.motion!.dx).toBeCloseTo((from + (to - from) * (5 / 6) - to) * W, 6)
        near(s.transform, IDENTITY, 9)
      } },
      // entr 58 fold, accel 100 % on the effect: at 250 ms of 500 its children are at p = 0.5² = 0.25.
      // ppt_w #ppt_w*2.5 → #ppt_w; ppt_h #ppt_h*0.01 → #ppt_h; ppt_y #ppt_h+1 → #ppt_y; fade 500.
      { key: 'sp:30', span: 500, check(s) {
        const p = eased(0.5, 1, 0)
        expect(p).toBeCloseTo(0.25, 12)
        expect(s.opacity).toBeCloseTo(p, 9)
        expect(s.transform[0]).toBeCloseTo(2.5 + (1 - 2.5) * p, 9)
        expect(s.transform[3]).toBeCloseTo(0.01 + (1 - 0.01) * p, 9)
        const h = geo('sp:30').h / H
        expect(s.motion!.dy).toBeCloseTo(((h + 1) + (yc('sp:30') - (h + 1)) * p - yc('sp:30')) * H, 6)
      } },
      // entr 47 descend: fade 1000; ppt_y #ppt_y-.1 → #ppt_y linear 1000.
      { key: 'sp:39', span: 1000, check(s) {
        expect(s.opacity).toBeCloseTo(0.5, 9)
        expect(s.motion!.dy).toBeCloseTo(-0.1 * 0.5 * H, 6)
      } },
      // entr 42 ascend: fade 1000; ppt_y #ppt_y+.1 → #ppt_y linear 1000.
      { key: 'sp:5', span: 1000, check(s) {
        expect(s.opacity).toBeCloseTo(0.5, 9)
        expect(s.motion!.dy).toBeCloseTo(0.1 * 0.5 * H, 6)
      } },
      // entr 10 fade: 2000 ms.
      { key: 'sp:6', span: 2000, check(s) {
        expect(s.opacity).toBeCloseTo(0.5, 9)
        near(s.transform, IDENTITY, 9)
      } },
    ]
    for (const [k, step] of steps.entries()) {
      // Before its click, the step's target is hidden (an entrance), and so is every later one.
      for (const later of steps.slice(k)) expect(r.state().get(later.key)?.visible, `${later.key} before step ${k + 1}`).toBe(false)
      r.s.next()
      await r.tick(0)
      expect(r.s.buildIndex).toBe(k + 1)
      await r.tick(step.span / 2)
      const mid = r.state().get(step.key)!
      expect(mid.visible).toBe(true)
      step.check(mid)
      // The file duration: still moving 1 ms before it, settled at it.
      await r.tick(step.span / 2 - 1)
      const before = r.state().get(step.key)!
      expect(Math.abs(1 - before.opacity) + Math.abs(before.motion?.dx ?? 0) + Math.abs(before.motion?.dy ?? 0) + Math.abs(before.transform[0] - 1)).toBeGreaterThan(1e-4)
      await r.tick(1)
      expectSettled(r.state().get(step.key))
    }
    // The breaks effect's xshear has no render mapping: diagnosed once, on its slide.
    expect(codes(r, 5)).toContain('behaviour-unsupported:behaviour 16')
    r.exit()
  })

  it('tdf44223 split transition, bldP animBg and entrance preset 2', async () => {
    const { rig: r, doc } = await show('tdf44223')
    // bldP animBg="1": the heart (spid 4) builds as one target with its background, not split into paragraphs.
    const plan = r.s.stage!.plan!
    const heart = plan.layers.filter(l => l.kind === 'target' && l.key.startsWith('sp:4'))
    expect(heart.map(l => l.kind === 'target' && [l.key, l.part ?? 'whole'])).toEqual([['sp:4', 'whole']])
    expect(r.state().get('sp:4')?.visible).toBe(false)
    // Entrance preset 2 (fly in), subtype 4 (from bottom): ppt_y 1+#ppt_h/2 → #ppt_y over 500 ms, ppt_x unchanged.
    r.s.next()
    await r.tick(0)
    await r.tick(250)
    const W = 1280, H = 720 // 16:9, 12,192,000 × 6,858,000 EMU
    expect(doc.slides[0].widthEmu / 9525).toBe(W)
    const g = baseGeometry(doc.slides[0]).get('sp:4')!
    const h = g.h / H, yc = (g.y + g.h / 2) / H
    const mid = r.state().get('sp:4')!
    expect(mid.visible).toBe(true)
    expect(mid.motion!.dy).toBeCloseTo(((1 + h / 2) + (yc - (1 + h / 2)) * 0.5 - yc) * H, 6)
    expect(mid.motion!.dx).toBeCloseTo(0, 9)
    // The fly-in's sound (p:audio on a p:sndTgt, in the effect's subTnLst) reaches the media player with the effect
    // (S2-10; catalogue row sound:effect). jsdom decodes no audio (canPlayType is ''), so here it runs silent with
    // media-codec on the effect's behaviour, diagnosed; the next test plays it.
    expect(r.diagnostics.some(d => d.code === 'media-codec' && d.slideIndex === 0 && d.effect?.startsWith('behaviour ') && d.message.includes('audio/x-wav'))).toBe(true)
    expect(codes(r, 0).filter(c => c.startsWith('media-stand-in'))).toEqual([])
    await r.tick(250)
    expectSettled(r.state().get('sp:4'))
    // Slide 2 splits in (orient vert, spd slow, p14:dur 1500): native, not a look-alike.
    r.s.next()
    await r.tick(0)
    expect(r.s.slideIndex).toBe(1)
    expect(r.s.stage!.transition).toEqual({ effect: 'split', duration: 1500 })
    const spec = await transitionOf('tdf44223', 2)
    expect(rowForTransition(spec.ns, spec.element)?.status).toBe('exact')
    expect(resolveTransition2D(spec).approximated).toBe(false)
    expect(transitionDraw(spec, 0.5, createRng(1)).mask).toEqual({ kind: 'split', progress: 0.5, params: { orient: 'vert', dir: 'out' } })
    await r.tick(750)
    expect(r.s.stage!.transition).toEqual({ effect: 'split', duration: 1500 })
    expect(codes(r, 1)).not.toContain('transition-approximate:split')
    // Its applause sound goes to the transition-sound controller (S2-7). jsdom decodes no audio (canPlayType is ''),
    // so here the split runs silent with media-codec, diagnosed once for the slide; a browser plays the WAV.
    expect(codes(r, 1)).toContain('media-codec:transition-sound')
    expect(codes(r, 1).filter(c => c.startsWith('media-stand-in'))).toEqual([])
    await r.tick(750)
    expect(r.s.stage!.transition).toBeUndefined()
    r.exit()
  })

  it('tdf44223 fly-in plays its wind.wav effect sound when the effect begins (S2-10)', async () => {
    // A browser that decodes WAV: canPlayType and play() stand in for the decoder; the bytes are the file's own part,
    // read through the S2-6 lease (a Blob URL), so the record, lease and timing paths are the real ones.
    const proto = HTMLMediaElement.prototype as unknown as { canPlayType: unknown; play: unknown; pause: unknown; load: unknown }
    const saved = { canPlayType: proto.canPlayType, play: proto.play, pause: proto.pause, load: proto.load }
    const U = URL as unknown as { createObjectURL?: unknown; revokeObjectURL?: unknown }
    const savedUrl = { createObjectURL: U.createObjectURL, revokeObjectURL: U.revokeObjectURL }
    const played: string[] = []
    const blobs: Blob[] = []
    proto.canPlayType = function () { return 'maybe' }
    proto.play = function (this: HTMLMediaElement) { played.push(this.getAttribute('data-officeview-show-effect-sound') ?? 'other'); return Promise.resolve() }
    proto.pause = function () {}
    proto.load = function () {}
    U.createObjectURL = (b: Blob) => { blobs.push(b); return `blob:effect-${blobs.length}` }
    U.revokeObjectURL = () => {}
    try {
      const { rig: r } = await show('tdf44223')
      await r.tick(0)
      // Nothing plays before the click that begins the fly-in.
      expect(played).toEqual([])
      r.s.next()
      for (let i = 0; i < 5; i++) await r.tick(16)
      const el = document.querySelector<HTMLAudioElement>('audio[data-officeview-show-effect-sound]')
      expect(el).not.toBeNull()
      expect(el!.hidden).toBe(true)
      expect(el!.getAttribute('src')).toMatch(/^blob:effect-/)
      expect(played.filter(p => p !== 'other')).toHaveLength(1)
      // The Blob is the file's audio1.wav (27.4 KiB is audio2, the transition's applause), as audio.
      const part = await (await JSZip.loadAsync(bytesOf('tdf44223'))).file('ppt/media/audio1.wav')!.async('uint8array')
      expect(blobs.map(b => b.size)).toContain(part.length)
      expect(codes(r, 0).filter(c => c.startsWith('media-'))).toEqual([])
      r.exit()
    } finally {
      Object.assign(proto, saved)
      Object.assign(U, savedUrl)
    }
  })

  it('tdf142915 auto-advances after 1000 ms and ignores clicks', async () => {
    // The deck's showPr says useTimings="0" ("Use Timings" off in Set Up Show), and the show honours it: advTm never
    // fires. That half is Office evidence about tdf142915. The second half changes that one attribute in memory
    // (useTimings="1"; the committed corpus file is never modified): it is engine evidence that advTm fires at
    // 1000 ms after each slide's transition when timings are on, not a claim about the deck as filed.
    // advClick="0" ignores clicks either way.
    const zip = await JSZip.loadAsync(bytesOf('tdf142915'))
    const props = await zip.file('ppt/presProps.xml')!.async('string')
    expect(props).toContain('useTimings="0"')
    const asFiled = await loadDeck(bytesOf('tdf142915'))
    let r = await startShow(asFiled, { seams })
    await r.tick(5000)
    await r.click(640, 360)
    expect(r.s.slideIndex).toBe(0)
    await r.key('ArrowRight') // keys are never gated by advClick
    expect(r.s.slideIndex).toBe(1)
    r.exit()
    document.body.replaceChildren()

    zip.file('ppt/presProps.xml', props.replace('useTimings="0"', 'useTimings="1"'))
    const timed = await loadDeck(await zip.generateAsync({ type: 'uint8array' }))
    r = await startShow(timed, { seams })
    const t0 = seams.clock.now()
    expect(r.s.slideIndex).toBe(0)
    // Slide 1: p14:dur="10" with no effect element (p:none) is drawn as a cut, so its slide time starts at once.
    await r.tick(500)
    await r.click(640, 360)
    expect(r.s.slideIndex).toBe(0)
    await r.tick(t0 + 999 - seams.clock.now())
    expect(r.s.slideIndex).toBe(0)
    await r.tick(1)
    expect(r.s.slideIndex).toBe(1)
    await r.click(640, 360)
    expect(r.s.slideIndex).toBe(1)
    // Slide 2: p14:dur="0", advTm 1000, then the end-of-show slide.
    await r.tick(999)
    expect(r.s.state.slide).toBe(1)
    await r.tick(1)
    expect(r.s.state.slide).toBe('end')
    r.exit()
  })

  it('bnc591147 trigger runs its interactive sequence; mediacall commands reach the media player; the avi shows its poster', async () => {
    const { rig: r } = await show('bnc591147')
    // The video (video.avi, a 10-byte stand-in file) shows its poster picture (image1.png): the frame is painted
    // with the poster's pixels, and no image was withheld.
    const frame = r.frameCanvas()!
    expect(frame).toBeDefined()
    const ctx = frame.getContext('2d')!
    const px = ctx.getImageData(Math.floor(frame.width / 2), Math.floor(frame.height / 2), 1, 1).data
    const white = px[0] === 255 && px[1] === 255 && px[2] === 255
    expect(white, 'poster pixels at the slide centre').toBe(false)
    expect(px[3]).toBe(255)
    expect(r.diagnostics.map(d => d.code)).not.toEqual(expect.arrayContaining(['image-workspace']))
    expect(r.diagnostics.map(d => d.code)).not.toEqual(expect.arrayContaining(['scratch-image']))
    // The slide's playFrom(0.0) mediacall starts with the slide and reaches the media player (S2-7), which refuses the
    // undecodable AVI (application/octet-stream) on its shape: the poster stays, diagnosed as media-codec.
    await r.tick(10)
    expect(codes(r, 0)).toContain('media-codec:4')
    expect(codes(r, 0).filter(c => c.startsWith('media-stand-in'))).toEqual([])
    expect(codes(r, 0)).not.toContain('behaviour-unsupported:behaviour 6')
    // A click on the poster (spid 4, also a refused ppaction://media link) runs its interactive sequence (togglePause).
    await r.click(480, 360)
    await r.tick(10)
    expect(r.s.cursor!.events.some(e => e.kind === 'trigger' && e.spid === '4')).toBe(true)
    expect(r.s.slideIndex).toBe(0)
    // togglePause on an unplayable node is a no-op (spec 6.5): the codec diagnostic stands, nothing new is said.
    expect(codes(r, 0).filter(c => c.startsWith('media-'))).toEqual(['media-codec:4'])
    expect(codes(r, 0)).toContain('link-refused:4::')
    r.exit()
  })

  it('airplane and morph decks play as Stage 3 rows: airplane is a shader on WebGL2 and its look-alike on Canvas 2D; Morph takes the morph route on both', async () => {
    // jsdom has no WebGL2, so the show runs on Canvas 2D; the WebGL2 half is the dispatch below (and, on a real
    // GPU, tests/browser/slideshow-corpus.spec.ts, which plays both decks on WebGL2).
    const webgl2 = { backend: 'webgl2', prepareTransition3D: () => true } as unknown as Compositor
    const canvas2d = { backend: 'canvas2d' } as unknown as Compositor
    for (const [name, element, duration, preset] of [['airplane-transition', 'prstTrans', 1250, 'airplane'], ['tdf149588_transparentSolidFill', 'morph', 1500, undefined]] as const) {
      const { rig: r } = await show(name)
      const spec = await transitionOf(name, 1)
      expect([spec.element, spec.preset]).toEqual([element, preset])
      const row = rowForTransition(spec.ns, spec.element, spec.preset)!
      expect(row).toMatchObject({ status: 'approximate', implementation: 'shader', stage: 3, verified: false, lookalike: 'transition:p:fade' })
      if (element === 'prstTrans') {
        // The show's backend registered every 3-D row (three/index.ts): WebGL2 runs airplane's shader.
        expect(dispatchTransition(spec, webgl2)).toBe('shader')
        // Canvas 2D draws the fade look-alike, diagnosed once, for the file's duration.
        expect(dispatchTransition(spec, canvas2d)).toBe('lookalike')
        expect(r.s.stage!.transition).toEqual({ effect: 'fade', duration })
        expect(codes(r, 0)).toContain('transition-approximate:prstTrans:airplane')
      } else {
        // Morph takes the morph route on both backends, diagnosed as an approximation.
        expect(dispatchTransition(spec, webgl2)).toBe('morph')
        expect(dispatchTransition(spec, canvas2d)).toBe('morph')
        expect(codes(r, 0)).toContain('transition-approximate:morph')
        // The deck's only slide carries the Morph, and a first slide enters from black, as PowerPoint does: nothing
        // to morph from, so the fade look-alike plays and says why (S3-6).
        expect(codes(r, 0)).toContain('transition-approximate:morph-fallback')
        expect(r.diagnostics.find(d => d.effect === 'morph-fallback')?.message).toContain('the first slide enters from black')
        expect(r.s.stage!.transition).toEqual({ effect: 'fade', duration })
      }
      expect(codes(r, 0).filter(c => c.startsWith('transition-shader-fallback'))).toEqual([])
      await r.tick(duration / 2)
      expect(r.s.stage!.transition).toEqual({ effect: 'fade', duration })
      await r.tick(duration / 2)
      expect(r.s.stage!.transition).toBeUndefined()
      r.exit()
      document.body.replaceChildren()
    }
  })

  it('tdf156808 plays the flythrough effect its catalogue row names (2-D, slow, 800 ms), not a stand-in', async () => {
    const { rig: r } = await show('tdf156808')
    const spec = await transitionOf('tdf156808', 1)
    expect([spec.ns, spec.element, spec.attrs.dir, spec.durationMs]).toEqual(['p14', 'flythrough', 'out', 800])
    const row = rowForTransition(spec.ns, spec.element)!
    expect([row.id, row.status, row.implementation, row.lookalike]).toEqual(['transition:p14:flythrough', 'exact', 'transform', undefined])
    expect(resolveTransition2D(spec)).toEqual({ spec, approximated: false })
    expect(r.s.stage!.transition).toEqual({ effect: 'flythrough', duration: 800 })
    await r.tick(400)
    expect(r.s.stage!.transition).toEqual({ effect: 'flythrough', duration: 800 })
    // At 50 %, fly through "out": the incoming slide is at 1.5 − 0.5·p = 1.25 × about the centre, half opaque.
    const at = transitionDraw(spec, 0.5, createRng(1))
    near(at.to!, [1.25, 0, 0, 1.25, -0.125, -0.125], 12)
    expect(at.toOpacity).toBeCloseTo(0.5, 12)
    expect(at.mask).toBeUndefined()
    expect(codes(r).filter(c => c.startsWith('transition-'))).toEqual([])
    await r.tick(400)
    expect(r.s.stage!.transition).toBeUndefined()
    r.exit()
  })

  it('every corpus deck presents without throwing and every diagnostic code is catalogued', async () => {
    const decks = readdirSync(CORPUS).filter(f => f.endsWith('.pptx')).sort()
    expect(decks.length).toBeGreaterThanOrEqual(16)
    const emitted = new Set<string>()
    const table: Array<{ deck: string; items: Awaited<ReturnType<typeof deckCoverage>> }> = []
    for (const file of decks) {
      const bytes = new Uint8Array(readFileSync(join(CORPUS, file)))
      const doc = await loadDeck(bytes)
      const r = await startShow(doc, { seams })
      await playThrough(r)
      expect(r.s.state.slide, file).toBe('end')
      for (const d of r.diagnostics) emitted.add(d.code)
      if (file.startsWith('lo-')) table.push({ deck: file.slice(3, -5), items: await deckCoverage(bytes, doc, r.diagnostics) })
      r.exit()
      document.body.replaceChildren()
    }
    // Every emitted code is in the diagnostic catalogue.
    expect([...emitted].filter(c => !(c in DIAGNOSTIC_CODES))).toEqual([])
    // So is every code literal in the slideshow sources (a new code cannot ship uncatalogued).
    const literal = /(?:code: ?|diagnose\((?:i, )?|add\(ctx, |report\(|unsupported\([^)]*?, )'([a-z0-9]+(?:-[a-z0-9]+)+)'/g
    const missing: string[] = []
    const walk = (dir: string) => {
      for (const e of readdirSync(dir, { withFileTypes: true })) {
        const p = join(dir, e.name)
        if (e.isDirectory()) walk(p)
        else if (/\.tsx?$/.test(e.name)) for (const m of readFileSync(p, 'utf8').matchAll(literal)) if (!(m[1] in DIAGNOSTIC_CODES)) missing.push(`${relative(ROOT, p)}: ${m[1]}`)
      }
    }
    walk(join(ROOT, 'src/slideshow'))
    expect(missing).toEqual([])
    // Owner rule: every effect except a "never" row plays as itself or its closest equivalent, and says so.
    const items = table.flatMap(t => t.items.map(it => ({ deck: t.deck, ...it })))
    expect(items.filter(it => it.coverage === 'refused')).toEqual([])
    expect(items.filter(it => /NOT diagnosed/.test(it.note))).toEqual([])
    expect(coverageMarkdown(table)).toContain('| tdf104788 |')
  }, 120_000)
})
