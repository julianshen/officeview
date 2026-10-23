/**
 * Slideshow plan S1-22 fix round 1.
 *
 * 1. At the benchmark size (3840 × 2160) a slide transition plays: the outgoing frame is drawn from the
 *    transition allowance its slide's rung already priced, not charged to the workspace a second time.
 *    jsdom runs the Canvas 2D compositor, which takes no drawing-buffer lease; the WebGL2 one leases
 *    3840 × 2160 × 4 B for it, so the desktop workspace here is the 512 MiB budget less that lease.
 * 2. The session runs the slide appearance enrichment itself (spec 8): a show started on a freshly parsed
 *    document, before any viewer painted it, shows the master's background and artwork.
 *
 * Decks are "Office-shaped synthetic". Every clock is driven.
 */
import { afterEach, beforeEach, expect, it } from 'vitest'
import { sessionDeps } from '../src/slideshow/player/session'
import { createSlideShowWorkspace, SLIDESHOW_WORKSPACE_LIMIT } from '../src/slideshow/workspace'
import { OfficePackage } from '../src/core/zip'
import { parsePptx } from '../src/pptx/parse'
import { buildPptxDeck, type PptxShapeSpec } from '../src/testdata/ooxml-builders'
import { installShowSeams, loadDeck, startShow, type ShowSeams } from './helpers/slideshow-show-rig'

let seams: ShowSeams | undefined
const realWorkspace = sessionDeps.workspace
beforeEach(() => { seams = undefined })
afterEach(() => {
  sessionDeps.workspace = realWorkspace
  Object.defineProperty(window, 'devicePixelRatio', { value: 1, configurable: true, writable: true })
  seams?.restore()
  document.body.replaceChildren()
})

const W = 12192000, H = 6858000 // 16:9
const box = (id: number, x: number, y: number, w: number, h: number, fill: string): PptxShapeSpec =>
  ({ id, name: `Box ${id}`, off: [String(x), String(y)], ext: [String(w), String(h)], fill })

/** Two full-slide-background slides; the second splits in (a mask transition) over 1500 ms. */
async function twoSlides(profileBytes: number, size: { width: number; height: number }, dpr: number) {
  seams = installShowSeams(size)
  Object.defineProperty(window, 'devicePixelRatio', { value: dpr, configurable: true, writable: true })
  sessionDeps.workspace = (profile) => createSlideShowWorkspace(profile, { limit: profileBytes })
  const doc = await loadDeck(await buildPptxDeck({
    slideSize: [W, H],
    slides: [
      { shapes: [box(2, 0, 0, W, H, 'DDEEFF'), box(3, 914400, 914400, 3657600, 1828800, '336699')] },
      { shapes: [box(2, 0, 0, W, H, 'FFEEDD'), box(3, 6096000, 2743200, 3657600, 1828800, '993366')],
        transition: { xml: '<p:transition spd="slow"><p:split orient="vert" dir="out"/></p:transition>' } },
    ],
  }))
  return startShow(doc, { seams, size })
}

it('at 3840 × 2160 on the desktop budget a slide transition plays (two resident slides plus one transition fit)', async () => {
  const fb = 3840 * 2160 * 4
  const r = await twoSlides(SLIDESHOW_WORKSPACE_LIMIT.desktop - fb, { width: 1920, height: 1080 }, 2)
  try {
    expect(r.s.stage!.backingStore).toEqual({ w: 3840, h: 2160, dpr: 2 })
    await r.tick(100)
    // Both slides are resident at the full rung before the change.
    expect(r.s.stage!.inspect().slides.map(s => [s.index, s.level])).toEqual([[0, 'full'], [1, 'full']])
    r.s.next()
    await r.tick(0)
    expect(r.s.slideIndex).toBe(1)
    expect(r.diagnostics.filter(d => d.code === 'transition-workspace')).toEqual([])
    expect(r.s.stage!.transition).toEqual({ effect: 'split', duration: 1000 })
    await r.tick(500)
    expect(r.s.stage!.transition).toEqual({ effect: 'split', duration: 1000 })
    await r.tick(500)
    expect(r.s.stage!.transition).toBeUndefined()
  } finally { r.exit() }
})

it('the mobile budget (256 MiB) holds a transition at its largest output (a 1024 × 1366 tablet at DPR 2, capped at 2160 device px high)', async () => {
  // The backing store is capped at DPR 2 and 3840 × 2160 device px: 2048 × 2732 scales to 1619 × 2160.
  const fb = 1619 * 2160 * 4
  const r = await twoSlides(SLIDESHOW_WORKSPACE_LIMIT.mobile - fb, { width: 1024, height: 1366 }, 2)
  try {
    expect(r.s.stage!.backingStore).toMatchObject({ w: 1619, h: 2160 })
    await r.tick(100)
    expect(r.s.stage!.inspect().slides.map(s => [s.index, s.level])).toEqual([[0, 'full'], [1, 'full']])
    r.s.next()
    await r.tick(0)
    expect(r.s.slideIndex).toBe(1)
    expect(r.diagnostics.filter(d => d.code === 'transition-workspace')).toEqual([])
    expect(r.s.stage!.transition).toEqual({ effect: 'split', duration: 1000 })
  } finally { r.exit() }
})

it('a transition frame that genuinely cannot fit is still a diagnosed cut', async () => {
  // The show's first slide transitions from a black frame charged to the workspace (no slide rung holds it).
  // A workspace that refuses that one lease makes the change a cut, diagnosed, never silent.
  seams = installShowSeams({ width: 1920, height: 1080 })
  Object.defineProperty(window, 'devicePixelRatio', { value: 2, configurable: true, writable: true })
  sessionDeps.workspace = (profile) => {
    const ws = createSlideShowWorkspace(profile)
    return { ...ws, limit: ws.limit, lease: (bytes, tag) => (tag === 'transition:black' ? undefined : ws.lease(bytes, tag)), reserveTerminal: (b) => ws.reserveTerminal(b), used: () => ws.used() }
  }
  const doc = await loadDeck(await buildPptxDeck({
    slideSize: [W, H],
    slides: [{ shapes: [box(2, 0, 0, W, H, 'DDEEFF')], transition: { xml: '<p:transition spd="slow"><p:split orient="vert" dir="out"/></p:transition>' } }],
  }))
  const r = await startShow(doc, { seams, size: { width: 1920, height: 1080 } })
  try {
    expect(r.s.stage!.transition).toBeUndefined()
    expect(r.diagnostics.filter(d => d.code === 'transition-workspace').map(d => [d.slideIndex, d.message]))
      .toEqual([[0, 'The black source frame did not fit the slideshow workspace; the first slide is a cut']])
  } finally { r.exit() }
})

it('a show started on a freshly parsed document shows the master background and artwork (the session runs the enrichment)', async () => {
  seams = installShowSeams({ width: 1280, height: 720 })
  Object.defineProperty(window, 'devicePixelRatio', { value: 1, configurable: true, writable: true })
  // Never painted by a viewer: parsePptx only, so no slide carries an underlay yet.
  const doc = await parsePptx(await OfficePackage.load(await buildPptxDeck({
    slideSize: [W, H],
    masters: [{
      bg: { gradient: { stops: [[0, '00AA44'], [100000, '00AA44']], angle: 90 } },
      shapes: [box(10, 0, 0, W, 685800, 'AA0000')],
      layouts: [{ shapes: [] }],
    }],
    slides: [{ layout: [0, 0], shapes: [box(2, 4572000, 2743200, 3048000, 1371600, '3366CC')] }],
  })))
  expect(doc.slides[0].underlay).toBeUndefined()
  const r = await startShow(doc, { seams, size: { width: 1280, height: 720 } })
  try {
    await r.tick(0)
    const c = r.frameCanvas()!
    const px = (x: number, y: number) => Array.from(c.getContext('2d')!.getImageData(x, y, 1, 1).data)
    expect(px(640, 600), 'master gradient background').toEqual([0, 170, 68, 255])
    expect(px(640, 30), 'master artwork band').toEqual([170, 0, 0, 255])
    expect(px(640, 360), 'slide shape').toEqual([51, 102, 204, 255])
  } finally { r.exit() }
})
