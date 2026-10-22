/**
 * Slideshow PR #18 review.
 *
 * P1. A slide whose preparation takes show time (image decoding, layer construction) starts its transition
 *     when it becomes drawable, not when the change was requested: the transition plays in full on screen,
 *     and its `advTm` counts from the transition's end measured from there.
 *
 * P2. The Canvas 2D fallback (WebGL2 unavailable, as in jsdom) composites on the terminal output canvas,
 *     which the permanent reserve already pays for: the stage mounts no second full-size canvas.
 *
 * The residency is wrapped so slide 1's `ensure` waits on a gate the test opens (its own file: `vi.mock`
 * is per module). jsdom with node-canvas; every clock is driven. Decks are "Office-shaped synthetic".
 */
import { afterEach, expect, it, vi } from 'vitest'
import { buildPptxDeck, type PptxShapeSpec } from '../src/testdata/ooxml-builders'
import { installShowSeams, loadDeck, flush, startShow, type ShowSeams } from './helpers/slideshow-show-rig'

const gate: { open?: Promise<void>; release?: () => void } = {}
vi.mock('../src/slideshow/residency', async (importOriginal) => {
  const real = await importOriginal<typeof import('../src/slideshow/residency')>()
  return {
    ...real,
    createResidency: (...args: Parameters<typeof real.createResidency>) => {
      const r = real.createResidency(...args)
      const ensure = r.ensure.bind(r)
      r.ensure = async (slide, level) => {
        if (slide === 1 && gate.open) await gate.open
        return ensure(slide, level)
      }
      return r
    },
  }
})

/** Slide metadata reads for `metaGate.slide` wait on `metaGate.open` (a slow slide to load). */
const metaGate: { slide?: number; open?: Promise<void> } = {}
vi.mock('../src/slideshow/read/store', async (importOriginal) => {
  const real = await importOriginal<typeof import('../src/slideshow/read/store')>()
  return {
    ...real,
    createSlideMetadataStore: (...args: Parameters<typeof real.createSlideMetadataStore>) => {
      const store = real.createSlideMetadataStore(...args)
      const acquire = store.acquire.bind(store)
      store.acquire = async (i, signal) => {
        if (i === metaGate.slide && metaGate.open) await metaGate.open
        return acquire(i, signal)
      }
      return store
    },
  }
})

let seams: ShowSeams | undefined
afterEach(() => {
  seams?.restore()
  seams = undefined
  gate.open = gate.release = undefined
  metaGate.slide = metaGate.open = undefined
  document.body.replaceChildren()
})

const W = 12192000, H = 6858000
const box = (id: number, fill: string): PptxShapeSpec => ({ id, name: `Box ${id}`, off: ['0', '0'], ext: [String(W), String(H)], fill })

it('P1: a slow-to-prepare slide starts its transition, and its advTm, when it becomes drawable', async () => {
  seams = installShowSeams({ width: 1280, height: 720 })
  gate.open = new Promise<void>((resolve) => { gate.release = resolve })
  const doc = await loadDeck(await buildPptxDeck({
    slideSize: [W, H],
    slides: [
      { shapes: [box(2, 'DDEEFF')] },
      // A 1000 ms fade, then advance 2000 ms after it ends.
      { shapes: [box(2, 'FFEEDD')], transition: { xml: '<p:transition spd="slow" advTm="2000"><p:fade/></p:transition>' } },
      { shapes: [box(2, 'EEFFDD')] },
    ],
  }))
  const r = await startShow(doc, { seams, size: { width: 1280, height: 720 } })
  try {
    await r.tick(100)
    r.s.next()
    await r.tick(0)
    expect(r.s.slideIndex).toBe(1)
    // Slide 1 is still preparing: 5000 ms of show time pass before it can be drawn.
    await r.tick(5000)
    expect(r.s.stage!.transition).toBeUndefined()
    gate.release!()
    await flush()
    const drawable = seams.clock.now()
    // The transition starts now, at progress 0: the slide's timeline begins one duration later.
    expect(r.s.stage!.slideTime(drawable)).toBe(-1000)
    r.s.tick()
    expect(r.s.stage!.transition).toEqual({ effect: 'fade', duration: 1000 })
    await r.tick(999)
    expect(r.s.stage!.transition).toEqual({ effect: 'fade', duration: 1000 })
    await r.tick(1)
    expect(r.s.stage!.transition).toBeUndefined()
    // advTm counts from the transition's end: 2000 ms after it, not before.
    await r.tick(1999)
    expect(r.s.slideIndex).toBe(1)
    await r.tick(1)
    expect(r.s.slideIndex).toBe(2)
  } finally { r.exit() }
})

it('P2: a Canvas 2D show composites on the reserved output canvas and mounts no second full-size canvas', async () => {
  seams = installShowSeams({ width: 1280, height: 720 })
  const doc = await loadDeck(await buildPptxDeck({
    slideSize: [W, H],
    slides: [{ shapes: [box(2, 'DDEEFF')] }, { shapes: [box(2, 'FFEEDD')] }],
  }))
  const r = await startShow(doc, { seams, size: { width: 1280, height: 720 } })
  try {
    await r.tick(100)
    expect(r.s.stage!.backend).toBe('canvas2d')
    expect(r.s.stage!.level).toBe('full')
    const output = r.host.querySelector<HTMLCanvasElement>('[data-officeview-show-canvas]')!
    expect(r.host.querySelectorAll('[data-officeview-show-stage] canvas')).toHaveLength(0)
    expect(r.host.querySelectorAll('[data-officeview-show-surface] canvas')).toHaveLength(1)
    expect(r.frameCanvas()).toBe(output)
    // The composited slide is on the output: its background colour at the centre.
    const px = output.getContext('2d')!.getImageData(output.width >> 1, output.height >> 1, 1, 1).data
    expect([...px.slice(0, 3)]).toEqual([0xDD, 0xEE, 0xFF])
    r.s.next()
    await r.tick(100)
    const px2 = output.getContext('2d')!.getImageData(output.width >> 1, output.height >> 1, 1, 1).data
    expect([...px2.slice(0, 3)]).toEqual([0xFF, 0xEE, 0xDD])
  } finally { r.exit() }
})

it('PR #18 review 3: navigation queued on a loading slide survives the slide change it causes', async () => {
  seams = installShowSeams({ width: 1280, height: 720 })
  const doc = await loadDeck(await buildPptxDeck({
    slideSize: [W, H],
    slides: [{ shapes: [box(2, 'DDEEFF')] }, { shapes: [box(2, 'FFEEDD')] }, { shapes: [box(2, 'EEFFDD')] }, { shapes: [box(2, 'DDDDDD')] }],
  }))
  let release!: () => void
  metaGate.slide = 1
  metaGate.open = new Promise<void>((resolve) => { release = resolve })
  const r = await startShow(doc, { seams, size: { width: 1280, height: 720 } })
  try {
    await r.tick(100)
    // Three fast Next presses: the first enters slide 1 (still loading), the other two queue there.
    r.s.next(); r.s.next(); r.s.next()
    expect(r.s.slideIndex).toBe(1)
    release()
    await r.tick(0)
    await r.tick(0)
    expect(r.s.slideIndex).toBe(3)
  } finally { r.exit() }
})
