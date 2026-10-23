/**
 * Slideshow PR #18 review 4: the compositor's output resize is part of the resize transaction. When the
 * compositor keeps its old drawing buffer (WebGL2 under workspace pressure), the stage keeps the old backing
 * store, layer scale and viewport, refunds the new terminal reserve, and diagnoses.
 *
 * The backend manager is wrapped (its own file: `vi.mock` is per module) so its compositor refuses any
 * output wider than `refuse.above`, as `webgl2.resizeOutput` does when the new framebuffer lease fails.
 * jsdom with node-canvas; every clock is driven. Decks are "Office-shaped synthetic".
 */
import { afterEach, expect, it, vi } from 'vitest'
import { sessionDeps } from '../src/slideshow/player/session'
import { createSlideShowWorkspace } from '../src/slideshow/workspace'
import type { SlideShowWorkspace } from '../src/slideshow/types'
import { buildPptxDeck } from '../src/testdata/ooxml-builders'
import { installShowSeams, loadDeck, startShow, type ShowSeams } from './helpers/slideshow-show-rig'

const refuse: { above?: number } = {}
vi.mock('../src/slideshow/render/backend', async (importOriginal) => {
  const real = await importOriginal<typeof import('../src/slideshow/render/backend')>()
  return {
    ...real,
    createBackendManager: (...args: Parameters<typeof real.createBackendManager>) => {
      const m = real.createBackendManager(...args)
      const c = m.current
      const setViewport = c.setViewport.bind(c)
      c.setViewport = (v) => (refuse.above !== undefined && v.deviceWidth > refuse.above ? false : setViewport(v))
      return m
    },
  }
})

let seams: ShowSeams | undefined
const realWorkspace = sessionDeps.workspace
afterEach(() => {
  sessionDeps.workspace = realWorkspace
  refuse.above = undefined
  seams?.restore()
  seams = undefined
  document.body.replaceChildren()
})

it('a refused compositor resize keeps the old backing store, layer scale and viewport, and refunds the new reserve', async () => {
  seams = installShowSeams({ width: 1280, height: 720 })
  let ws!: SlideShowWorkspace
  sessionDeps.workspace = (profile) => (ws = createSlideShowWorkspace(profile))
  const doc = await loadDeck(await buildPptxDeck({
    slideSize: [12192000, 6858000],
    slides: [{ shapes: [{ id: 2, name: 'Box', off: ['914400', '914400'], ext: ['3657600', '1828800'], fill: '336699', paragraphs: [{ runs: [{ text: 'A' }] }] }] }],
  }))
  const r = await startShow(doc, { seams, size: { width: 1280, height: 720 } })
  try {
    await r.tick(100)
    const before = { backing: r.s.stage!.backingStore, scale: r.s.stage!.inspect().slides[0].scale, used: ws.used() }
    expect(before.backing).toEqual({ w: 1280, h: 720, dpr: 1 })
    const output = r.host.querySelector<HTMLCanvasElement>('[data-officeview-show-canvas]')!
    refuse.above = 1280
    Object.defineProperty(window, 'innerWidth', { value: 1920, configurable: true, writable: true })
    Object.defineProperty(window, 'innerHeight', { value: 1080, configurable: true, writable: true })
    window.dispatchEvent(new Event('resize'))
    await r.tick(0)
    expect(r.s.stage!.backingStore).toEqual(before.backing)
    expect([output.width, output.height]).toEqual([1280, 720])
    expect(r.s.stage!.inspect().slides[0].scale).toBe(before.scale)
    expect(ws.used()).toBe(before.used)
    expect(r.diagnostics.filter(d => d.code === 'resize-kept-old').length).toBe(1)
    // Once the compositor can take it, the next resize goes through.
    refuse.above = undefined
    window.dispatchEvent(new Event('resize'))
    await r.tick(0)
    expect(r.s.stage!.backingStore).toEqual({ w: 1920, h: 1080, dpr: 1 })
  } finally { r.exit() }
})
