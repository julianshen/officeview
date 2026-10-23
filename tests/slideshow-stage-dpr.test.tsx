/**
 * Slideshow plan S1-20 review M2: a resize that keeps the backing store's
 * pixel size but changes its DPR must re-rasterise a reduced-rung slide,
 * whose layer scale follows the DPR.
 *
 * The residency is wrapped so every slide settles on the reduced rung
 * (its own file: `vi.mock` is per module). jsdom with node-canvas; every
 * clock is driven. Decks are "Office-shaped synthetic".
 */
import { afterEach, beforeEach, expect, it, vi } from 'vitest'
import { runShow } from '../src/slideshow'
import { sessionDeps } from '../src/slideshow/player/session'
import { ManualClock } from '../src/slideshow/clock'
import type { ShowStart } from '../src/components/present'
import { buildPptxDeck } from '../src/testdata/ooxml-builders'
import { OfficePackage } from '../src/core/zip'
import { parsePptx, pptxPackageHandle } from '../src/pptx/parse'

vi.mock('../src/slideshow/residency', async (importOriginal) => {
  const real = await importOriginal<typeof import('../src/slideshow/residency')>()
  return {
    ...real,
    createResidency: (...args: Parameters<typeof real.createResidency>) => {
      const r = real.createResidency(...args)
      const ensure = r.ensure.bind(r)
      r.ensure = (slide, level) => ensure(slide, level === 'full' ? 'reduced' : level)
      return r
    },
  }
})

const clock = new ManualClock()
let size = { w: 1920, h: 1440, dpr: 1.5 }
beforeEach(() => {
  vi.spyOn(sessionDeps, 'clock').mockImplementation(() => clock)
  vi.spyOn(sessionDeps, 'frames').mockImplementation(() => () => {})
  vi.spyOn(window, 'innerWidth', 'get').mockImplementation(() => size.w)
  vi.spyOn(window, 'innerHeight', 'get').mockImplementation(() => size.h)
  vi.spyOn(window, 'devicePixelRatio', 'get').mockImplementation(() => size.dpr)
})
afterEach(() => {
  vi.restoreAllMocks()
  document.body.replaceChildren()
})
const flush = async () => { for (let i = 0; i < 20; i++) await new Promise((r) => setTimeout(r, 0)) }

it('M2: a DPR-only change of the backing store re-rasterises the reduced rung at the new scale', async () => {
  const doc = await parsePptx(await OfficePackage.load(await buildPptxDeck({
    slides: [{ shapes: [{ id: 2, off: ['0', '0'], ext: ['914400', '914400'], fill: '336699', paragraphs: [{ runs: [{ text: 'A' }] }] }] }],
  })))
  const host = document.createElement('div')
  host.setAttribute('data-officeview-show', '')
  document.body.append(host)
  const ac = new AbortController()
  const s = await runShow({
    doc, handle: pptxPackageHandle(doc), host,
    token: { release() {}, released: false, owns: () => false },
    platform: { hasActivation: () => true, requestFullscreen: async () => {}, exitFullscreen: async () => {}, fullscreenElement: () => null, open: () => null },
    fullscreen: Promise.resolve('entered'), retryFullscreen: async () => 'entered',
    from: 'beginning', motion: 'auto',
    callbacks: { currentSlide: () => 0, restoreFocus() {}, scrollToSlide() {} },
    signal: ac.signal, generation: 1, isCurrent: () => !ac.signal.aborted, cleanup: () => ac.abort(),
  } as ShowStart)
  await flush()
  // 1920 x 1440 CSS at DPR 1.5: a 2880 x 2160 backing store (capped by 2160 rows); reduced layers at DPR 1.
  expect(s.stage!.backingStore).toEqual({ w: 2880, h: 2160, dpr: 1.5 })
  const before = s.stage!.inspect().slides[0]
  expect(before.level).toBe('reduced')
  expect(before.scale).toBeCloseTo(2, 5)
  // 2880 x 2160 CSS at DPR 1: the same pixel size, a different DPR
  size = { w: 2880, h: 2160, dpr: 1 }
  window.dispatchEvent(new Event('resize'))
  await flush()
  expect(s.stage!.backingStore).toEqual({ w: 2880, h: 2160, dpr: 1 })
  for (let i = 0; i < 5; i++) s.tick()
  const after = s.stage!.inspect().slides[0]
  expect(after.level).toBe('reduced')
  expect(after.scale).toBeCloseTo(3, 5)
  expect(after.pending).toBe(0)
  s.exit()
})
