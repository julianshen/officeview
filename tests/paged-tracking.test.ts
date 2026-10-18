import { describe, it, expect, vi, beforeEach, afterEach } from 'vitest'
import { createElement, type RefObject } from 'react'
import { render, renderHook, waitFor, fireEvent, act, within } from '@testing-library/react'
import { docxModel, mockViewerLayout, type FakePage } from './helpers/paged-viewer'
import { usePageTracking } from '../src/components/paged/usePageTracking'
import { clampPan, type Transform } from '../src/core/zoom'
import { contentYAt, enterZoom, exitZoom, pageOffsets, panYForContentTop, remapContentY, type ScrollPanState } from '../src/components/scrollPan'
import { mostVisiblePage } from '../src/components/paged/pageTracking'
import type { PageSpec } from '../src/render/paint'

const h = vi.hoisted(() => ({ pages: null as FakePage[] | null }))
vi.mock('../src/render/paint', async (importOriginal) => {
  const m = await importOriginal<typeof import('../src/render/paint')>()
  const { fakePaintables: fake } = await import('./helpers/paged-viewer')
  return { ...m, getPaintables: (doc: never, opts: never) => (h.pages ? Promise.resolve(fake(h.pages)) : m.getPaintables(doc, opts)) }
})

const spec = (widthPx: number, heightPx: number): PageSpec => ({ widthPx, heightPx }) as PageSpec

describe('page tracking (pure)', () => {
  it('two 2000px pages in an 800px viewport switch at the dominance crossover', () => {
    const o = pageOffsets([spec(1000, 2000), spec(1000, 2000)], 1000, 0)
    expect(o).toEqual({ top: [0, 2000], height: [2000, 2000] })
    // page 0 visible 600, page 1 visible 200
    expect(mostVisiblePage(o, 1400, 800)).toBe(0)
    // page 0 visible 200, page 1 visible 600
    expect(mostVisiblePage(o, 1800, 800)).toBe(1)
    // exactly balanced: the lower index wins
    expect(mostVisiblePage(o, 1600, 800)).toBe(0)
    expect(mostVisiblePage(o, 1601, 800)).toBe(1)
  })

  it('mixed page sizes compare visible pixels, not ratios', () => {
    // page 0: 200px tall, fully visible; page 1: 2000px tall, 500px visible
    const o = pageOffsets([spec(1000, 200), spec(1000, 2000)], 1000, 0)
    expect(mostVisiblePage(o, 0, 700)).toBe(1)
  })

  it('maps the viewport through the zoom transform', () => {
    const o = pageOffsets([spec(1000, 1000), spec(1000, 1000), spec(1000, 1000)], 1000, 0)
    // zoom 2, panned so content y 1500..1900 fills an 800px viewport
    const s: ScrollPanState = { scrollTop: 0, panY: -16 - 1500 * 2, zoom: 2, padding: 16 }
    const top = contentYAt(0, s)
    const bottom = contentYAt(800, s)
    expect(top).toBeCloseTo(1500)
    expect(bottom).toBeCloseTo(1900)
    expect(mostVisiblePage(o, top, bottom - top)).toBe(1)
  })

  it('scrollPan: contentYAt includes native scrollTop; enterZoom and exitZoom keep the anchored content point fixed', () => {
    const s: ScrollPanState = { scrollTop: 1000, panY: 0, zoom: 1, padding: 16 }
    const y = contentYAt(400, s)
    expect(y).toBe(1384)
    const z = enterZoom(s, 2, 400)
    expect(z.scrollTop).toBe(0)
    expect(Math.abs(contentYAt(400, { scrollTop: 0, panY: z.panY, zoom: 2, padding: 16 }) - y)).toBeLessThanOrEqual(0.5)
    const back = exitZoom({ scrollTop: 0, panY: z.panY, zoom: 2, padding: 16 }, 400, 1e9)
    expect(back.panY).toBe(0)
    expect(Math.abs(contentYAt(400, { scrollTop: back.scrollTop, panY: 0, zoom: 1, padding: 16 }) - y)).toBeLessThanOrEqual(0.5)
    // exitZoom clamps to the scroll range
    expect(exitZoom({ scrollTop: 0, panY: z.panY, zoom: 2, padding: 16 }, 400, 500).scrollTop).toBe(500)
    expect(exitZoom({ scrollTop: 0, panY: 100, zoom: 2, padding: 16 }, 0, 500).scrollTop).toBe(0)
  })

  it('panYForContentTop is the exact inverse of contentYAt', () => {
    const a = { scrollTop: 0, zoom: 2, padding: 16 }
    expect(panYForContentTop(1000, { ...a, panY: 0 })).toBe(-2016)
    expect(contentYAt(0, { ...a, panY: -2016 })).toBe(1000)
    const b = { scrollTop: 40, zoom: 3, padding: 16 }
    expect(panYForContentTop(500, { ...b, panY: 0 })).toBe(-1476)
    expect(contentYAt(0, { ...b, panY: -1476 })).toBe(500)
    let seed = 7
    const rand = () => { seed = (seed * 16807) % 2147483647; return seed / 2147483647 }
    for (let i = 0; i < 200; i++) {
      const s: ScrollPanState = { scrollTop: rand() * 5000, panY: 0, zoom: 1 + rand() * 5, padding: rand() * 40 }
      const target = rand() * 100000
      const panY = panYForContentTop(target, s)
      expect(Math.abs(contentYAt(0, { ...s, panY }) - target)).toBeLessThan(1e-9 * Math.max(1, target))
    }
  })

  it('remapContentY keeps page index and page-local offset; gaps do not scale', () => {
    const fit = pageOffsets([spec(1000, 1000), spec(1000, 1000)], 2000, 20)
    const nat = pageOffsets([spec(1000, 1000), spec(1000, 1000)], 1000, 20)
    expect(fit).toEqual({ top: [0, 2020], height: [2000, 2000] })
    expect(nat).toEqual({ top: [0, 1020], height: [1000, 1000] })
    expect(remapContentY(2020 + 500, fit, nat)).toBe(1270)
    expect(remapContentY(2010, fit, nat)).toBe(1010)
    // before page 0 unchanged; after the last page offset from its bottom, unscaled
    expect(remapContentY(-5, fit, nat)).toBe(-5)
    expect(remapContentY(4020 + 30, fit, nat)).toBe(2020 + 30)
  })
})

/** A container element with a fixed client box (jsdom has no layout). */
function box(height: number, width = 1000): HTMLDivElement {
  const el = document.createElement('div')
  Object.defineProperty(el, 'clientHeight', { configurable: true, value: height })
  Object.defineProperty(el, 'clientWidth', { configurable: true, value: width })
  document.body.appendChild(el)
  return el
}

describe('usePageTracking', () => {
  let frames: FrameRequestCallback[] = []
  let rafCalls = 0
  beforeEach(() => {
    frames = []
    rafCalls = 0
    vi.stubGlobal('requestAnimationFrame', (cb: FrameRequestCallback) => { rafCalls++; frames.push(cb); return frames.length })
    vi.stubGlobal('cancelAnimationFrame', () => {})
  })
  afterEach(() => { vi.unstubAllGlobals() })
  const flush = () => act(() => { const f = frames; frames = []; for (const cb of f) cb(0) })
  const specs = Array.from({ length: 6 }, () => spec(1000, 1000))
  const identity: Transform = { zoom: 1, panX: 0, panY: 0 }

  it('updates at most once per animation frame and only on change', () => {
    const el = box(800)
    const ref = { current: el } as RefObject<HTMLElement | null>
    let renders = 0
    const { result } = renderHook(() => { renders++; return usePageTracking(ref, { specs, displayWidth: 1000, gap: 16 }, identity, () => {}) })
    flush()
    const baseRenders = renders
    const baseRaf = rafCalls
    // scroll within page 0: 50 events in one frame
    for (let i = 0; i < 50; i++) { el.scrollTop = 10 + i; el.dispatchEvent(new Event('scroll')) }
    expect(rafCalls - baseRaf).toBe(1)
    flush()
    expect(result.current.current).toBe(0)
    expect(renders).toBe(baseRenders)
    // cross into page 2
    el.scrollTop = 16 + 2 * 1016 + 100
    for (let i = 0; i < 50; i++) el.dispatchEvent(new Event('scroll'))
    expect(rafCalls - baseRaf).toBe(2)
    flush()
    expect(result.current.current).toBe(2)
    expect(renders).toBe(baseRenders + 1)
  })

  it('goTo scrolls the viewer container, never scrollIntoView', () => {
    Element.prototype.scrollIntoView ??= () => {}
    const spy = vi.spyOn(Element.prototype, 'scrollIntoView').mockImplementation(() => {})
    const el = box(800)
    const { result } = renderHook(() => usePageTracking({ current: el }, { specs, displayWidth: 1000, gap: 16 }, identity, () => {}))
    act(() => result.current.goTo(3))
    expect(el.scrollTop).toBe(16 + 3 * 1016)
    expect(spy).not.toHaveBeenCalled()
    spy.mockRestore()
  })

  it('goTo while zoomed sets panY through clampPan', () => {
    const el = box(800)
    const t: Transform = { zoom: 2, panX: -100, panY: -300 }
    const set = vi.fn()
    // a short last page, so it can never reach the viewport top
    const layout = { specs: [...specs.slice(0, 5), spec(1000, 200)], displayWidth: 1000, gap: 16 }
    const { result } = renderHook(() => usePageTracking({ current: el }, layout, t, set))
    const contentHeight = 5 * 1000 + 200 + 5 * 16
    const viewport = { width: 1000, height: 800, contentWidth: 1000, contentHeight }
    act(() => result.current.goTo(2))
    expect(set).toHaveBeenLastCalledWith(clampPan({ ...t, panY: panYForContentTop(2 * 1016, { scrollTop: 0, panY: t.panY, zoom: 2, padding: 16 }) }, viewport))
    // the last page cannot reach the top: clamped, not the raw inverse
    act(() => result.current.goTo(5))
    const raw = panYForContentTop(5 * 1016, { scrollTop: 0, panY: t.panY, zoom: 2, padding: 16 })
    const last = set.mock.calls.at(-1)![0] as Transform
    expect(last.panY).toBeGreaterThan(raw)
    expect(last).toEqual(clampPan({ ...t, panY: raw }, viewport))
    expect(el.scrollTop).toBe(0)
  })
})

describe('page tracking in the viewer', () => {
  const size = { width: 1000, height: 800 }
  let restore: () => void
  beforeEach(() => { h.pages = Array.from({ length: 8 }, () => ({ w: 800, h: 1000 })); restore = mockViewerLayout(size) })
  afterEach(() => { restore(); h.pages = null })

  async function mount(document = docxModel(['x'])) {
    const { OfficeDoc } = await import('../src/components/OfficeDoc')
    const ui = render(createElement(OfficeDoc, { document }))
    await waitFor(() => expect(ui.container.querySelectorAll('canvas').length).toBe(8))
    const bar = await ui.findByRole('toolbar', { name: 'Document toolbar' })
    return { ...ui, bar, region: ui.getByRole('region', { name: 'Document viewer' }), OfficeDoc }
  }
  const position = (bar: HTMLElement) => within(bar).getByTestId('officeview-page-position').textContent
  const transform = (c: HTMLElement) => {
    const t = (c.querySelector('[data-testid="officeview-content"]') as HTMLElement).style.transform
    const m = /translate3d\((-?[\d.e-]+)px, (-?[\d.e-]+)px, 0\) scale\(([\d.e-]+)\)/.exec(t)
    return m ? { panX: Number(m[1]), panY: Number(m[2]), zoom: Number(m[3]) } : { panX: 0, panY: 0, zoom: 1 }
  }

  it('native scroll to a later page, then zoom: tracking and goTo stay correct', async () => {
    const ui = await mount()
    // natural layout: 800px pages (container 1000 - 2*16 = 968 > 800), 1000px tall, gap 16
    ui.region.scrollTop = 16 + 3 * 1016
    fireEvent.scroll(ui.region)
    await waitFor(() => expect(position(ui.bar)).toBe('4 / 8'))
    fireEvent.click(within(ui.bar).getByLabelText('Zoom in'))
    expect(ui.region.scrollTop).toBe(0)
    expect(ui.region.style.overflow).toBe('hidden')
    await new Promise((r) => setTimeout(r, 40))
    expect(position(ui.bar)).toBe('4 / 8')
    fireEvent.click(within(ui.bar).getByLabelText('Next page'))
    await waitFor(() => expect(position(ui.bar)).toBe('5 / 8'))
    fireEvent.click(within(ui.bar).getByLabelText('Next page'))
    await waitFor(() => expect(position(ui.bar)).toBe('6 / 8'))
    const t = transform(ui.container)
    const top5 = 5 * 1016
    expect(t.panY).toBeCloseTo(panYForContentTop(top5, { scrollTop: 0, panY: 0, zoom: t.zoom, padding: 16 }), 6)
    expect(Math.abs(contentYAt(0, { scrollTop: 0, panY: t.panY, zoom: t.zoom, padding: 16 }) - top5)).toBeLessThanOrEqual(1)
    // back out to 1: native scrolling resumes, page 6 still tracked
    fireEvent.click(within(ui.bar).getByLabelText('Zoom out'))
    expect(ui.region.style.overflow).toBe('auto')
    expect(transform(ui.container).zoom).toBe(1)
    const anchored = contentYAt(400, { scrollTop: 0, panY: t.panY, zoom: t.zoom, padding: 16 })
    expect(Math.abs(contentYAt(400, { scrollTop: ui.region.scrollTop, panY: 0, zoom: 1, padding: 16 }) - anchored)).toBeLessThanOrEqual(1)
    fireEvent.scroll(ui.region)
    await new Promise((r) => setTimeout(r, 40))
    expect(position(ui.bar)).toBe('6 / 8')
  })

  it('a replacement document with the same page count resets to page 1', async () => {
    const ui = await mount()
    ui.region.scrollTop = 16 + 4 * 1016
    fireEvent.scroll(ui.region)
    await waitFor(() => expect(position(ui.bar)).toBe('5 / 8'))
    ui.rerender(createElement(ui.OfficeDoc, { document: docxModel(['y']) }))
    await waitFor(() => expect(position(ui.bar)).toBe('1 / 8'))
    expect(ui.region.scrollTop).toBe(0)
  })
})
