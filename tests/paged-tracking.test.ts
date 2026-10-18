import { describe, it, expect } from 'vitest'
import { contentYAt, enterZoom, exitZoom, pageOffsets, panYForContentTop, remapContentY, type ScrollPanState } from '../src/components/scrollPan'
import { mostVisiblePage } from '../src/components/paged/pageTracking'
import type { PageSpec } from '../src/render/paint'

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
