/**
 * Piece 4 Task 1: the DOM-free slide preparation runs on the server (bun,
 * no DOM) under `@napi-rs/canvas` (`serverBackend`, the process-wide
 * `CanvasBackend`).
 */
import { afterAll, beforeAll, describe, expect, it } from 'bun:test'
import { installCanvasBackend, type CanvasBackend } from '../../src/core/canvas-backend'
import { hitRegionsOf, prepareSlide, rasterPreparedLayer } from '../../src/slideshow/prepare'
import { createDiagnosticSink } from '../../src/slideshow/diagnostics'
import type { CanvasFactory, Lease, SlideShowWorkspace, TimeNode } from '../../src/slideshow/types'
import { flyInDeck, loadShowDeck, triggerDeck } from '../../src/testdata/show-decks'
import { serverBackend } from '../src/backend'

let backend: CanvasBackend
let previous: CanvasBackend | undefined
beforeAll(() => { backend = serverBackend(); previous = installCanvasBackend(backend) })
afterAll(() => { installCanvasBackend(previous) })

function roomyWorkspace(): SlideShowWorkspace {
  return {
    limit: Number.POSITIVE_INFINITY, used: () => 0, reserveTerminal: () => undefined,
    lease: (bytes, tag) => {
      const lease: Lease = { tag, bytes, grow: () => true, shrink() {}, transfer: () => lease, release() {} }
      return lease
    },
  }
}

const factory = (): CanvasFactory => (w, h) => backend.createCanvas(Math.max(1, Math.ceil(w)), Math.max(1, Math.ceil(h))) as never
const measure = () => backend.createCanvas(1, 1).getContext('2d') as unknown as CanvasRenderingContext2D

/** Every spid an interactive sequence is triggered by (`onClick` conditions with a shape target). */
function triggerSpids(root: TimeNode): string[] {
  const out = new Set<string>()
  const walk = (n: TimeNode, depth: number) => {
    if (depth > 64) return
    for (const c of n.ctn.start) if (c.evt === 'onClick' && c.target && 'spid' in c.target) out.add(c.target.spid)
    if ('children' in n) for (const k of n.children) walk(k, depth + 1)
  }
  walk(root, 0)
  return [...out]
}

describe('prepare headless', () => {
  it('prepareSlide and rasterPreparedLayer run under bun with @napi-rs/canvas', async () => {
    expect(typeof document).toBe('undefined')
    const d = await loadShowDeck(await flyInDeck())
    const meta = await d.meta(0)
    const ps = prepareSlide({ index: 0, slide: d.doc.slides[0], meta, watermark: false })
    expect(ps.index).toBe(0)
    expect(ps.layers.length).toBeGreaterThanOrEqual(3)
    expect(ps.layers[0]).toMatchObject({ kind: 'static', background: true })
    expect([...ps.animated]).toEqual(['sp:3'])
    expect([...ps.hidden]).toEqual(['sp:3'])
    expect(ps.base.get('sp:3')).toMatchObject({ x: 380, y: 260, w: 200, h: 200 })
    const sink = createDiagnosticSink()
    for (const layer of ps.layers) {
      const r = rasterPreparedLayer(ps, layer, {
        scale: 1, images: [], resolveFont: f => f, canvas: factory(), workspace: roomyWorkspace(), maxTextureSize: 8192, sink,
      })
      if ('refused' in r) expect(r.refused).toBe('empty')
      else expect(r.canvas.width).toBeGreaterThan(0)
    }
  })

  it('hit regions are slide-px rects', async () => {
    const d = await loadShowDeck(await triggerDeck())
    const meta = await d.meta(0)
    const timing = meta.timing
    if (!timing || 'refused' in timing) throw new Error('triggerDeck has no timing')
    const spids = triggerSpids(timing.root)
    expect(spids).toEqual(['3'])
    const ps = prepareSlide({ index: 0, slide: d.doc.slides[0], meta, watermark: false })
    const regions = hitRegionsOf(ps, spids, measure(), f => f)
    expect([...regions.keys()]).toEqual(['3'])
    const rects = regions.get('3')!
    expect(rects).toHaveLength(1)
    expect(rects[0].w).toBeGreaterThan(0)
    expect(rects[0].h).toBeGreaterThan(0)
    // The button is at (80, 560), 160 × 80 slide px; the painted box pads it slightly.
    expect(rects[0].x).toBeLessThanOrEqual(80)
    expect(rects[0].x + rects[0].w).toBeGreaterThanOrEqual(240)
    expect(rects[0].y + rects[0].h).toBeLessThanOrEqual(720)
  })

  it('a requested group and a requested descendant both own the descendant leaf', async () => {
    const d = await loadShowDeck(await triggerDeck())
    const meta = await d.meta(0)
    const slide = d.doc.slides[0]
    const [box, button] = slide.shapes
    // Group 9 holds both shapes with an identity child transform (ext = chExt, chOff = off).
    const E = 9144000, H = 6858000
    const group = {
      ...box, xEmu: 0, yEmu: 0, widthEmu: E, heightEmu: H, rotationDeg: 0, fill: undefined, line: undefined, textBody: undefined,
      source: { ...box.source!, id: '9', name: 'Group' },
      group: { ext: { width: E, height: H }, chExt: { width: E, height: H }, chOff: { x: 0, y: 0 } },
      children: [box, button],
    } as typeof box
    const grouped = { ...slide, shapes: [group] }
    const ps = prepareSlide({ index: 0, slide: grouped, meta, watermark: false })
    const regions = hitRegionsOf(ps, ['9', '3'], measure(), f => f)
    expect(regions.get('9')).toHaveLength(2)
    expect(regions.get('3')).toHaveLength(1)
    expect(regions.get('9')).toContainEqual(regions.get('3')![0])
  })
})
