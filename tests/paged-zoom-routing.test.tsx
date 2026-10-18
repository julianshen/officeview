import { describe, it, expect, vi, beforeEach, afterEach } from 'vitest'
import { render, waitFor, fireEvent, act } from '@testing-library/react'
import { docxModel, mockViewerLayout, type FakePage } from './helpers/paged-viewer'
import { contentYAt } from '../src/components/scrollPan'

// Synthetic pages, and the lazy zoom routing held until `release`.
const h = vi.hoisted(() => {
  let release!: () => void
  const gate = new Promise<void>((r) => { release = r })
  return { pages: null as FakePage[] | null, gate, release: () => release() }
})
vi.mock('../src/render/paint', async (importOriginal) => {
  const m = await importOriginal<typeof import('../src/render/paint')>()
  const { fakePaintables: fake } = await import('./helpers/paged-viewer')
  return { ...m, getPaintables: (doc: never, opts: never) => (h.pages ? Promise.resolve(fake(h.pages)) : m.getPaintables(doc, opts)) }
})
const PAD = 16
const size = { width: 1000, height: 800 }
let restore: () => void
beforeEach(() => { h.pages = Array.from({ length: 8 }, () => ({ w: 800, h: 1000 })); restore = mockViewerLayout(size) })
afterEach(() => restore())

function transformOf(el: HTMLElement) {
  const t = (el.querySelector('[data-testid="officeview-content"]') as HTMLElement).style.transform
  const m = /translate3d\((-?[\d.e-]+)px, (-?[\d.e-]+)px, 0\) scale\(([\d.e-]+)\)/.exec(t)
  return m ? { panX: Number(m[1]), panY: Number(m[2]), zoom: Number(m[3]) } : { panX: 0, panY: 0, zoom: 1 }
}

describe('inputs that arrive before the lazy zoom routing', () => {
  // One test: the routing module loads once per file, so every case is queued before the single release.
  it('a pinch started on a later page keeps its anchor, and queued input never reaches a replacement or unmounted document', async () => {
    const { OfficeDoc, routerLoader } = await import('../src/components/OfficeDoc')
    const load = routerLoader.load
    routerLoader.load = () => h.gate.then(load)
    const a = render(<OfficeDoc document={docxModel(['a'])} showToolbar={false} />)
    const b = render(<OfficeDoc document={docxModel(['b'])} showToolbar={false} />)
    const c = render(<OfficeDoc document={docxModel(['c'])} showToolbar={false} />)
    for (const ui of [a, b, c]) await waitFor(() => expect(ui.container.querySelectorAll('canvas').length).toBe(8))
    // A: natively scrolled to page 3, then a two-finger pinch x2 around y = 150
    const regionA = a.container.querySelector('[role=region]') as HTMLElement
    regionA.scrollTop = 16 + 2 * 1016
    const anchored = contentYAt(150, { scrollTop: regionA.scrollTop, panY: 0, zoom: 1, padding: PAD })
    const canvas = a.container.querySelector('canvas')!
    fireEvent.pointerDown(canvas, { pointerId: 1, clientX: 400, clientY: 100, pointerType: 'touch' })
    fireEvent.pointerDown(canvas, { pointerId: 2, clientX: 400, clientY: 200, pointerType: 'touch' })
    fireEvent.pointerMove(canvas, { pointerId: 2, clientX: 400, clientY: 250, pointerType: 'touch' })
    fireEvent.pointerMove(canvas, { pointerId: 1, clientX: 400, clientY: 50, pointerType: 'touch' })
    // B: a Ctrl-wheel queued, then the document is replaced
    fireEvent.wheel(b.container.querySelector('canvas')!, { ctrlKey: true, deltaY: -250, clientX: 100, clientY: 100 })
    b.rerender(<OfficeDoc document={docxModel(['b2'])} showToolbar={false} />)
    await waitFor(() => expect(b.container.querySelectorAll('canvas').length).toBe(8))
    // C: a Ctrl-wheel queued, then the viewer unmounts
    fireEvent.wheel(c.container.querySelector('canvas')!, { ctrlKey: true, deltaY: -250, clientX: 100, clientY: 100 })
    const cContainer = c.container
    c.unmount()
    expect(transformOf(a.container).zoom).toBe(1)
    await act(async () => { h.release(); await new Promise((r) => setTimeout(r, 30)) })
    const t = transformOf(a.container)
    expect(t.zoom).toBeCloseTo(2, 3)
    expect(regionA.scrollTop).toBe(0)
    // final fingers 50 and 250: distance x2, midpoint 150; its content point is unchanged
    expect(Math.abs(contentYAt(150, { scrollTop: 0, panY: t.panY, zoom: t.zoom, padding: PAD }) - anchored)).toBeLessThanOrEqual(1)
    expect(transformOf(b.container).zoom).toBe(1)
    expect(cContainer.innerHTML).toBe('')
  })
})
