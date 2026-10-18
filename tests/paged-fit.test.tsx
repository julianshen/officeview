import { describe, it, expect, vi, beforeEach, afterEach } from 'vitest'
import { render, waitFor, fireEvent, act, within, type RenderResult } from '@testing-library/react'
import { docxModel, mockViewerLayout, type FakePage } from './helpers/paged-viewer'
import { contentYAt, pageOffsets, remapContentY } from '../src/components/scrollPan'
import { stepZoom } from '../src/core/zoom'
import { OfficePackage } from '../src/core/zip'
import { parseDocx } from '../src/docx/parse'
import { buildDocx } from '../src/testdata/ooxml-builders'

const h = vi.hoisted(() => ({ pages: null as FakePage[] | null }))
vi.mock('../src/render/paint', async (importOriginal) => {
  const m = await importOriginal<typeof import('../src/render/paint')>()
  const { fakePaintables: fake } = await import('./helpers/paged-viewer')
  return { ...m, getPaintables: (doc: never, opts: never) => (h.pages ? Promise.resolve(fake(h.pages)) : m.getPaintables(doc, opts)) }
})

const PAD = 16
const size = { width: 1000, height: 800 }
let restoreLayout: () => void
const pages = (n: number, w = 600, ht = 800): FakePage[] => Array.from({ length: n }, () => ({ w, h: ht }))

beforeEach(() => {
  h.pages = pages(3)
  size.width = 1000
  size.height = 800
  restoreLayout = mockViewerLayout(size)
})
afterEach(() => { restoreLayout() })

async function mount(props: Record<string, unknown> = {}) {
  const { OfficeDoc } = await import('../src/components/OfficeDoc')
  const ui = render(<OfficeDoc document={docxModel(['x'])} {...props} />)
  await waitFor(() => expect(ui.container.querySelectorAll('canvas').length).toBeGreaterThan(0))
  const bar = await ui.findByRole('toolbar', { name: 'Document toolbar' })
  return { ...ui, bar, region: ui.getByRole('region', { name: 'Document viewer' }) }
}
const canvasWidth = (ui: RenderResult) => (ui.container.querySelector('canvas') as HTMLCanvasElement).style.width
const content = (ui: RenderResult) => ui.getByTestId('officeview-content')
/** zoom / panX / panY parsed from the content transform ('none' is identity). */
function transformOf(ui: RenderResult): { zoom: number; panX: number; panY: number } {
  const t = content(ui).style.transform
  if (!t || t === 'none') return { zoom: 1, panX: 0, panY: 0 }
  const m = /translate3d\((-?[\d.e-]+)px, (-?[\d.e-]+)px, 0\) scale\(([\d.e-]+)\)/.exec(t)!
  return { panX: Number(m[1]), panY: Number(m[2]), zoom: Number(m[3]) }
}
const resize = async (width: number) => { size.width = width; await act(async () => { window.dispatchEvent(new Event('resize')) }) }

describe('fit-to-width layout mode', () => {
  it('fit-to-width lays pages out at the pane width without a transform', async () => {
    const ui = await mount()
    const fit = within(ui.bar).getByLabelText('Fit to width')
    expect(fit.getAttribute('aria-pressed')).toBe('false')
    expect(canvasWidth(ui)).toBe('600px')
    fireEvent.click(fit)
    expect(canvasWidth(ui)).toBe('968px')
    expect(content(ui).style.transform).toBe('none')
    expect(ui.region.style.overflow).toBe('auto')
    expect(fit.getAttribute('aria-pressed')).toBe('true')
    expect(within(ui.bar).getByLabelText('Reset zoom').textContent).toBe('161%')
  })

  it('a pointer drag in fit mode selects text, not pan', async () => {
    h.pages = null
    const doc = await parseDocx(await OfficePackage.load(await buildDocx([{ runs: [{ text: 'alpha beta gamma' }] }])))
    const ui = await mount({ document: doc })
    fireEvent.click(within(ui.bar).getByLabelText('Fit to width'))
    const canvas = ui.container.querySelector('canvas') as HTMLCanvasElement
    const w = parseFloat(canvas.style.width)
    canvas.getBoundingClientRect = () => ({ left: 0, top: 0, width: w, height: w * 1056 / 816, right: w, bottom: w * 1056 / 816, x: 0, y: 0, toJSON: () => ({}) }) as DOMRect
    const s = w / 816
    fireEvent.pointerDown(canvas, { pointerId: 1, clientX: 100 * s, clientY: 104 * s, button: 0, pointerType: 'mouse' })
    fireEvent.pointerMove(canvas, { pointerId: 1, clientX: 260 * s, clientY: 104 * s, pointerType: 'mouse' })
    fireEvent.pointerMove(canvas, { pointerId: 1, clientX: 262 * s, clientY: 104 * s, pointerType: 'mouse' })
    fireEvent.pointerUp(canvas, { pointerId: 1, clientX: 262 * s, clientY: 104 * s, pointerType: 'mouse' })
    const copy = within(ui.bar).getByLabelText('Copy selection') as HTMLButtonElement
    await waitFor(() => expect(copy.disabled).toBe(false))
    expect(content(ui).style.transform).toBe('none')
    expect(ui.region.style.cursor).toBe('auto')
  })

  it('zooming from fit mode starts at the equivalent zoom and turns fit off', async () => {
    const ui = await mount()
    fireEvent.click(within(ui.bar).getByLabelText('Fit to width'))
    ui.region.scrollTop = 900
    const fitOff = pageOffsets(h.pages!.map((p) => ({ widthPx: p.w, heightPx: p.h })), 968, 16)
    const natOff = pageOffsets(h.pages!.map((p) => ({ widthPx: p.w, heightPx: p.h })), 600, 16)
    const yNat = remapContentY(contentYAt(400, { scrollTop: 900, panY: 0, zoom: 1, padding: PAD }), fitOff, natOff)
    fireEvent.click(within(ui.bar).getByLabelText('Zoom in'))
    const t = transformOf(ui)
    expect(t.zoom).toBeCloseTo(stepZoom(968 / 600, 1), 2)
    expect(within(ui.bar).getByLabelText('Fit to width').getAttribute('aria-pressed')).toBe('false')
    expect(canvasWidth(ui)).toBe('600px')
    expect(ui.region.scrollTop).toBe(0)
    expect(Math.abs(contentYAt(400, { scrollTop: 0, panY: t.panY, zoom: t.zoom, padding: PAD }) - yNat)).toBeLessThanOrEqual(1)
  })

  it('leaving fit mode on a later page keeps the anchor through page-relative remapping', async () => {
    h.pages = pages(10)
    const specs = h.pages.map((p) => ({ widthPx: p.w, heightPx: p.h }))
    const fitOff = pageOffsets(specs, 968, 24)
    const natOff = pageOffsets(specs, 600, 24)
    for (const c of [
      { at: fitOff.top[7] + 300, expected: natOff.top[7] + 300 * (600 / 968) },
      { at: fitOff.top[6] + fitOff.height[6] + 10, expected: natOff.top[6] + natOff.height[6] + 10 },
    ]) {
      const ui = await mount({ pageGapPx: 24 })
      fireEvent.click(within(ui.bar).getByLabelText('Fit to width'))
      // the anchor (viewport centre, y = 400) sits on `at`
      ui.region.scrollTop = c.at + PAD - 400
      fireEvent.click(within(ui.bar).getByLabelText('Zoom in'))
      const t = transformOf(ui)
      expect(Math.abs(contentYAt(400, { scrollTop: 0, panY: t.panY, zoom: t.zoom, padding: PAD }) - c.expected)).toBeLessThanOrEqual(1)
      ui.unmount()
    }
  })

  it('zoom out from fit follows stepZoom: 1.613 → 1.0756 (zoomed), then 1 (native natural layout)', async () => {
    h.pages = pages(6)
    const specs = h.pages.map((p) => ({ widthPx: p.w, heightPx: p.h }))
    const ui = await mount()
    const readout = within(ui.bar).getByLabelText('Reset zoom')
    fireEvent.click(within(ui.bar).getByLabelText('Fit to width'))
    expect(readout.textContent).toBe('161%')
    ui.region.scrollTop = 2000
    const yNat = remapContentY(contentYAt(400, { scrollTop: 2000, panY: 0, zoom: 1, padding: PAD }), pageOffsets(specs, 968, 16), pageOffsets(specs, 600, 16))
    fireEvent.click(within(ui.bar).getByLabelText('Zoom out'))
    expect(transformOf(ui).zoom).toBeCloseTo(1.0756, 3)
    expect(readout.textContent).toBe('108%')
    expect(ui.region.style.overflow).toBe('hidden')
    fireEvent.click(within(ui.bar).getByLabelText('Zoom out'))
    expect(readout.textContent).toBe('100%')
    expect(canvasWidth(ui)).toBe('600px')
    expect(ui.region.style.overflow).toBe('auto')
    expect(content(ui).style.transform).toBe('none')
    expect(Math.abs(contentYAt(400, { scrollTop: ui.region.scrollTop, panY: 0, zoom: 1, padding: PAD }) - yNat)).toBeLessThanOrEqual(1)
    // Ctrl-wheel: ratio applied to z0 (×0.5 → clamps to 1, native natural layout)
    fireEvent.click(within(ui.bar).getByLabelText('Fit to width'))
    fireEvent.wheel(ui.container.querySelector('canvas')!, { ctrlKey: true, deltaY: 250 * Math.log(2), clientX: 300, clientY: 300 })
    expect(readout.textContent).toBe('100%')
    expect(canvasWidth(ui)).toBe('600px')
    expect(within(ui.bar).getByLabelText('Fit to width').getAttribute('aria-pressed')).toBe('false')
    // pinch: ×1.5 from z0 → 2.42
    fireEvent.click(within(ui.bar).getByLabelText('Fit to width'))
    const canvas = ui.container.querySelector('canvas')!
    fireEvent.pointerDown(canvas, { pointerId: 1, clientX: 100, clientY: 100, pointerType: 'touch' })
    fireEvent.pointerDown(canvas, { pointerId: 2, clientX: 200, clientY: 100, pointerType: 'touch' })
    fireEvent.pointerMove(canvas, { pointerId: 2, clientX: 250, clientY: 100, pointerType: 'touch' })
    expect(readout.textContent).toBe('242%')
    expect(within(ui.bar).getByLabelText('Fit to width').getAttribute('aria-pressed')).toBe('false')
  })

  it('fit mode follows container resize; narrow widths keep the natural fallback', async () => {
    const ui = await mount()
    fireEvent.click(within(ui.bar).getByLabelText('Fit to width'))
    expect(canvasWidth(ui)).toBe('968px')
    await resize(1400)
    expect(canvasWidth(ui)).toBe('1368px')
    for (const w of [0, 1, 31, 32]) {
      await resize(w)
      expect(canvasWidth(ui), `fit ${w}`).toBe('600px')
    }
    await resize(33)
    expect(canvasWidth(ui)).toBe('1px')
    ui.unmount()
    const plain = await mount()
    for (const w of [0, 1, 31, 32]) {
      await resize(w)
      expect(canvasWidth(plain), `natural ${w}`).toBe('600px')
    }
    await resize(33)
    expect(canvasWidth(plain)).toBe('1px')
  })

  it('fit mode caps at MAX_ZOOM × natural width', async () => {
    h.pages = pages(2, 100, 140)
    const ui = await mount()
    fireEvent.click(within(ui.bar).getByLabelText('Fit to width'))
    expect(canvasWidth(ui)).toBe('600px')
  })

  it('fit mode keeps off-screen page backing stores at natural-mode size', async () => {
    h.pages = pages(30)
    const observed = new Map<Element, (e: { isIntersecting: boolean }[]) => void>()
    class IO {
      cb: (e: { isIntersecting: boolean }[]) => void
      constructor(cb: (e: { isIntersecting: boolean }[]) => void) { this.cb = cb }
      observe(el: Element) { observed.set(el, this.cb) }
      disconnect() {}
    }
    vi.stubGlobal('IntersectionObserver', IO)
    try {
      const ui = await mount()
      fireEvent.click(within(ui.bar).getByLabelText('Fit to width'))
      const canvases = [...ui.container.querySelectorAll('canvas')]
      await act(async () => { canvases.forEach((c, i) => observed.get(c)?.([{ isIntersecting: i < 2 }])) })
      await waitFor(() => expect(canvases[0].width).toBe(968))
      expect(canvases[1].width).toBe(968)
      for (const c of canvases.slice(2)) expect(c.width).toBeLessThanOrEqual(600)
      expect(canvases[29].width).toBe(600)
    } finally {
      vi.unstubAllGlobals()
    }
  })
})
