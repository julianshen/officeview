import { describe, it, expect, vi, beforeEach, afterEach } from 'vitest'
import { render, waitFor, fireEvent, within, type RenderResult } from '@testing-library/react'
import { docxModel, mockViewerLayout, type FakePage } from './helpers/paged-viewer'
import { showLoader, type PresentPlatform } from '../src/components/present'

/** jsdom platform: activation granted and the full-screen request resolves, so
 * a start reaches the synchronous focus handoff (the chunk stays gated). */
const platform: PresentPlatform = {
  hasActivation: () => true,
  requestFullscreen: () => Promise.resolve(),
  exitFullscreen: () => Promise.resolve(),
  fullscreenElement: () => null,
  open: () => null,
}

// Synthetic pages when `h.pages` is set; the real painter otherwise.
const h = vi.hoisted(() => ({ pages: null as FakePage[] | null }))
vi.mock('../src/render/paint', async (importOriginal) => {
  const m = await importOriginal<typeof import('../src/render/paint')>()
  const { fakePaintables: fake } = await import('./helpers/paged-viewer')
  return { ...m, getPaintables: (doc: never, opts: never) => (h.pages ? Promise.resolve(fake(h.pages)) : m.getPaintables(doc, opts)) }
})

const size = { width: 1000, height: 800 }
let restoreLayout: () => void
let idle: (() => void)[] = []
const PPTX = { slides: [{}, {}, {}], widthEmu: 1, heightEmu: 1 }
const pptxPages: FakePage[] = [{ w: 960, h: 540 }, { w: 960, h: 540 }, { w: 960, h: 540 }]

beforeEach(() => {
  h.pages = pptxPages
  size.width = 1000
  size.height = 800
  restoreLayout = mockViewerLayout(size)
  idle = []
  ;(window as unknown as { requestIdleCallback: (cb: () => void) => number }).requestIdleCallback = (cb) => { idle.push(cb); return idle.length }
  ;(window as unknown as { cancelIdleCallback: (id: number) => void }).cancelIdleCallback = () => {}
})

afterEach(() => {
  restoreLayout()
  delete (window as unknown as Record<string, unknown>).requestIdleCallback
  delete (window as unknown as Record<string, unknown>).cancelIdleCallback
})

const REPORT = docxModel(['The quick brown fox', 'jumps over the lazy dog'])

async function mount({ document: doc, ...props }: { document?: unknown } & Record<string, unknown> = {}): Promise<RenderResult & { region: HTMLElement }> {
  const { OfficeDoc } = await import('../src/components/OfficeDoc')
  const ui = render(<OfficeDoc document={(doc ?? REPORT) as never} {...props} />)
  await waitFor(() => expect(ui.container.querySelectorAll('canvas').length).toBeGreaterThan(0))
  return { ...ui, region: ui.getByRole('region', { name: 'Document viewer' }) }
}
const toolbar = async (ui: RenderResult) => ui.findByRole('toolbar', { name: 'Document toolbar' })

describe('paged Present button', () => {
  it('PPTX shows Present with From beginning, From current slide and Presenter View (stage 2)', async () => {
    const ui = await mount({ document: PPTX })
    const bar = await toolbar(ui)
    const present = within(bar).getByLabelText('Present')
    expect(present.tagName).toBe('BUTTON')
    expect(present.querySelector('svg[data-icon="presentation"]')).toBeTruthy()
    const more = within(bar).getByLabelText('Present options')
    expect(more.getAttribute('aria-haspopup')).toBe('menu')
    expect(more.getAttribute('aria-expanded')).toBe('false')
    fireEvent.click(more)
    const menu = ui.getByRole('menu', { name: 'Present options' })
    expect(more.getAttribute('aria-expanded')).toBe('true')
    const items = [...menu.querySelectorAll<HTMLElement>('[role^=menuitem]')]
    expect(items.map((i) => i.textContent)).toEqual(['From beginning', 'From current slide', 'Presenter View'])
  })

  it('the Presenter View item starts the show with the presenter window, inside the click', async () => {
    const load = vi.spyOn(showLoader, 'load').mockImplementation(() => new Promise(() => {}) as never)
    const open = vi.fn(() => null)
    try {
      const ui = await mount({ document: PPTX, __presentPlatform: { ...platform, open } })
      const bar = await toolbar(ui)
      fireEvent.click(within(bar).getByLabelText('Present options'))
      fireEvent.click(ui.getByRole('menuitem', { name: 'Presenter View' }))
      // presenterView: true opens the popup synchronously in the gesture (spec 4.1), from the beginning (as Alt+F5).
      expect(open).toHaveBeenCalledTimes(1)
      expect(open.mock.calls[0]).toEqual([expect.anything(), 'about:blank', 'popup'])
      expect(document.activeElement).toBe(ui.container.querySelector('[data-officeview-show]'))
      expect(ui.queryByRole('menu')).toBeNull()
    } finally {
      load.mockRestore()
    }
  })

  it('Present is hidden with showToolbar=false and allowPresent=false; never shown for DOCX or XLSX', async () => {
    const none = await mount({ document: PPTX, showToolbar: false })
    await new Promise((r) => setTimeout(r, 30))
    expect(none.queryByRole('toolbar')).toBeNull()
    none.unmount()
    const denied = await mount({ document: PPTX, allowPresent: false })
    const bar = await toolbar(denied)
    expect(within(bar).queryByLabelText('Present')).toBeNull()
    expect(within(bar).queryByLabelText('Present options')).toBeNull()
    // the rest of the toolbar (nav, zoom) is still there for PPTX
    expect(within(bar).getByRole('button', { name: 'Next slide' })).toBeTruthy()
    denied.unmount()
    const word = await mount()
    expect(within(await toolbar(word)).queryByLabelText('Present')).toBeNull()
    word.unmount()
    const sheet = await mount({ document: { sheets: [{}, {}] } })
    expect(within(await toolbar(sheet)).queryByLabelText('Present')).toBeNull()
    sheet.unmount()
  })

  it('pointer enter, focus and menu open call preload once', async () => {
    const load = vi.spyOn(showLoader, 'load')
    try {
      const ui = await mount({ document: PPTX })
      const bar = await toolbar(ui)
      const main = within(bar).getByLabelText('Present')
      const more = within(bar).getByLabelText('Present options')
      fireEvent.pointerEnter(main)
      fireEvent.focus(main)
      fireEvent.pointerEnter(more)
      fireEvent.click(more)
      expect(load).toHaveBeenCalledTimes(1)
    } finally {
      load.mockRestore()
    }
  })

  it('the preload hint prepares screen placement on every hover and focus, loading the chunk once (S2-3)', async () => {
    const prepareScreens = vi.fn(() => Promise.resolve())
    const load = vi.spyOn(showLoader, 'load').mockImplementation(() => Promise.resolve({ prepareScreens }) as never)
    try {
      const ui = await mount({ document: PPTX })
      const bar = await toolbar(ui)
      const main = within(bar).getByLabelText('Present')
      const more = within(bar).getByLabelText('Present options')
      fireEvent.pointerEnter(main)
      fireEvent.focus(more)
      fireEvent.click(more)
      await waitFor(() => expect(prepareScreens).toHaveBeenCalledTimes(3))
      // Keyed by the viewer's document: the synchronous start reads the slot for host.ownerDocument.
      expect(prepareScreens.mock.calls.every((c) => (c as unknown[])[0] === main.ownerDocument)).toBe(true)
      expect(load).toHaveBeenCalledTimes(1)
    } finally {
      load.mockRestore()
    }
  })

  it('busy state while the chunk loads', async () => {
    let release!: (m: unknown) => void
    const chunk = new Promise((resolve) => { release = resolve })
    const load = vi.spyOn(showLoader, 'load').mockImplementation(() => chunk as never)
    try {
      const ui = await mount({ document: PPTX })
      const bar = await toolbar(ui)
      const main = within(bar).getByLabelText('Present') as HTMLButtonElement
      fireEvent.click(main)
      expect(main.getAttribute('aria-busy')).toBe('true')
      expect(main.querySelector('svg[data-icon="loader"]')).toBeTruthy()
      release({})
      await waitFor(() => expect(main.getAttribute('aria-busy')).toBeNull())
      expect(main.querySelector('svg[data-icon="loader"]')).toBeNull()
    } finally {
      load.mockRestore()
    }
  })

  it('starting from a menu item puts focus on the show host', async () => {
    const load = vi.spyOn(showLoader, 'load').mockImplementation(() => new Promise(() => {}) as never)
    try {
      const ui = await mount({ document: PPTX, __presentPlatform: platform })
      const bar = await toolbar(ui)
      fireEvent.click(within(bar).getByLabelText('Present options'))
      fireEvent.click(ui.getByRole('menuitem', { name: 'From current slide' }))
      const host = ui.container.querySelector('[data-officeview-show]') as HTMLElement
      // the host holds focus (Esc and the show keys must reach it); the menu
      // closed and the split button did not steal the focus back
      expect(ui.queryByRole('menu')).toBeNull()
      expect(document.activeElement).toBe(host)
    } finally {
      load.mockRestore()
    }
  })

  it('compact: Present lives in More, and the zoom group stays', async () => {
    size.width = 375
    const load = vi.spyOn(showLoader, 'load').mockImplementation(() => new Promise(() => {}) as never)
    try {
      const ui = await mount({ document: PPTX, __presentPlatform: platform })
      const bar = await toolbar(ui)
      // the compact toolbar keeps its groups; Presentation is a More item
      expect(within(bar).queryByLabelText('Present')).toBeNull()
      expect(within(bar).queryByLabelText('Present options')).toBeNull()
      expect(within(bar).getByRole('button', { name: 'Zoom out' })).toBeTruthy()
      fireEvent.click(within(bar).getByLabelText('More options'))
      const items = [...ui.getByRole('menu').querySelectorAll<HTMLElement>('[role^=menuitem]')]
      expect(items.map((i) => i.textContent)).toEqual(['From beginning', 'From current slide', 'Presenter View', 'Fit to width', 'Copy selection', 'Print'])
      fireEvent.click(ui.getByRole('menuitem', { name: 'From current slide' }))
      const host = ui.container.querySelector('[data-officeview-show]') as HTMLElement
      expect(document.activeElement).toBe(host)
      expect(ui.queryByRole('menu')).toBeNull()
    } finally {
      load.mockRestore()
    }
  })
})
