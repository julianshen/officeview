import { describe, it, expect, vi, beforeEach, afterEach } from 'vitest'
import { render, waitFor, fireEvent, act, within, type RenderResult } from '@testing-library/react'
import type { ComponentType } from 'react'
import { docxModel, mockViewerLayout, stubPointer, type FakePage } from './helpers/paged-viewer'
import { OfficePackage } from '../src/core/zip'
import { parseDocx } from '../src/docx/parse'
import { buildDocx } from '../src/testdata/ooxml-builders'

// Synthetic pages when `h.pages` is set; the real painter otherwise.
const h = vi.hoisted(() => ({ pages: null as FakePage[] | null, gate: null as Promise<void> | null, open: true }))
vi.mock('../src/render/paint', async (importOriginal) => {
  const m = await importOriginal<typeof import('../src/render/paint')>()
  const { fakePaintables: fake } = await import('./helpers/paged-viewer')
  return { ...m, getPaintables: (doc: never, opts: never) => (h.pages ? Promise.resolve(fake(h.pages)) : m.getPaintables(doc, opts)) }
})
// The paged chrome is a lazy chunk. Its module is cached after the first
// load, so `gate` holds the chunk's component suspended instead (React use()),
// which is what OfficeDoc's Suspense boundary sees while the import is open.
vi.mock('../src/components/paged/PagedChrome', async (importOriginal) => {
  const m = await importOriginal<typeof import('../src/components/paged/PagedChrome')>()
  // the legacy Suspense protocol, as React.lazy itself does: throw the
  // pending promise until it settles
  const Gated = (props: Parameters<typeof m.default>[0]) => { if (h.gate && !h.open) throw h.gate; return m.default(props) }
  return { ...m, default: Gated }
})

const size = { width: 1000, height: 800 }
let restoreLayout: () => void
let idle: (() => void)[] = []
const writeText = vi.fn().mockResolvedValue(undefined)

beforeEach(() => {
  h.pages = [{ w: 800, h: 1000, text: 'alpha one' }, { w: 800, h: 1000, text: 'alpha two' }, { w: 800, h: 1000, text: 'three' }]
  h.gate = null
  h.open = true
  size.width = 1000
  size.height = 800
  restoreLayout = mockViewerLayout(size)
  idle = []
  ;(window as unknown as { requestIdleCallback: (cb: () => void) => number }).requestIdleCallback = (cb) => { idle.push(cb); return idle.length }
  ;(window as unknown as { cancelIdleCallback: (id: number) => void }).cancelIdleCallback = () => {}
  writeText.mockClear()
  Object.defineProperty(navigator, 'clipboard', { value: { writeText }, configurable: true })
})
afterEach(() => {
  restoreLayout()
  delete (window as unknown as Record<string, unknown>).requestIdleCallback
  delete (window as unknown as Record<string, unknown>).cancelIdleCallback
})

async function flushIdle() {
  for (let i = 0; i < 1000 && idle.length > 0; i++) {
    const batch = idle
    idle = []
    await act(async () => { for (const cb of batch) cb() })
  }
}

const REPORT = docxModel(['The quick brown fox', 'jumps over the lazy dog', '日本語'])
const REPORT_WORDS = 12 // 4 + 5 + 3 (each CJK code point is a word)

async function mount(props: Record<string, unknown> = {}, Doc?: ComponentType<any>): Promise<RenderResult & { region: HTMLElement }> {
  const OfficeDoc = Doc ?? (await import('../src/components/OfficeDoc')).OfficeDoc
  const ui = render(<OfficeDoc document={REPORT} {...props} />)
  await waitFor(() => expect(ui.container.querySelectorAll('canvas').length).toBeGreaterThan(0))
  return { ...ui, region: ui.getByRole('region', { name: 'Document viewer' }) }
}
const toolbar = (ui: RenderResult) => ui.findByRole('toolbar', { name: 'Document toolbar' })
const status = (c: HTMLElement) => c.querySelector('[data-testid="officeview-status"]') as HTMLElement | null
const resize = async (width: number) => { size.width = width; await act(async () => { window.dispatchEvent(new Event('resize')) }) }

describe('paged toolbar and status line', () => {
  it('renders title, page position and word count for a DOCX', async () => {
    const ui = await mount({ title: 'Report.docx' })
    const bar = await toolbar(ui)
    const title = within(bar).getByText('Report.docx')
    expect(title.getAttribute('title')).toBe('Report.docx')
    expect(title.className).toContain('ov:truncate')
    expect(within(bar).getByTestId('officeview-page-position').textContent).toBe('1 / 3')
    expect(status(ui.container)!.textContent).toBe('Page 1 of 3 · — words')
    // the count may schedule or commit after one drain, so drain on every poll
    await waitFor(async () => { await flushIdle(); expect(status(ui.container)!.textContent).toBe(`Page 1 of 3 · ${REPORT_WORDS} words`) }, { timeout: 3000 })
  })

  it('formats large word counts with a thousands separator', async () => {
    const ui = await mount({ document: docxModel([Array.from({ length: 2418 }, (_, i) => `w${i}`).join(' ')]) })
    await toolbar(ui)
    // the count may schedule or commit after one drain, so drain on every poll
    await waitFor(async () => { await flushIdle(); expect(status(ui.container)!.textContent).toBe('Page 1 of 3 · 2,418 words') }, { timeout: 3000 })
  })

  it('uses Slide wording and no word count for PPTX', async () => {
    h.pages = [{ w: 960, h: 540 }, { w: 960, h: 540 }]
    const ui = await mount({ document: { slides: [{}, {}], widthEmu: 1, heightEmu: 1 } })
    const bar = await toolbar(ui)
    expect(within(bar).getByTestId('officeview-page-position').textContent).toBe('Slide 1 / 2')
    expect(within(bar).getByLabelText('Next slide')).toBeTruthy()
    await flushIdle()
    expect(status(ui.container)!.textContent).toBe('Slide 1 of 2')
  })

  it('static XLSX pages say Page, with no word count', async () => {
    h.pages = [{ w: 800, h: 600 }, { w: 800, h: 600 }]
    const ui = await mount({ document: { sheets: [{}, {}] } })
    await toolbar(ui)
    await flushIdle()
    expect(status(ui.container)!.textContent).toBe('Page 1 of 2')
  })

  it('hides Print and Copy when disallowed', async () => {
    const allowed = await mount()
    const bar = await toolbar(allowed)
    expect(within(bar).getByLabelText('Copy selection')).toBeTruthy()
    expect(within(bar).getByLabelText('Print')).toBeTruthy()
    allowed.unmount()
    const denied = await mount({ allowCopy: false, allowPrint: false })
    const bar2 = await toolbar(denied)
    expect(within(bar2).queryByLabelText('Copy selection')).toBeNull()
    expect(within(bar2).queryByLabelText('Print')).toBeNull()
    denied.unmount()
    // compact: not in the More menu either
    size.width = 375
    const compact = await mount({ allowCopy: false, allowPrint: false })
    fireEvent.click(within(await toolbar(compact)).getByLabelText('More options'))
    const menu = compact.getByRole('menu')
    expect(within(menu).queryByText('Copy selection')).toBeNull()
    expect(within(menu).queryByText('Print')).toBeNull()
    expect(within(menu).getByText('Fit to width')).toBeTruthy()
  })

  it('showToolbar=false renders no toolbar, status or placeholders; showSearch/showZoomControls hide their groups', async () => {
    const ui = await mount({ showToolbar: false })
    await new Promise((r) => setTimeout(r, 30))
    expect(ui.queryByRole('toolbar')).toBeNull()
    expect(status(ui.container)).toBeNull()
    expect(ui.container.querySelector('[data-testid="officeview-toolbar-placeholder"]')).toBeNull()
    expect(ui.container.querySelector('[data-testid="officeview-status-slot"]')).toBeNull()
    ui.unmount()
    const bare = await mount({ showSearch: false, showZoomControls: false })
    const bar = await toolbar(bare)
    expect(within(bar).queryByRole('searchbox')).toBeNull()
    expect(within(bar).queryByTestId('officeview-zoom-controls')).toBeNull()
    expect(within(bar).queryByLabelText('Zoom in')).toBeNull()
    expect(within(bar).getByTestId('officeview-page-position')).toBeTruthy()
  })

  it('placeholders reserve regular and compact heights before the chunk resolves', async () => {
    for (const c of [{ width: 1000, coarse: false, top: '40px' }, { width: 1000, coarse: true, top: '48px' }, { width: 375, coarse: false, top: '48px' }]) {
      let release!: () => void
      h.open = false
    h.gate = new Promise<void>((r) => { release = r }).then(() => { h.open = true })
      size.width = c.width
      const pointer = stubPointer(c.coarse)
      const { OfficeDoc } = await import('../src/components/OfficeDoc')
      const ui = await mount({}, OfficeDoc)
      const top = ui.container.querySelector('[data-testid="officeview-toolbar-placeholder"]') as HTMLElement
      const bottom = ui.container.querySelector('[data-testid="officeview-status-slot"]') as HTMLElement
      expect(top.style.height, JSON.stringify(c)).toBe(c.top)
      expect(bottom.style.height).toBe('24px')
      expect(ui.queryByRole('toolbar')).toBeNull()
      release()
      const bar = await toolbar(ui)
      expect(bar.style.height).toBe(c.top)
      expect(status(ui.container)!.style.height).toBe('24px')
      expect(ui.container.querySelector('[data-testid="officeview-toolbar-placeholder"]')).toBeNull()
      ui.unmount()
      pointer.restore()
    }
  })

  it('compact mode collapses Find and moves Fit/Copy/Print into More', async () => {
    size.width = 375
    const ui = await mount({ title: 'Report.docx' })
    const bar = await toolbar(ui)
    expect(within(bar).queryByText('Report.docx')).toBeNull()
    expect(within(bar).queryByRole('searchbox')).toBeNull()
    const find = within(bar).getByLabelText('Find')
    expect(find.getAttribute('aria-expanded')).toBe('false')
    expect(within(bar).queryByLabelText('Fit to width')).toBeNull()
    expect(within(bar).queryByLabelText('Print')).toBeNull()
    const more = within(bar).getByLabelText('More options')
    expect(more.getAttribute('aria-haspopup')).toBe('menu')
    expect(more.getAttribute('aria-expanded')).toBe('false')
    expect(more.querySelector('svg[data-icon="ellipsis"]')).toBeTruthy()
    // every toolbar target carries the 44px class in compact mode
    for (const b of within(bar).getAllByRole('button')) expect(b.className, b.getAttribute('aria-label') ?? '').toMatch(/ov:(coarse:)?min-h-11/)
    fireEvent.click(find)
    const input = await within(bar).findByRole('searchbox')
    expect(document.activeElement).toBe(input)
  })

  it('More menu: opens focused on first item, arrows/Home/End move, Escape and outside click close and return focus', async () => {
    size.width = 375
    const ui = await mount()
    const more = within(await toolbar(ui)).getByLabelText('More options')
    fireEvent.click(more)
    expect(more.getAttribute('aria-expanded')).toBe('true')
    const menu = ui.getByRole('menu')
    const items = [...menu.querySelectorAll<HTMLElement>('[role^=menuitem]')]
    expect(items.map((i) => i.textContent)).toEqual(['Fit to width', 'Copy selection', 'Print'])
    expect(document.activeElement).toBe(items[0])
    fireEvent.keyDown(items[0], { key: 'ArrowDown' })
    expect(document.activeElement).toBe(items[1])
    fireEvent.keyDown(items[1], { key: 'End' })
    expect(document.activeElement).toBe(items[2])
    fireEvent.keyDown(items[2], { key: 'ArrowDown' })
    expect(document.activeElement).toBe(items[0])
    fireEvent.keyDown(items[0], { key: 'ArrowUp' })
    expect(document.activeElement).toBe(items[2])
    fireEvent.keyDown(items[2], { key: 'Home' })
    expect(document.activeElement).toBe(items[0])
    fireEvent.keyDown(items[0], { key: 'Escape' })
    expect(ui.queryByRole('menu')).toBeNull()
    expect(document.activeElement).toBe(more)
    fireEvent.click(more)
    expect(ui.getByRole('menu')).toBeTruthy()
    fireEvent.pointerDown(document.body)
    expect(ui.queryByRole('menu')).toBeNull()
    expect(document.activeElement).toBe(more)
    // choosing Fit toggles fit mode and closes the menu
    fireEvent.click(more)
    fireEvent.click(within(ui.getByRole('menu')).getByText('Fit to width'))
    expect(ui.queryByRole('menu')).toBeNull()
    fireEvent.click(more)
    expect(within(ui.getByRole('menu')).getByRole('menuitemcheckbox').getAttribute('aria-checked')).toBe('true')
  })

  it('Cmd/Ctrl-F expands compact Find and focuses it; Escape collapses and returns focus to the Find button', async () => {
    size.width = 375
    const ui = await mount()
    const bar = await toolbar(ui)
    fireEvent.keyDown(ui.region, { key: 'f', ctrlKey: true })
    const input = await within(bar).findByRole('searchbox')
    await waitFor(() => expect(document.activeElement).toBe(input))
    fireEvent.keyDown(input, { key: 'Escape' })
    expect(within(bar).queryByRole('searchbox')).toBeNull()
    expect(document.activeElement).toBe(within(bar).getByLabelText('Find'))
  })

  it('desktop Find Escape returns focus to the document region', async () => {
    const ui = await mount()
    const input = within(await toolbar(ui)).getByRole('searchbox')
    input.focus()
    fireEvent.keyDown(input, { key: 'Escape' })
    expect(document.activeElement).toBe(ui.region)
  })

  it('crossing the compact threshold keeps query and active match and moves focus to the counterpart', async () => {
    const ui = await mount()
    const bar = await toolbar(ui)
    const input = within(bar).getByRole('searchbox') as HTMLInputElement
    fireEvent.change(input, { target: { value: 'alpha' } })
    await waitFor(() => expect(within(bar).getByTestId('officeview-search-status').textContent).toBe('1 of 2'))
    fireEvent.click(within(bar).getByLabelText('Next match'))
    await waitFor(() => expect(within(bar).getByTestId('officeview-search-status').textContent).toBe('2 of 2'))
    input.focus()
    await resize(375)
    const compactInput = within(bar).getByRole('searchbox') as HTMLInputElement
    expect(compactInput.value).toBe('alpha')
    expect(within(bar).getByTestId('officeview-search-status').textContent).toBe('2 of 2')
    expect(document.activeElement).toBe(compactInput)
    await resize(1000)
    expect(document.activeElement).toBe(within(bar).getByRole('searchbox'))
    // Fit (desktop) ↔ More (compact)
    within(bar).getByLabelText('Fit to width').focus()
    await resize(375)
    expect(document.activeElement).toBe(within(bar).getByLabelText('More options'))
    await resize(1000)
    expect(document.activeElement).toBe(within(bar).getByLabelText('Fit to width'))
  })

  it('zoom readout follows Ctrl-wheel and pinch zoom synchronously', async () => {
    const ui = await mount()
    const bar = await toolbar(ui)
    const readout = within(bar).getByLabelText('Reset zoom')
    expect(readout.textContent).toBe('100%')
    const canvas = ui.container.querySelector('canvas')!
    fireEvent.wheel(canvas, { ctrlKey: true, deltaY: -250, clientX: 100, clientY: 100 })
    expect(readout.textContent).toBe('272%')
    fireEvent.click(readout)
    expect(readout.textContent).toBe('100%')
    fireEvent.pointerDown(canvas, { pointerId: 1, clientX: 100, clientY: 100, pointerType: 'touch' })
    fireEvent.pointerDown(canvas, { pointerId: 2, clientX: 200, clientY: 100, pointerType: 'touch' })
    fireEvent.pointerMove(canvas, { pointerId: 2, clientX: 300, clientY: 100, pointerType: 'touch' })
    expect(readout.textContent).toBe('200%')
  })

  it('next/prev move the current page and update the status', async () => {
    const ui = await mount()
    const bar = await toolbar(ui)
    const prev = within(bar).getByLabelText('Previous page') as HTMLButtonElement
    const next = within(bar).getByLabelText('Next page') as HTMLButtonElement
    expect(prev.disabled).toBe(true)
    fireEvent.click(next)
    // page tops: 0, 1016, 2032 (1000px pages, 16px gap); padding 16
    expect(ui.region.scrollTop).toBe(16 + 1016)
    fireEvent.scroll(ui.region)
    await waitFor(() => expect(within(bar).getByTestId('officeview-page-position').textContent).toBe('2 / 3'))
    expect(status(ui.container)!.textContent).toMatch(/^Page 2 of 3/)
    fireEvent.click(next)
    fireEvent.scroll(ui.region)
    await waitFor(() => expect(within(bar).getByTestId('officeview-page-position').textContent).toBe('3 / 3'))
    expect(next.disabled).toBe(true)
    fireEvent.click(prev)
    expect(ui.region.scrollTop).toBe(16 + 1016)
  })
})

describe('copy and Find focus requests', () => {
  async function realDoc() {
    h.pages = null
    return parseDocx(await OfficePackage.load(await buildDocx([{ runs: [{ text: 'alpha beta gamma' }] }, { runs: [{ text: 'second line of text' }] }])))
  }

  it('copy confirmation still announces Copied', async () => {
    const doc = await realDoc()
    const ui = await mount({ document: doc })
    const bar = await toolbar(ui)
    const canvas = ui.container.querySelector('canvas') as HTMLCanvasElement
    canvas.getBoundingClientRect = () => ({ left: 0, top: 0, width: 816, height: 1056, right: 816, bottom: 1056, x: 0, y: 0, toJSON: () => ({}) }) as DOMRect
    const copy = within(bar).getByLabelText('Copy selection') as HTMLButtonElement
    expect(copy.disabled).toBe(true)
    fireEvent.doubleClick(canvas, { clientX: 160, clientY: 104, bubbles: true })
    await waitFor(() => expect(copy.disabled).toBe(false))
    fireEvent.click(copy)
    await waitFor(() => expect(copy.getAttribute('data-state')).toBe('copied'))
    expect(writeText).toHaveBeenCalledWith('beta')
    const live = ui.container.querySelector('[data-testid="officeview-copy-status"]')!
    expect([live.getAttribute('role'), live.getAttribute('aria-live'), live.textContent]).toEqual(['status', 'polite', 'Copied'])
  })

  it('Cmd/Ctrl-F before the chunk loads focuses the search input on mount', async () => {
    let release!: () => void
    h.open = false
    h.gate = new Promise<void>((r) => { release = r }).then(() => { h.open = true })
    const { OfficeDoc } = await import('../src/components/OfficeDoc')
    const ui = await mount({}, OfficeDoc)
    expect(ui.queryByRole('searchbox')).toBeNull()
    fireEvent.keyDown(ui.region, { key: 'f', ctrlKey: true })
    release()
    const search = await ui.findByRole('searchbox')
    await waitFor(() => expect(document.activeElement).toBe(search))
  })

  it('a queued request is honoured once, not replayed when the toolbar remounts', async () => {
    let release!: () => void
    h.open = false
    h.gate = new Promise<void>((r) => { release = r }).then(() => { h.open = true })
    const { OfficeDoc } = await import('../src/components/OfficeDoc')
    const ui = await mount({}, OfficeDoc)
    fireEvent.keyDown(ui.region, { key: 'f', ctrlKey: true })
    release()
    await waitFor(() => expect(document.activeElement).toBe(ui.getByRole('searchbox')))
    ui.getByRole('searchbox').blur()
    ui.rerender(<OfficeDoc document={REPORT} showToolbar={false} />)
    expect(ui.queryByRole('searchbox')).toBeNull()
    ui.rerender(<OfficeDoc document={REPORT} />)
    const search = await ui.findByRole('searchbox')
    await new Promise((r) => setTimeout(r, 20))
    expect(document.activeElement).not.toBe(search)
  })

  it('Cmd/Ctrl-F while search is disabled is dropped: enabling search later does not steal focus', async () => {
    let release!: () => void
    h.open = false
    h.gate = new Promise<void>((r) => { release = r }).then(() => { h.open = true })
    const { OfficeDoc } = await import('../src/components/OfficeDoc')
    const ui = await mount({ showSearch: false }, OfficeDoc)
    fireEvent.keyDown(ui.region, { key: 'f', ctrlKey: true })
    release()
    await toolbar(ui)
    fireEvent.keyDown(ui.region, { key: 'f', ctrlKey: true })
    ui.region.focus()
    ui.rerender(<OfficeDoc document={REPORT} showSearch />)
    const search = await ui.findByRole('searchbox')
    await new Promise((r) => setTimeout(r, 20))
    expect(document.activeElement).not.toBe(search)
    expect(document.activeElement).toBe(ui.region)
  })
})

describe('PagedToolbar Find requests', () => {
  const base = {
    kind: 'document' as const, current: 1, total: 3, onPrev() {}, onNext() {}, zoom: 1, onZoomIn() {}, onZoomOut() {}, onResetZoom() {},
    onFit() {}, fitActive: false, copyState: 'idle' as const, showSearch: true, showZoom: true, onEscapeToDocument() {},
    search: <input type="search" aria-label="Search document" />,
  }
  it.each([false, true])('a second request after focus moved away refocuses Find (compact %s)', async (compact) => {
    const { PagedToolbar } = await import('../src/components/paged/PagedToolbar')
    const outside = document.createElement('button')
    document.body.appendChild(outside)
    const ui = render(<PagedToolbar {...base} compact={compact} focusSearchRequest={1} />)
    await waitFor(() => expect(document.activeElement).toBe(ui.getByRole('searchbox')))
    outside.focus()
    ui.rerender(<PagedToolbar {...base} compact={compact} focusSearchRequest={2} />)
    await waitFor(() => expect(document.activeElement).toBe(ui.getByRole('searchbox')))
    outside.remove()
  })

  it('compact with zoom controls hidden offers no Fit, and no More menu when nothing is left', async () => {
    const { PagedToolbar } = await import('../src/components/paged/PagedToolbar')
    const ui = render(<PagedToolbar {...base} compact showZoom={false} onCopy={() => {}} focusSearchRequest={0} />)
    fireEvent.click(ui.getByLabelText('More options'))
    expect(within(ui.getByRole('menu')).queryByText('Fit to width')).toBeNull()
    expect(within(ui.getByRole('menu')).getByText('Copy selection')).toBeTruthy()
    ui.unmount()
    const bare = render(<PagedToolbar {...base} compact showZoom={false} focusSearchRequest={0} />)
    expect(bare.queryByLabelText('More options')).toBeNull()
  })

  it('compact menu items are 44px targets on any pointer', async () => {
    const { PagedToolbar } = await import('../src/components/paged/PagedToolbar')
    const ui = render(<PagedToolbar {...base} compact onPrint={() => {}} focusSearchRequest={0} />)
    fireEvent.click(ui.getByLabelText('More options'))
    for (const item of ui.getByRole('menu').querySelectorAll('[role^=menuitem]')) expect(item.className.split(' ')).toContain('ov:min-h-11')
  })

  it('no numeric position while the pages are pending', async () => {
    const { PagedToolbar } = await import('../src/components/paged/PagedToolbar')
    const { PagedStatus } = await import('../src/components/paged/PagedStatus')
    const ui = render(<><PagedToolbar {...base} current={1} total={0} compact={false} focusSearchRequest={0} /><PagedStatus kind="document" current={1} total={0} words={null} /></>)
    expect(ui.getByTestId('officeview-page-position').textContent).not.toMatch(/\d/)
    expect(ui.getByTestId('officeview-status').textContent).not.toMatch(/\d/)
  })

  it('the status position span carries its own test id, independent of child order (final review M7)', async () => {
    const { PagedStatus } = await import('../src/components/paged/PagedStatus')
    const ui = render(<PagedStatus kind="slides" current={3} total={12} words={1234} />)
    expect(ui.getByTestId('officeview-status-position').textContent).toBe('Slide 3 of 12')
    expect(ui.getByTestId('officeview-status').contains(ui.getByTestId('officeview-status-position'))).toBe(true)
  })
})

describe('pending pages in the viewer', () => {
  it('the toolbar and status never read "1 / 0" while the paintables resolve', async () => {
    let resolve!: () => void
    const held = new Promise<void>((r) => { resolve = r })
    const pages = h.pages!
    h.pages = null
    const paint = await import('../src/render/paint')
    const spy = vi.spyOn(paint, 'getPaintables').mockImplementation(async () => { await held; const { fakePaintables: fake } = await import('./helpers/paged-viewer'); return fake(pages) })
    try {
      const { OfficeDoc } = await import('../src/components/OfficeDoc')
      const ui = render(<OfficeDoc document={REPORT} />)
      const bar = await toolbar(ui)
      expect(within(bar).getByTestId('officeview-page-position').textContent).not.toMatch(/\d/)
      expect(status(ui.container)!.textContent).not.toMatch(/\d/)
      resolve()
      await waitFor(() => expect(within(bar).getByTestId('officeview-page-position').textContent).toBe('1 / 3'))
    } finally { spy.mockRestore() }
  })
})

describe('backdrop token (review I4)', () => {
  it('the scroll area uses --ov-color-backdrop unless the host passes background; pages stay white', async () => {
    const themed = await mount()
    expect(themed.region.style.background).toBe('var(--ov-color-backdrop)')
    expect((themed.container.querySelector('canvas') as HTMLElement).style.background).toBe('rgb(255, 255, 255)')
    themed.unmount()
    const host = await mount({ background: '#123456' })
    expect(host.region.style.background).toBe('rgb(18, 52, 86)')
  })
})
