/**
 * Slideshow plan S2-5: the Presenter View UI and the single-screen swap
 * (spec 5.1 single-screen mode, 5.2 Presenter View contents, 6.3 protection
 * on presenter surfaces, 9 Presenter View accessibility).
 *
 * jsdom with node-canvas. The popup is a fake window over a second document
 * (as in slideshow-presenter-window.test.tsx); the session runs through
 * `runShow` with a recording stage, a ManualClock and injected monotonic and
 * wall clocks. Builder decks are "Office-shaped synthetic". Nothing waits on
 * wall time.
 */
import { describe, it, expect, vi, beforeEach, afterEach } from 'vitest'
import { readFileSync } from 'node:fs'
import { createRoot } from 'react-dom/client'
import { flushSync } from 'react-dom'
import { runShow } from '../src/slideshow'
import { sessionDeps, type SessionController } from '../src/slideshow/player/session'
import { ManualClock } from '../src/slideshow/clock'
import type { Stage } from '../src/slideshow/player/stage'
import { preparedPlacement, type PresentFullscreenOptions, type ShowStart } from '../src/components/present'
import { buildPptxDeck } from '../src/testdata/ooxml-builders'
import { OfficePackage } from '../src/core/zip'
import { parsePptx, pptxPackageHandle } from '../src/pptx/parse'
import type { PptxDocument } from '../src/pptx/types'
import type { WatermarkOptions } from '../src/core/watermark'
import { createElapsedTimer, formatElapsed, formatWallClock } from '../src/slideshow/presenter/timer'
import { NotesPane, NOTES_SIZES, NOTES_DEFAULT } from '../src/slideshow/presenter/notes'
import { SlideGrid, GRID_CELL, gridWindow } from '../src/slideshow/presenter/grid'
import { diagnosticKind, PANEL_KIND } from '../src/slideshow/presenter/diagnostics'
import { DIAGNOSTIC_CODES } from '../src/slideshow/diagnostics'
import { previewTimeline } from '../src/slideshow/presenter/preview'

type FakePopup = Window & { closeCalls: number }
function fakePopup(): FakePopup {
  const t = new EventTarget() as unknown as FakePopup & { closed: boolean }
  Object.assign(t, {
    document: document.implementation.createHTMLDocument(''),
    closed: false,
    closeCalls: 0,
    moveTo: vi.fn(),
    resizeTo: vi.fn(),
    close() {
      t.closeCalls++
      if (t.closed) return
      t.closed = true
      t.dispatchEvent(new Event('pagehide'))
    },
  })
  return t
}

const box = { id: 2, prst: 'rect', off: ['0', '0'] as [string, string], ext: ['914400', '914400'] as [string, string] }
const SCRIPT = 'Say <script>window.__pwned = 1</script> & <b>bold</b>'
const KIOSK = '<p:presentationPr xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main"><p:showPr><p:kiosk/></p:showPr></p:presentationPr>'
async function deck(opts: { hidden?: number; kiosk?: boolean } = {}): Promise<PptxDocument> {
  const bytes = await buildPptxDeck({
    ...(opts.kiosk ? { presProps: KIOSK } : {}),
    slides: [
      { shapes: [box], notes: [[SCRIPT], ['Second paragraph']] },
      { shapes: [box], hidden: opts.hidden === 1 },
      { shapes: [box] },
    ],
  })
  return parsePptx(await OfficePackage.load(bytes))
}

const recordingStage = (): Stage => ({
  leave() {}, async enter() {}, slideTime: (t) => t, frame: () => undefined, stateAt: () => new Map(),
  plan: undefined, toSlide: () => undefined, runAt: () => undefined, resize() {}, level: 'full', backend: 'canvas2d', backingStore: { w: 1, h: 1, dpr: 1 }, transition: undefined, inspect: () => ({ current: undefined, slides: [], lastDraws: [] }), dispose() {},
})

let now = 0
const WALL = Date.UTC(2026, 9, 23, 14, 5, 0)
beforeEach(() => {
  now = 0
  vi.spyOn(sessionDeps, 'clock').mockImplementation(() => new ManualClock())
  vi.spyOn(sessionDeps, 'stage').mockImplementation(() => recordingStage())
  vi.spyOn(sessionDeps, 'frames').mockImplementation(() => () => {})
  vi.spyOn(sessionDeps, 'now').mockImplementation(() => now)
  vi.spyOn(sessionDeps, 'wall').mockImplementation(() => WALL + now)
})
afterEach(() => {
  vi.restoreAllMocks()
  document.body.replaceChildren()
  document.head.replaceChildren()
  preparedPlacement.delete(document)
  delete (window as { getScreenDetails?: unknown }).getScreenDetails
  try { localStorage.clear() } catch { /* none */ }
})

const flush = async () => { for (let i = 0; i < 20; i++) await new Promise((r) => setTimeout(r, 0)) }

interface ShowOpts { watermark?: WatermarkOptions; protect?: { copy?: boolean; print?: boolean }; retry?: (o?: PresentFullscreenOptions) => Promise<'entered' | 'refused'> }
async function show(doc: PptxDocument, popup: Window | null | undefined, o: ShowOpts = {}) {
  // The viewer's theme root carries the protection policy (spec 6.3), the host sits beside the page region.
  const themeRoot = document.createElement('div')
  themeRoot.setAttribute('data-ov-theme-root', '')
  const region = document.createElement('div')
  region.setAttribute('data-officeview-root', '')
  if (o.protect?.copy === false) region.style.userSelect = 'none'
  if (o.protect?.print === false) themeRoot.setAttribute('data-ov-print', 'false')
  const host = document.createElement('div')
  host.tabIndex = -1
  themeRoot.append(region, host)
  document.body.append(themeRoot)
  const ac = new AbortController()
  const log: string[] = []
  const cleanup = vi.fn(() => { log.push('cleanup'); ac.abort(); popup?.close() })
  const start: ShowStart = {
    doc, handle: pptxPackageHandle(doc), host,
    token: { release() {}, released: false, owns: () => false },
    platform: { hasActivation: () => true, requestFullscreen: async () => {}, exitFullscreen: async () => {}, fullscreenElement: () => null, open: () => null },
    fullscreen: Promise.resolve('entered'),
    retryFullscreen: o.retry ?? (async () => 'entered'),
    from: 'beginning', motion: 'auto', presenterView: popup === undefined ? undefined : true, popup,
    ...(o.watermark ? { watermark: o.watermark } : {}),
    // The copy policy reaches the show as the start's allowCopy (PR #18 review 8), not the region's CSS.
    ...(o.protect?.copy === false ? { allowCopy: false } : {}),
    callbacks: {
      onStart: () => log.push('onStart'),
      onEnd: () => log.push('onEnd'),
      onDiagnostic: (d) => log.push(`diag:${d.code}`),
      currentSlide: () => 0, restoreFocus() {}, scrollToSlide() {},
    },
    signal: ac.signal, generation: 1, isCurrent: () => !ac.signal.aborted, cleanup,
  }
  const s = await runShow(start)
  await flush()
  return { s, log, cleanup, host }
}

const q = <T extends Element = HTMLElement>(root: ParentNode, sel: string) => root.querySelector<T & Element>(sel) as unknown as T
const button = (root: ParentNode, name: string) => q<HTMLButtonElement>(root, `button[aria-label="${name}"]`)
const key = (target: EventTarget, k: string, init: KeyboardEventInit = {}) => {
  const e = new KeyboardEvent('keydown', { key: k, bubbles: true, cancelable: true, ...init })
  target.dispatchEvent(e)
  return e.defaultPrevented
}
async function step(s: SessionController) { s.tick(); await flush() }

function renderInto(el: HTMLElement, node: React.ReactNode) {
  const root = createRoot(el)
  flushSync(() => root.render(node))
  return root
}

describe('Presenter View (S2-5)', () => {
  it('shows current, next, notes, elapsed, wall clock and Slide n of m', async () => {
    const popup = fakePopup()
    const r = await show(await deck(), popup)
    const pdoc = popup.document
    const view = q(pdoc, '[data-officeview-presenter-view]')
    expect(view).not.toBeNull()
    // Status: "Slide n of m".
    expect(q(pdoc, '[data-officeview-presenter-status]').textContent).toBe('Slide 1 of 3')
    // Current: the mirror pane holds a leased surface and copies the rendered audience frame.
    const audience = r.host.querySelector<HTMLCanvasElement>('[data-officeview-show-canvas]')!
    const actx = audience.getContext('2d')!
    actx.fillStyle = '#ff0000'
    actx.fillRect(0, 0, audience.width, audience.height)
    await step(r.s)
    const mirror = q<HTMLCanvasElement>(pdoc, '[data-officeview-presenter-current] canvas')
    expect(mirror.width).toBeGreaterThan(0)
    const px = [...mirror.getContext('2d')!.getImageData(Math.floor(mirror.width / 2), Math.floor(mirror.height / 2), 1, 1).data].slice(0, 3)
    expect(px).toEqual([255, 0, 0])
    // Next: the next visible slide's first checkpoint.
    const next = q(pdoc, '[data-officeview-presenter-next]')
    expect(next.getAttribute('aria-label')).toBe('Next: slide 2')
    expect(next.querySelector('canvas')).not.toBeNull()
    // Notes as text.
    expect(q(pdoc, '[data-officeview-presenter-notes]').textContent).toContain(SCRIPT)
    expect(pdoc.querySelector('script, b')).toBeNull()
    // Elapsed and wall clock on the injected clocks.
    expect(q(pdoc, '[data-officeview-presenter-elapsed]').textContent).toBe('0:00:00')
    expect(q(pdoc, '[data-officeview-presenter-clock]').textContent).toBe(formatWallClock(WALL))
    now = 65_000
    await step(r.s)
    expect(q(pdoc, '[data-officeview-presenter-elapsed]').textContent).toBe('0:01:05')
    expect(q(pdoc, '[data-officeview-presenter-clock]').textContent).toBe(formatWallClock(WALL + 65_000))
    // Moving on updates status and Next; the last slide's Next is the end of the show.
    r.s.next()
    await step(r.s)
    expect(q(pdoc, '[data-officeview-presenter-status]').textContent).toBe('Slide 2 of 3')
    r.s.next()
    await step(r.s)
    expect(q(pdoc, '[data-officeview-presenter-next]').getAttribute('aria-label')).toBe('Next: end of slide show')
    r.s.exit()
  })

  it('A+ and A- change the notes size; notes keep bullets and paragraph breaks', async () => {
    const popup = fakePopup()
    const r = await show(await deck(), popup)
    const pdoc = popup.document
    const notes = q(pdoc, '[data-officeview-presenter-notes]')
    expect(notes.style.fontSize).toBe(`${NOTES_SIZES[NOTES_DEFAULT]}px`)
    button(pdoc, 'Larger notes text').click()
    await flush()
    expect(notes.style.fontSize).toBe(`${NOTES_SIZES[NOTES_DEFAULT + 1]}px`)
    button(pdoc, 'Smaller notes text').click()
    await flush()
    button(pdoc, 'Smaller notes text').click()
    await flush()
    expect(notes.style.fontSize).toBe(`${NOTES_SIZES[NOTES_DEFAULT - 1]}px`)
    // Each paragraph is its own block holding one text node.
    const ps = notes.querySelectorAll('p')
    expect(ps).toHaveLength(2)
    for (const p of ps) expect([...p.childNodes].every((n) => n.nodeType === Node.TEXT_NODE)).toBe(true)
    // The size is clamped at both ends.
    for (let i = 0; i < NOTES_SIZES.length + 2; i++) { button(pdoc, 'Larger notes text').click(); await flush() }
    expect(notes.style.fontSize).toBe(`${NOTES_SIZES[NOTES_SIZES.length - 1]}px`)
    expect(button(pdoc, 'Larger notes text').disabled).toBe(true)
    r.s.exit()

    // Bullets and levels, rendered on their own.
    const el = document.createElement('div')
    document.body.append(el)
    const root = renderInto(el, <NotesPane notes={[{ level: 0, bullet: '•', text: 'One' }, { level: 1, bullet: '–', text: 'Two' }, { level: 0, text: 'Plain' }, { level: 0, bullet: 'arabicPeriod', text: 'Auto' }]} selectable />)
    const paras = [...el.querySelectorAll('p')]
    expect(paras.map((p) => p.textContent)).toEqual(['• One', '– Two', 'Plain', '• Auto'])
    expect(paras[1]!.style.marginInlineStart).toBe('1.5em')
    root.unmount()
  })

  it('timer pause and reset on the injected clock', async () => {
    const t = createElapsedTimer(1000)
    expect(t.elapsed(4000)).toBe(3000)
    t.pause(4000)
    expect(t.paused).toBe(true)
    expect(t.elapsed(9000)).toBe(3000)
    t.resume(10_000)
    expect(t.elapsed(12_000)).toBe(5000)
    t.reset(12_000)
    expect(t.elapsed(12_500)).toBe(500)
    t.pause(13_000)
    t.reset(14_000)
    expect(t.elapsed(20_000)).toBe(0)
    expect(formatElapsed(3_723_000)).toBe('1:02:03')
    expect(formatElapsed(-5)).toBe('0:00:00')

    // In the window: pause holds the readout, reset zeroes it, and both are announced.
    const popup = fakePopup()
    const r = await show(await deck(), popup)
    const pdoc = popup.document
    now = 5000
    await step(r.s)
    expect(q(pdoc, '[data-officeview-presenter-elapsed]').textContent).toBe('0:00:05')
    button(pdoc, 'Pause timer').click()
    await flush()
    expect(q(pdoc, '[data-officeview-presenter-timer-status]').textContent).toBe('Timer paused')
    now = 60_000
    await step(r.s)
    expect(q(pdoc, '[data-officeview-presenter-elapsed]').textContent).toBe('0:00:05')
    button(pdoc, 'Resume timer').click()
    await flush()
    now = 62_000
    await step(r.s)
    expect(q(pdoc, '[data-officeview-presenter-elapsed]').textContent).toBe('0:00:07')
    button(pdoc, 'Reset timer').click()
    await flush()
    expect(q(pdoc, '[data-officeview-presenter-elapsed]').textContent).toBe('0:00:00')
    expect(q(pdoc, '[data-officeview-presenter-timer-status]').textContent).toBe('Timer reset')
    // The show clock is not the timer: pausing the show (S) does not stop it.
    expect(r.s.state.paused).toBe(false)
    r.s.exit()
  })

  it('controls: previous, next, black, end show; grid jumps to a slide including hidden', async () => {
    const popup = fakePopup()
    const r = await show(await deck({ hidden: 1 }), popup)
    const pdoc = popup.document
    button(pdoc, 'Next slide').click()
    await flush()
    // Next skips the hidden slide.
    expect(r.s.slideIndex).toBe(2)
    button(pdoc, 'Previous slide').click()
    await flush()
    expect(r.s.slideIndex).toBe(0)
    const black = button(pdoc, 'Black screen')
    black.click()
    await step(r.s)
    expect(r.s.state.screen).toBe('black')
    expect(black.getAttribute('aria-pressed')).toBe('true')
    black.click()
    await step(r.s)
    expect(r.s.state.screen).toBe('slide')
    expect(black.getAttribute('aria-pressed')).toBe('false')

    const all = button(pdoc, 'See all slides')
    all.click()
    await flush()
    expect(all.getAttribute('aria-expanded')).toBe('true')
    const grid = q(pdoc, '[data-officeview-presenter-grid]')
    const items = [...grid.querySelectorAll<HTMLButtonElement>('button[data-slide]')]
    expect(items).toHaveLength(3)
    expect(items[1]!.getAttribute('aria-label')).toBe('Slide 2 (hidden)')
    expect(items[0]!.getAttribute('aria-current')).toBe('true')
    items[1]!.click()
    await flush()
    expect(r.s.slideIndex).toBe(1)
    expect(q(pdoc, '[data-officeview-presenter-grid]')).toBeNull()

    button(pdoc, 'End show').click()
    expect(r.log).toContain('onEnd')
    expect(popup.closed).toBe(true)
  })

  it('the grid holds at most 48 thumbnails and frees them as it scrolls', () => {
    expect(gridWindow({ top: 0, width: 0, height: 0 }, 300)).toEqual([0, 48])
    expect(gridWindow({ top: 0, width: 0, height: 0 }, 5)).toEqual([0, 5])
    const el = document.createElement('div')
    document.body.append(el)
    const held = new Set<number>()
    const thumb = vi.fn((i: number) => { held.add(i); const c = document.createElement('canvas'); c.width = 320; c.height = 180; return c })
    const release = vi.fn((i: number) => held.delete(i))
    const root = renderInto(el, <SlideGrid count={300} current={0} hidden={() => false} thumb={thumb} release={release} paint={() => {}} onJump={() => {}} />)
    expect(held.size).toBe(48)
    expect(thumb).toHaveBeenCalledTimes(48)
    const scroller = q(el, '[data-officeview-presenter-grid]')
    Object.defineProperty(scroller, 'clientWidth', { value: GRID_CELL.w * 2, configurable: true })
    Object.defineProperty(scroller, 'clientHeight', { value: GRID_CELL.h * 2, configurable: true })
    // jsdom does not scroll: the scroller's offset is defined directly.
    Object.defineProperty(scroller, 'scrollTop', { value: GRID_CELL.h * 50, configurable: true })
    flushSync(() => { scroller.dispatchEvent(new Event('scroll')) })
    expect(release).toHaveBeenCalledWith(0)
    expect(held.has(0)).toBe(false)
    expect(held.has(100)).toBe(true)
    expect(held.size).toBeLessThanOrEqual(48)
    root.unmount()
    // Unmounting frees every held thumbnail.
    expect(held.size).toBe(0)
  })

  it('diagnostics panel lists approximate, pending and refused effects once', async () => {
    expect(diagnosticKind({ code: 'transition-approximate', message: '' })).toBe('approximate')
    expect(diagnosticKind({ code: 'transition-stand-in', message: '' })).toBe('approximate')
    expect(diagnosticKind({ code: 'media-autoplay', message: '' })).toBe('pending')
    expect(diagnosticKind({ code: 'link-refused', message: '' })).toBe('refused')
    expect(diagnosticKind({ code: 'transition-unknown', message: '' })).toBe('refused')
    const popup = fakePopup()
    const r = await show(await deck(), popup)
    const sink = r.s.presenterSource.sink
    for (let i = 0; i < 3; i++) {
      sink.add({ code: 'transition-approximate', slideIndex: 0, effect: 'p14:vortex', message: 'Vortex is approximated by a fade' })
      sink.add({ code: 'media-autoplay', slideIndex: 0, effect: 'video:4', message: 'Click to play' })
      sink.add({ code: 'link-refused', slideIndex: 0, effect: 'link:5', message: 'javascript: link refused' })
    }
    await step(r.s)
    await step(r.s)
    const panel = q(popup.document, '[data-officeview-presenter-diagnostics]')
    expect(panel.localName).toBe('details')
    expect(panel.querySelector('summary')!.textContent).toContain('3')
    const text = panel.textContent ?? ''
    for (const m of ['Vortex is approximated by a fade', 'Click to play', 'javascript: link refused']) {
      expect(text.split(m)).toHaveLength(2)
    }
    const groups = [...panel.querySelectorAll('[data-kind]')].map((g) => g.getAttribute('data-kind'))
    expect(groups).toEqual(['refused', 'approximate', 'pending'])
    // Never on the audience picture.
    expect(r.host.textContent).not.toContain('Vortex')
    r.s.exit()
  })

  it('keys typed in the slide-number field stay in the field except Esc', async () => {
    const popup = fakePopup()
    const r = await show(await deck(), popup)
    const field = q<HTMLInputElement>(popup.document, 'input[aria-label="Go to slide"]')
    expect(field).not.toBeNull()
    for (const k of ['ArrowRight', 'ArrowDown', 'PageDown', 'n', 'b', 'w', ' ', '2', 'Home', 'End', 'Backspace']) {
      expect(key(field, k)).toBe(false)
    }
    await flush()
    expect(r.s.slideIndex).toBe(0)
    expect(r.s.state.screen).toBe('slide')
    expect(r.s.state.buffer).toBe('')
    // The field's own Enter goes to the typed slide.
    field.value = '3'
    key(field, 'Enter')
    await flush()
    expect(r.s.slideIndex).toBe(2)
    // Esc still exits.
    key(field, 'Escape')
    expect(r.log).toContain('onEnd')
  })

  it('Alt+F5 and the toggle swap the full-screen surface between audience and presenter', async () => {
    // Not wanted at start ('none'): the audience shows; Alt+F5 swaps the same surface to Presenter View.
    const r = await show(await deck(), undefined)
    const surface = q(r.host, '[data-officeview-show-surface]')
    expect(r.s.presenter).toBe('none')
    expect(q(r.host, '[data-officeview-presenter-view]')).toBeNull()
    expect(key(r.host, 'F5', { altKey: true })).toBe(true)
    await flush()
    expect(r.s.presenter).toBe('swap')
    const view = q(r.host, '[data-officeview-presenter-view]')
    expect(view).not.toBeNull()
    expect(view.getAttribute('data-host')).toBe('swap')
    expect(surface.hidden).toBe(true)
    // Presenter keys work on the swapped surface; the notes keep theirs.
    key(r.host, 'ArrowRight')
    await flush()
    expect(r.s.slideIndex).toBe(1)
    key(q(r.host, '[data-officeview-presenter-notes]'), 'ArrowRight')
    await flush()
    expect(r.s.slideIndex).toBe(1)
    // Wheel and swipe over the swapped-in view scroll it; they are not show input.
    const wheel = new WheelEvent('wheel', { deltaY: 120, bubbles: true, cancelable: true })
    q(r.host, '[data-officeview-presenter-notes]').dispatchEvent(wheel)
    await flush()
    expect(wheel.defaultPrevented).toBe(false)
    expect(r.s.slideIndex).toBe(1)
    // A focused presenter button keeps only Enter and Space, as in the popup.
    expect(key(button(r.host, 'Black screen'), ' ')).toBe(false)
    key(button(r.host, 'Black screen'), 'ArrowRight')
    await flush()
    expect(r.s.slideIndex).toBe(2)
    // The toggle swaps back to the audience; Alt+F5 forth again, from a focused presenter button too.
    button(r.host, 'Show audience view').click()
    await flush()
    expect(q(r.host, '[data-officeview-presenter-view]')).toBeNull()
    expect(surface.hidden).toBe(false)
    key(r.host, 'F5', { altKey: true })
    await flush()
    key(button(r.host, 'Next slide'), 'F5', { altKey: true })
    await flush()
    expect(q(r.host, '[data-officeview-presenter-view]')).toBeNull()
    // Ctrl+Alt+F5 is not the toggle.
    key(r.host, 'F5', { altKey: true, ctrlKey: true })
    await flush()
    expect(q(r.host, '[data-officeview-presenter-view]')).toBeNull()
    r.s.exit()
    expect(r.host.childElementCount).toBe(0)

    // Blocked popup: Presenter View was asked for, so the single screen starts on it.
    const b = await show(await deck(), null)
    expect(b.s.presenter).toBe('swap')
    expect(q(b.host, '[data-officeview-presenter-view]')).not.toBeNull()
    key(b.host, 'F5', { altKey: true })
    await flush()
    expect(q(b.host, '[data-officeview-presenter-view]')).toBeNull()
    // Esc exits from either view.
    key(b.host, 'Escape')
    b.s.exit()

    // With the popup adopted, Alt+F5 changes nothing.
    const popup = fakePopup()
    const w = await show(await deck(), popup)
    key(popup.document.body, 'F5', { altKey: true })
    await flush()
    expect(w.s.presenter).toBe('window')
    expect(q(w.host, '[data-officeview-presenter-view]')).toBeNull()
    w.s.exit()
  })

  it('labels and focus management inside the popup document', async () => {
    // jsdom focuses only in a document with a window (the fake popup's has none, and an iframe's is another
    // realm jsdom refuses listeners from), so the focus calls themselves are recorded.
    const focused: Element[] = []
    const real = HTMLElement.prototype.focus
    vi.spyOn(HTMLElement.prototype, 'focus').mockImplementation(function (this: HTMLElement, o?: FocusOptions) { focused.push(this); real.call(this, o) })
    const popup = fakePopup()
    const r = await show(await deck(), popup)
    const pdoc = popup.document
    const view = q(pdoc, '[data-officeview-presenter-view]')
    expect(view.getAttribute('aria-label')).toBe('Presenter View')
    // Focus moved into the popup's own document.
    expect(view.contains(focused.at(-1)!)).toBe(true)
    // Every control has a name; the panes and toolbar are labelled.
    for (const b of view.querySelectorAll('button')) expect((b.getAttribute('aria-label') ?? b.textContent ?? '').trim()).not.toBe('')
    expect(q(pdoc, '[data-officeview-presenter-notes]').getAttribute('aria-label')).toBe('Speaker notes')
    expect(q(pdoc, '[role="toolbar"]').getAttribute('aria-label')).toBe('Slide controls')
    expect(q(pdoc, '[data-officeview-presenter-current]').getAttribute('aria-label')).toBe('Current slide')
    // Live regions: slide changes and the timer's state; the running readout is a timer, not a live region.
    expect(q(pdoc, '[data-officeview-presenter-status]').getAttribute('aria-live')).toBe('polite')
    expect(q(pdoc, '[data-officeview-presenter-timer-status]').getAttribute('aria-live')).toBe('polite')
    expect(q(pdoc, '[data-officeview-presenter-elapsed]').getAttribute('role')).toBe('timer')
    // The grid takes focus to the current slide and gives it back to its opener.
    const all = button(pdoc, 'See all slides')
    all.click()
    await flush()
    expect(focused.at(-1)!.getAttribute('aria-current')).toBe('true')
    expect(q(pdoc, '[data-officeview-presenter-grid]').contains(focused.at(-1)!)).toBe(true)
    q<HTMLButtonElement>(pdoc, '[data-officeview-presenter-grid] button[data-slide="2"]').click()
    await flush()
    expect(r.s.slideIndex).toBe(2)
    expect(focused.at(-1)).toBe(all)
    // The grid covers the controls: its own close button closes it and focus goes back the same way.
    all.click()
    await flush()
    button(pdoc, 'Close all slides').click()
    await flush()
    expect(q(pdoc, '[data-officeview-presenter-grid]')).toBeNull()
    expect(focused.at(-1)).toBe(all)
    // Every focus Presenter View moved stayed in the popup's document; none was taken from the opener.
    const mine = focused.slice(focused.indexOf(view))
    expect(mine.every((el) => el.ownerDocument === pdoc)).toBe(true)
    r.s.exit()
  })

  it('allowCopy=false makes notes unselectable; allowPrint=false hides the presenter popup in print; grid thumbnails carry the watermark', async () => {
    const watermark: WatermarkOptions = { text: 'OV-STEP', color: '#ff00ff', opacity: 1, fontSizePt: 320, placement: 'center' }
    const popup = fakePopup()
    const r = await show(await deck(), popup, { watermark, protect: { copy: false, print: false } })
    const pdoc = popup.document
    const notes = q(pdoc, '[data-officeview-presenter-notes]')
    expect(notes.style.userSelect).toBe('none')
    const root = q(pdoc, '[data-officeview-presenter]')
    expect(root.getAttribute('data-ov-print')).toBe('false')
    const css = readFileSync('src/styles/officeview.css', 'utf8')
    expect(css).toMatch(/@media print[\s\S]*\[data-ov-theme-root\]\[data-officeview-presenter\]\[data-ov-print="false"\][\s\S]*display: none !important/)
    button(pdoc, 'See all slides').click()
    await flush()
    const thumb = q<HTMLCanvasElement>(pdoc, '[data-officeview-presenter-grid] button[data-slide="1"] canvas')
    expect(thumb.width).toBe(320)
    const data = thumb.getContext('2d')!.getImageData(0, 0, thumb.width, thumb.height).data
    let magenta = false
    for (let i = 0; i < data.length; i += 4) if (data[i]! > 180 && data[i + 2]! > 180 && data[i + 1]! < 120) { magenta = true; break }
    expect(magenta).toBe(true)
    r.s.exit()

    // Unprotected: notes are selectable and the popup prints.
    const p2 = fakePopup()
    const u = await show(await deck(), p2)
    expect(q(p2.document, '[data-officeview-presenter-notes]').style.userSelect).toBe('')
    expect(q(p2.document, '[data-officeview-presenter]').getAttribute('data-ov-print')).toBeNull()
    u.s.exit()
  })

  it('Swap displays, Use both screens and the drag hint; a refusal in the popup is retried from the opener', async () => {
    const presenterScreen = { availLeft: 0, availTop: 0, availWidth: 1440, availHeight: 900, isPrimary: true, label: 'laptop' }
    const audienceScreen = { availLeft: 1440, availTop: 0, availWidth: 1920, availHeight: 1080, label: 'projector' }
    preparedPlacement.set(document, { popupFeatures: 'popup=yes,left=0,top=0,width=1440,height=900', audienceScreen, presenterScreen })
    const asked: Array<string | undefined> = []
    let answer: 'entered' | 'refused' = 'refused'
    const popup = fakePopup()
    const r = await show(await deck(), popup, { retry: async (o) => { asked.push(o?.screen?.label); return answer } })
    const pdoc = popup.document
    expect(button(pdoc, 'Use both screens')).toBeNull()
    expect(pdoc.querySelector('[data-officeview-drag-hint]')).toBeNull()
    // The click in the popup gives the opener no activation: the request is refused, the opener offers the retry.
    button(pdoc, 'Swap displays').click()
    await flush()
    expect(asked).toEqual(['laptop'])
    const notice = () => r.host.querySelector<HTMLElement>('[data-officeview-show-notice][data-placement]')
    expect(notice()).not.toBeNull()
    expect(popup.moveTo).not.toHaveBeenCalled()
    answer = 'entered'
    q<HTMLButtonElement>(notice()!, 'button').click()
    await flush()
    expect(asked).toEqual(['laptop', 'laptop'])
    expect(notice()).toBeNull()
    expect(popup.moveTo).toHaveBeenCalledWith(1440, 0)
    // Swapped: a second swap goes back.
    button(pdoc, 'Swap displays').click()
    await flush()
    expect(asked.at(-1)).toBe('projector')
    r.s.exit()

    // Not yet granted, but the API exists: "Use both screens" asks on that click.
    preparedPlacement.delete(document)
    const details = { screens: [presenterScreen, audienceScreen] }
    const get = vi.fn(() => Promise.resolve(details))
    ;(window as { getScreenDetails?: unknown }).getScreenDetails = get
    const p2 = fakePopup()
    const u = await show(await deck(), p2, { retry: async () => 'entered' })
    expect(button(p2.document, 'Swap displays')).toBeNull()
    button(p2.document, 'Use both screens').click()
    expect(get).toHaveBeenCalledTimes(1)
    await flush()
    expect(p2.moveTo).toHaveBeenCalledWith(0, 0)
    expect(button(p2.document, 'Swap displays')).not.toBeNull()
    u.s.exit()

    // No Window Management API (and nothing prepared): the one-time drag hint, in the popup.
    delete (window as { getScreenDetails?: unknown }).getScreenDetails
    preparedPlacement.delete(document)
    const p3 = fakePopup()
    const h = await show(await deck(), p3)
    expect(p3.document.querySelector('[data-officeview-drag-hint]')).not.toBeNull()
    expect(button(p3.document, 'Use both screens')).toBeNull()
    h.s.exit()
  })

  describe('fix round 1', () => {
    it('I1: in kiosk, presenter input reaches the show; audience input stays blocked', async () => {
      const popup = fakePopup()
      const r = await show(await deck({ kiosk: true }), popup)
      expect(r.s.state.kiosk).toBe(true)
      const pdoc = popup.document
      // Audience keys on the host do nothing in kiosk.
      key(r.host, 'ArrowRight')
      key(r.host, 'b')
      await flush()
      expect(r.s.slideIndex).toBe(0)
      expect(r.s.state.screen).toBe('slide')
      // The presenter's Black screen button and keys work.
      button(pdoc, 'Black screen').click()
      await step(r.s)
      expect(r.s.state.screen).toBe('black')
      key(pdoc.body, 'w')
      await flush()
      expect(r.s.state.screen).toBe('slide')
      key(pdoc.body, 'w')
      await flush()
      expect(r.s.state.screen).toBe('white')
      key(pdoc.body, 'ArrowRight')
      await flush()
      expect(r.s.state.screen).toBe('slide')
      key(pdoc.body, 'ArrowRight')
      await flush()
      expect(r.s.slideIndex).toBe(1)
      r.s.exit()

      // The swapped-in view is presenter input too.
      const b = await show(await deck({ kiosk: true }), null)
      key(q(b.host, '[data-officeview-presenter-view]'), 'ArrowRight')
      await flush()
      expect(b.s.slideIndex).toBe(1)
      b.s.exit()
    })

    it('I2: allowPrint=false keeps the single-screen swap view (and its notes) out of print', async () => {
      const r = await show(await deck(), null, { protect: { print: false } })
      const box = q(r.host, '[data-officeview-presenter-swap]')
      expect(box.getAttribute('data-ov-print')).toBe('false')
      expect(box.querySelector('[data-officeview-presenter-notes]')).not.toBeNull()
      const css = readFileSync('src/styles/officeview.css', 'utf8')
      expect(css).toMatch(/@media print[\s\S]*\[data-ov-theme-root\] \[data-officeview-presenter-swap\]\[data-ov-print="false"\][\s\S]*display: none !important/)
      r.s.exit()
      const u = await show(await deck(), null)
      expect(q(u.host, '[data-officeview-presenter-swap]').getAttribute('data-ov-print')).toBeNull()
      u.s.exit()
    })

    it('M1: the preview hands its holds back when the show reaches its end', async () => {
      const popup = fakePopup()
      const r = await show(await deck(), popup)
      const model = r.s.presenterSource.model!
      r.s.goTo(2)
      await step(r.s)
      await flush()
      expect(previewTimeline(model, 2)).toBeDefined()
      r.s.next()
      await step(r.s)
      await flush()
      expect(r.s.state.slide).toBe('end')
      expect(previewTimeline(model, 2)).toBeUndefined()
      r.s.exit()
    })

    it('M2: Next repaints when the stage decodes its pictures after the first paint', async () => {
      const images: Array<CanvasImageSource | undefined> = []
      const calls: number[] = []
      vi.mocked(sessionDeps.stage).mockImplementation((deps) => ({
        ...recordingStage(),
        previewOf: (i: number) => { calls.push(i); return { slide: deps.doc.slides[i]!, state: new Map(), images, resolveFont: (f: string) => f, builds: [] } },
      }))
      const popup = fakePopup()
      const r = await show(await deck(), popup)
      await step(r.s)
      const before = calls.filter((i) => i === 1).length
      expect(before).toBeGreaterThan(0)
      await step(r.s)
      expect(calls.filter((i) => i === 1).length).toBe(before)
      // A picture arrives: the next frame repaints Next once.
      images[0] = document.createElement('canvas')
      await step(r.s)
      await step(r.s)
      expect(calls.filter((i) => i === 1).length).toBe(before + 1)
      r.s.exit()
    })

    it('M3: a window with more visible cells than the cap leases the visible ones first and marks the rest', () => {
      // 12 columns x 6 visible rows = 72 visible cells: the window starts at the first visible cell.
      expect(gridWindow({ top: GRID_CELL.h * 10, width: GRID_CELL.w * 12, height: GRID_CELL.h * 6 }, 300)).toEqual([120, 168])
      // A small box still keeps a row of margin either side.
      expect(gridWindow({ top: GRID_CELL.h * 10, width: GRID_CELL.w * 2, height: GRID_CELL.h * 2 }, 300)).toEqual([18, 26])
      const el = document.createElement('div')
      document.body.append(el)
      const root = renderInto(el, <SlideGrid count={60} current={0} hidden={() => false} thumb={() => { const c = document.createElement('canvas'); c.width = 320; c.height = 180; return c }} release={() => true} paint={() => {}} onJump={() => {}} />)
      const cells = [...el.querySelectorAll('[data-officeview-presenter-thumb]')]
      expect(cells.filter((c) => c.querySelector('canvas')).length).toBe(48)
      expect(cells[50]!.hasAttribute('data-unheld')).toBe(true)
      expect(cells[10]!.hasAttribute('data-unheld')).toBe(false)
      root.unmount()
    })

    it('M4: the presenter codes are classified explicitly', () => {
      expect(diagnosticKind({ code: 'presenter-paint', message: '' })).toBe('refused')
      expect(diagnosticKind({ code: 'presenter-thumbnail-approximate', message: '' })).toBe('approximate')
      expect(diagnosticKind({ code: 'presenter-styles-blocked', message: '' })).toBe('refused')
      expect(diagnosticKind({ code: 'placement-fullscreen-refused', message: '' })).toBe('refused')
    })

    it('S2-9: the panel groups agree with the diagnostic catalogue classes', () => {
      const BY_CLASS = { refused: 'refused', approximate: 'approximate', 'stand-in': 'approximate', info: 'other' } as const
      for (const [code, kind] of Object.entries(PANEL_KIND)) {
        // Every explicit entry is a catalogued code; only a resource code may be placed by presentation.
        expect(DIAGNOSTIC_CODES, code).toHaveProperty([code])
        const cls = DIAGNOSTIC_CODES[code]
        if (cls !== 'resource') expect(kind, code).toBe(BY_CLASS[cls])
      }
      // Every other catalogued code follows its class, whatever its wording.
      for (const [code, cls] of Object.entries(DIAGNOSTIC_CODES)) {
        if (cls === 'resource' || code in PANEL_KIND) continue
        expect(diagnosticKind({ code, message: '' }), code).toBe(BY_CLASS[cls])
      }
      // Wording that once misplaced a code: kiosk-restart is a refusal, not pending; timing-requires is a note.
      expect(diagnosticKind({ code: 'kiosk-restart', message: '' })).toBe('refused')
      expect(diagnosticKind({ code: 'timing-requires', message: '' })).toBe('other')
      expect(diagnosticKind({ code: 'media-codec', message: '' })).toBe('refused')
      expect(diagnosticKind({ code: 'constructor', message: '' })).toBe('other')
    })

    it('M5: swapping back hands focus to a focusable host', async () => {
      const r = await show(await deck(), undefined)
      r.host.removeAttribute('tabindex')
      key(r.host, 'F5', { altKey: true })
      await flush()
      key(r.host, 'F5', { altKey: true })
      await flush()
      expect(r.host.getAttribute('tabindex')).toBe('-1')
      expect(document.activeElement).toBe(r.host)
      r.s.exit()
    })
  })
})
