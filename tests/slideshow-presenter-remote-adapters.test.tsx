/**
 * Streaming Stage 3 piece 3, Task 5: Presenter View over an image-backed
 * slide source (a streamed deck: no PptxDocument, no ShowModel — finished
 * server-marked slide bitmaps and a server-drawn notes image).
 *
 * jsdom with node-canvas. Presenter View is rendered directly with a
 * hand-built session and source (no `runShow`: there is no deck), over a
 * real SlideShowWorkspace. Nothing waits on wall time.
 */
import { describe, it, expect, vi, afterEach } from 'vitest'
import { createRoot } from 'react-dom/client'
import { flushSync } from 'react-dom'
import { PresenterView, type PresenterSource } from '../src/slideshow/presenter/PresenterView'
import { NotesPane, type ImageNotes } from '../src/slideshow/presenter/notes'
import { createElapsedTimer } from '../src/slideshow/presenter/timer'
import { createSlideShowWorkspace } from '../src/slideshow/workspace'
import type { SessionController, SessionState } from '../src/slideshow/player/session'
import type { NoteParagraph, SlideShowDiagnostic } from '../src/slideshow/types'

afterEach(() => {
  vi.restoreAllMocks()
  document.body.replaceChildren()
})

const flush = async () => { for (let i = 0; i < 10; i++) await new Promise((r) => setTimeout(r, 0)) }

/** Top-left pixel of a canvas, as [r, g, b, a]. */
const px = (c: HTMLCanvasElement, x = 0, y = 0) => [...c.getContext('2d')!.getImageData(x, y, 1, 1).data]

interface Remote {
  source: PresenterSource
  session: SessionController
  slideImage: ReturnType<typeof vi.fn>
  previewOf: ReturnType<typeof vi.fn>
  bump(): void
  setSlide(n: number | 'end'): void
  ready: Set<number>
  notes: { current: ImageNotes | readonly NoteParagraph[] }
}

/** A streamed deck of `count` slides; `hidden` deck indices are flagged. Ready slides paint solid red. */
function remote(count: number, opts: { hidden?: readonly number[]; ready?: readonly number[] } = {}): Remote {
  const subs = new Set<() => void>()
  let version = 0
  const ready = new Set(opts.ready ?? Array.from({ length: count }, (_, i) => i))
  const notes: Remote['notes'] = { current: [] }
  const slideImage = vi.fn((i: number, canvas: HTMLCanvasElement) => {
    if (!ready.has(i)) return false
    const ctx = canvas.getContext('2d')!
    ctx.fillStyle = '#f00'
    ctx.fillRect(0, 0, canvas.width, canvas.height)
    return true
  })
  const previewOf = vi.fn(() => { throw new Error('previewOf must not be called for an image-backed source') })
  const workspace = createSlideShowWorkspace('desktop')
  const sink = { add: vi.fn(), list: (): readonly SlideShowDiagnostic[] => [] }
  const order = Array.from({ length: count }, (_, i) => i).filter((i) => !opts.hidden?.includes(i))
  let state: SessionState = {
    slide: 0, build: 0, screen: 'slide', buffer: '', paused: false, clockPaused: false, kiosk: false, loop: false, order, exited: false,
  }
  const timer = createElapsedTimer(0)
  const source = {
    subscribe(cb: () => void) { subs.add(cb); return () => { subs.delete(cb) } },
    get version() { return version },
    slideCount: count,
    slideImage,
    hidden: (i: number) => !!opts.hidden?.includes(i),
    notes: () => notes.current,
    diagnostics: () => [],
    sink,
    timer,
    now: () => 0,
    wall: () => 0,
    slideTime: () => 0,
    previewOptions: { showAnimation: true, reducedMotion: false, sink },
    workspace,
    protection: { copy: true, print: true },
    audienceFrame: () => undefined,
    onFrame: () => () => {},
    previewOf,
    key() {},
    input() {},
    toggleView() {},
    placement: undefined,
    touch() {},
  } as unknown as PresenterSource
  const session = {
    get state() { return state },
    cursor: undefined,
    presenterSource: source,
    goTo: vi.fn(), next: vi.fn(), previous: vi.fn(), exit: vi.fn(),
  } as unknown as SessionController
  const bump = () => { version++; for (const cb of subs) cb() }
  return {
    source, session, slideImage, previewOf, ready, notes, bump,
    setSlide(n) { state = { ...state, slide: n }; bump() },
  }
}

function mount(r: Remote) {
  const el = document.createElement('div')
  document.body.append(el)
  const root = createRoot(el)
  flushSync(() => root.render(<PresenterView session={r.session} host="swap" doc={document} />))
  return { el, root }
}

const nextCanvas = (el: HTMLElement) => el.querySelector<HTMLCanvasElement>('[data-officeview-presenter-next] canvas')!

describe('Presenter View over an image-backed source (Task 5)', () => {
  it('counts slides from slideCount and paints Next with slideImage(after), never previewOf', async () => {
    const r = remote(5)
    const { el, root } = mount(r)
    await flush()
    expect(el.querySelector('[data-officeview-presenter-status]')!.textContent).toBe('Slide 1 of 5')
    expect(r.slideImage).toHaveBeenCalledWith(1, expect.any(HTMLCanvasElement))
    expect(px(nextCanvas(el))).toEqual([255, 0, 0, 255])
    expect(el.querySelector('[data-officeview-presenter-next] figcaption')!.textContent).toBe('Next slide (2)')

    flushSync(() => r.setSlide(3))
    await flush()
    expect(el.querySelector('[data-officeview-presenter-status]')!.textContent).toBe('Slide 4 of 5')
    expect(r.slideImage).toHaveBeenLastCalledWith(4, expect.any(HTMLCanvasElement))

    // The last slide: Next is the end of the show, nothing painted.
    flushSync(() => r.setSlide(4))
    await flush()
    expect(el.querySelector('[data-officeview-presenter-end]')).not.toBeNull()
    expect(r.previewOf).not.toHaveBeenCalled()
    root.unmount()
  })

  it('skips a hidden slide for Next (the order) and shows a neutral placeholder until the bitmap arrives', async () => {
    const r = remote(4, { hidden: [1], ready: [0] })
    const { el, root } = mount(r)
    await flush()
    // Next skips hidden slide 1.
    expect(r.slideImage).toHaveBeenCalledWith(2, expect.any(HTMLCanvasElement))
    expect(r.slideImage).not.toHaveBeenCalledWith(1, expect.anything())
    // Not ready: the preview is cleared to the black the local letterbox paints, never a stale bitmap.
    const c = nextCanvas(el)
    c.getContext('2d')!.fillStyle = '#0f0'
    c.getContext('2d')!.fillRect(0, 0, c.width, c.height)
    flushSync(() => r.bump())
    expect(px(c)).toEqual([0, 0, 0, 255])
    // The bitmap lands: the next version change repaints it.
    r.ready.add(2)
    flushSync(() => r.bump())
    expect(px(c)).toEqual([255, 0, 0, 255])
    expect(r.previewOf).not.toHaveBeenCalled()
    root.unmount()
  })

  it('the grid paints thumbnails with slideImage(i), leaves hidden slides out, and retries the ones not ready', async () => {
    const r = remote(4, { hidden: [2], ready: [0, 1] })
    const { el, root } = mount(r)
    await flush()
    r.slideImage.mockClear()
    flushSync(() => (el.querySelector('button[aria-label="See all slides"]') as HTMLButtonElement).click())
    await flush()
    const cells = [...el.querySelectorAll<HTMLButtonElement>('[data-officeview-presenter-grid] li button')]
    expect(cells.map((b) => b.dataset.slide)).toEqual(['0', '1', '3'])
    const thumbCalls = () => r.slideImage.mock.calls.filter(([, c]) => (c as HTMLCanvasElement).closest('[data-officeview-presenter-grid]')).map(([i]) => i)
    expect(thumbCalls().sort()).toEqual([0, 1, 3])
    const thumb3 = cells[2]!.querySelector('canvas')!
    expect(px(thumb3)).toEqual([0, 0, 0, 255])

    // Only the thumbnail that was not ready is retried on the next version change.
    r.slideImage.mockClear()
    r.ready.add(3)
    flushSync(() => r.bump())
    expect(thumbCalls()).toEqual([3])
    expect(px(thumb3)).toEqual([255, 0, 0, 255])
    r.slideImage.mockClear()
    flushSync(() => r.bump())
    expect(thumbCalls()).toEqual([])

    // A jump goes to the deck index.
    flushSync(() => cells[2]!.click())
    expect(r.session.goTo).toHaveBeenCalledWith(3)
    expect(r.previewOf).not.toHaveBeenCalled()
    root.unmount()
  })

  it('the go-to field is bounded by slideCount', async () => {
    const r = remote(3)
    const { el, root } = mount(r)
    await flush()
    const field = el.querySelector<HTMLInputElement>('input[aria-label="Go to slide"]')!
    const go = (v: string) => { field.value = v; field.dispatchEvent(new KeyboardEvent('keydown', { key: 'Enter', bubbles: true })) }
    go('4')
    expect(r.session.goTo).not.toHaveBeenCalled()
    go('3')
    expect(r.session.goTo).toHaveBeenCalledWith(2)
    root.unmount()
  })

  it('image notes: status text while loading or failed, a painted unselectable canvas when ready', async () => {
    const r = remote(2)
    const paint = vi.fn((c: HTMLCanvasElement) => {
      c.width = 40
      c.height = 20
      const ctx = c.getContext('2d')!
      ctx.fillStyle = '#00f'
      ctx.fillRect(0, 0, 40, 20)
    })
    r.notes.current = { kind: 'image', state: 'loading', paint }
    const { el, root } = mount(r)
    await flush()
    const region = el.querySelector<HTMLElement>('[data-officeview-presenter-notes]')!
    expect(region.textContent).toBe('Loading notes…')
    expect(paint).not.toHaveBeenCalled()
    // No text-size controls for an image.
    expect(el.querySelector('button[aria-label="Larger notes text"]')).toBeNull()

    r.notes.current = { kind: 'image', state: 'ready', paint }
    flushSync(() => r.bump())
    const canvas = el.querySelector<HTMLCanvasElement>('[data-officeview-presenter-notes] canvas')!
    expect(canvas.getAttribute('aria-label')).toBe('Speaker notes (image)')
    expect(canvas.getAttribute('role')).toBe('img')
    expect(canvas.getAttribute('draggable')).toBe('false')
    expect(paint).toHaveBeenCalledTimes(1)
    expect(paint).toHaveBeenCalledWith(canvas)
    expect(px(canvas)).toEqual([0, 0, 255, 255])
    // Not selectable, not copyable, whatever the copy policy says.
    expect(el.querySelector<HTMLElement>('[data-officeview-presenter-notes]')!.style.userSelect).toBe('none')
    const menu = new MouseEvent('contextmenu', { bubbles: true, cancelable: true })
    canvas.dispatchEvent(menu)
    expect(menu.defaultPrevented).toBe(true)
    const drag = new Event('dragstart', { bubbles: true, cancelable: true })
    canvas.dispatchEvent(drag)
    expect(drag.defaultPrevented).toBe(true)
    // The same notes object across renders paints once.
    flushSync(() => r.bump())
    expect(paint).toHaveBeenCalledTimes(1)

    r.notes.current = { kind: 'image', state: 'error', paint }
    flushSync(() => r.bump())
    expect(el.querySelector('[data-officeview-presenter-notes] canvas')).toBeNull()
    expect(el.querySelector('[data-officeview-presenter-notes]')!.textContent).toBe('Notes could not be loaded')
    root.unmount()
  })

  it('image notes display at the server\'s 720 CSS px, not their device-pixel backing (DPR 2 → 1,440 px), and fit a narrower pane (PR #75)', async () => {
    const r = remote(2)
    r.notes.current = { kind: 'image', state: 'ready', paint: (c) => { c.width = 1440; c.height = 400 } }
    const { el, root } = mount(r)
    await flush()
    const canvas = el.querySelector<HTMLCanvasElement>('[data-officeview-presenter-notes] canvas')!
    expect(canvas.width).toBe(1440)
    expect(canvas.style.maxWidth).toBe('720px')
    expect(canvas.style.width).toBe('100%')
    expect(canvas.style.height).toBe('auto')
    root.unmount()
  })

  it('NotesPane keeps the text variant unchanged', () => {
    const el = document.createElement('div')
    document.body.append(el)
    const root = createRoot(el)
    flushSync(() => root.render(<NotesPane notes={[{ level: 0, text: 'Hello' }]} selectable />))
    expect(el.querySelector('[data-officeview-presenter-notes] p')!.textContent).toBe('Hello')
    expect(el.querySelector('button[aria-label="Larger notes text"]')).not.toBeNull()
    root.unmount()
  })
})
