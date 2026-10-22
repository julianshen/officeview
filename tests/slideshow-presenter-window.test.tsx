/**
 * Slideshow plan S2-2: the Presenter View popup lifecycle (spec 4.1 steps 1-2
 * and popup refusal, 5.1 second window, closing, single-screen fallback, 5.3,
 * 7 popup boundary, 9 Presenter View focus).
 *
 * jsdom. The popup is a fake window over a second document (an EventTarget
 * with `document`, `closed` and `close()`; closing fires `pagehide`, as a real
 * popup does). The eager start runs with an injected PresentPlatform and a
 * deferred chunk; the session runs through `runShow` with a recording stage.
 * Builder decks are "Office-shaped synthetic". Nothing waits on wall time.
 */
import { describe, it, expect, vi, beforeEach, afterEach } from 'vitest'
import { runShow } from '../src/slideshow'
import { sessionDeps } from '../src/slideshow/player/session'
import { adoptPresenterWindow } from '../src/slideshow/presenter/window'
import { createDiagnosticSink } from '../src/slideshow/diagnostics'
import { ManualClock } from '../src/slideshow/clock'
import type { Stage } from '../src/slideshow/player/stage'
import * as slideshow from '../src/slideshow'
import { acquireShowToken, defaultPlatform, preparedPlacement, showLoader, startPresent, SlideShowError, type PresentPlatform, type ShowStart, type SlideShowCallbacks } from '../src/components/present'
import { buildPptxDeck } from '../src/testdata/ooxml-builders'
import { OfficePackage } from '../src/core/zip'
import { parsePptx, pptxPackageHandle } from '../src/pptx/parse'
import type { PptxDocument } from '../src/pptx/types'

type FakePopup = Window & { closeCalls: number }
function fakePopup(): FakePopup {
  const t = new EventTarget() as unknown as FakePopup & { closed: boolean }
  Object.assign(t, {
    document: document.implementation.createHTMLDocument(''),
    closed: false,
    closeCalls: 0,
    close() {
      t.closeCalls++
      if (t.closed) return
      t.closed = true
      t.dispatchEvent(new Event('pagehide'))
    },
  })
  return t
}

const W = 9144000, H = 6858000
const box = { id: 2, prst: 'rect', off: ['0', '0'] as [string, string], ext: ['914400', '914400'] as [string, string] }
const SCRIPT = 'Say <script>window.__pwned = 1</script> & <b>bold</b>'
async function notesDeck(): Promise<PptxDocument> {
  const bytes = await buildPptxDeck({ slides: [{ shapes: [box], notes: [[SCRIPT], ['Second paragraph']] }, { shapes: [box] }, { shapes: [box] }] })
  return parsePptx(await OfficePackage.load(bytes))
}

const recordingStage = (): Stage => ({
  leave() {}, async enter() {}, slideTime: (t) => t, frame() {}, stateAt: () => new Map(),
  plan: undefined, toSlide: () => undefined, resize() {}, level: 'full', backend: 'canvas2d', backingStore: { w: 1, h: 1, dpr: 1 }, transition: undefined, inspect: () => ({ current: undefined, slides: [], lastDraws: [] }), dispose() {},
})
beforeEach(() => {
  vi.spyOn(sessionDeps, 'clock').mockImplementation(() => new ManualClock())
  vi.spyOn(sessionDeps, 'stage').mockImplementation(() => recordingStage())
  vi.spyOn(sessionDeps, 'frames').mockImplementation(() => () => {})
})
afterEach(() => {
  vi.restoreAllMocks()
  document.body.replaceChildren()
  document.head.replaceChildren()
})

const flush = async () => { for (let i = 0; i < 20; i++) await new Promise((r) => setTimeout(r, 0)) }

async function show(doc: PptxDocument, popup: Window | null | undefined) {
  const host = document.createElement('div')
  host.tabIndex = -1
  document.body.append(host)
  const ac = new AbortController()
  const log: string[] = []
  const cleanup = vi.fn(() => { log.push('cleanup'); ac.abort(); popup?.close() })
  const start: ShowStart = {
    doc, handle: pptxPackageHandle(doc), host,
    token: { release() {}, released: false, owns: () => false },
    platform: { hasActivation: () => true, requestFullscreen: async () => {}, exitFullscreen: async () => {}, fullscreenElement: () => null, open: () => null },
    fullscreen: Promise.resolve('entered'),
    retryFullscreen: async () => 'entered',
    from: 'beginning', motion: 'auto', presenterView: popup === undefined ? undefined : true, popup,
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

// ---------------------------------------------------------------------------
// The eager prologue
// ---------------------------------------------------------------------------

interface Deferred<T> { promise: Promise<T>; resolve(v: T): void; reject(e: unknown): void }
function deferred<T>(): Deferred<T> {
  let resolve!: (v: T) => void, reject!: (e: unknown) => void
  const promise = new Promise<T>((a, b) => { resolve = a; reject = b })
  return { promise, resolve, reject }
}
const PPTX = { slides: [0, 1].map((index) => ({ index, widthEmu: W, heightEmu: H, shapes: [] })), slideWidthEmu: W, slideHeightEmu: H, images: [] } as unknown as PptxDocument

function platform(log: string[], popup: Window | null): PresentPlatform {
  return {
    hasActivation: () => true,
    requestFullscreen: (_el, o) => { log.push(`requestFullscreen${o?.screen ? `:${o.screen.label}` : ''}`); return Promise.resolve() },
    exitFullscreen: () => Promise.resolve(),
    fullscreenElement: () => null,
    open: (_w, url, features) => { log.push(`open:${url}:${features}`); return popup },
  }
}
function freshHost(): HTMLElement {
  const doc = document.implementation.createHTMLDocument('t')
  const host = doc.createElement('div')
  host.style.visibility = 'hidden'
  doc.body.append(host)
  return host
}
const callbacks = (): Partial<SlideShowCallbacks> => ({})

describe('presenter popup: the synchronous start', () => {
  let loads: Deferred<typeof slideshow>[]
  let log: string[]
  beforeEach(() => {
    log = []
    loads = []
    vi.spyOn(showLoader, 'load').mockImplementation(() => {
      log.push('import')
      const d = deferred<typeof slideshow>()
      loads.push(d)
      return d.promise
    })
  })

  it('popup is opened before requestFullscreen in the same synchronous prologue', async () => {
    const host = freshHost()
    const popup = fakePopup()
    const run = startPresent({ doc: PPTX, host, opts: { presenterView: true }, platform: platform(log, popup), generation: 1, callbacks: callbacks(), allow: true, motion: 'auto' })
    // Everything privileged happened before the first await.
    expect(log).toEqual(['open:about:blank:popup', 'requestFullscreen', 'import'])
    host.dispatchEvent(new KeyboardEvent('keydown', { key: 'Escape' }))
    await run.catch(() => {})

    // A prepared placement (the preload hint) gives the popup's features and the audience screen, still synchronously.
    log.length = 0
    const host2 = freshHost()
    preparedPlacement.set(host2.ownerDocument, { popupFeatures: 'popup,left=0,top=0,width=800,height=600', audienceScreen: { availWidth: 1920, availHeight: 1080, label: 'external' } })
    const run2 = startPresent({ doc: PPTX, host: host2, platform: platform(log, fakePopup()), generation: 1, callbacks: callbacks(), allow: true, motion: 'auto' })
    expect(log).toEqual(['open:about:blank:popup,left=0,top=0,width=800,height=600', 'requestFullscreen:external', 'import'])
    host2.dispatchEvent(new KeyboardEvent('keydown', { key: 'Escape' }))
    await run2.catch(() => {})

    // presenterView: false, 'auto' with no prepared screen, and 'auto' on an extended display without a
    // prepared screen (permission not granted) open nothing: a plain start goes straight to full screen.
    log.length = 0
    for (const [presenterView, ext] of [[false, false], ['auto', false], ['auto', true], [undefined, true]] as const) {
      Object.defineProperty(window.screen, 'isExtended', { value: ext, configurable: true })
      const h = document.createElement('div')
      document.body.append(h)
      const r = startPresent({ doc: PPTX, host: h, opts: { presenterView }, platform: platform(log, fakePopup()), generation: 1, callbacks: callbacks(), allow: true, motion: 'auto' })
      h.dispatchEvent(new KeyboardEvent('keydown', { key: 'Escape' }))
      await r.catch(() => {})
      h.remove()
    }
    delete (window.screen as { isExtended?: boolean }).isExtended
    expect(log.filter((l) => l.startsWith('open'))).toEqual([])
    expect(log.filter((l) => l.startsWith('requestFullscreen'))).toHaveLength(4)
  })

  it('a throwing platform open is a blocked popup: the start goes on and the show token is released at its end (review M4)', async () => {
    const host = freshHost()
    const p = platform(log, null)
    p.open = () => { log.push('open:throws'); throw new Error('no popups here') }
    const run = startPresent({ doc: PPTX, host, opts: { presenterView: true }, platform: p, generation: 1, callbacks: callbacks(), allow: true, motion: 'auto' })
    expect(log).toEqual(['open:throws', 'requestFullscreen', 'import'])
    host.dispatchEvent(new KeyboardEvent('keydown', { key: 'Escape' }))
    expect(await run.then(() => 'resolved', (x: SlideShowError) => x.code)).toBe('cancelled')
    const token = acquireShowToken(host)
    expect(token).toBeDefined()
    token!.release()
  })

  it('the opener unloading during load closes the provisional popup', async () => {
    const host = document.createElement('div')
    document.body.append(host)
    const popup = fakePopup()
    const run = startPresent({ doc: PPTX, host, opts: { presenterView: true }, platform: platform(log, popup), generation: 1, callbacks: callbacks(), allow: true, motion: 'auto' })
    expect(popup.closed).toBe(false)
    window.dispatchEvent(new Event('pagehide'))
    expect(popup.closed).toBe(true)
    host.dispatchEvent(new KeyboardEvent('keydown', { key: 'Escape' }))
    expect(await run.then(() => 'resolved', (x: SlideShowError) => x.code)).toBe('cancelled')
    // The listener goes with the start: a later pagehide no longer reaches the popup.
    const calls = popup.closeCalls
    window.dispatchEvent(new Event('pagehide'))
    expect(popup.closeCalls).toBe(calls)
    host.remove()
  })

  it('a failed load closes the provisional popup', async () => {
    const host = freshHost()
    const popup = fakePopup()
    const run = startPresent({ doc: PPTX, host, opts: { presenterView: true }, platform: platform(log, popup), generation: 1, callbacks: callbacks(), allow: true, motion: 'auto' })
    expect(popup.closed).toBe(false)
    loads[0].reject(new Error('chunk failed'))
    const e = await run.catch((x: unknown) => x)
    expect(e).toBeInstanceOf(SlideShowError)
    expect((e as SlideShowError).code).toBe('load-failed')
    expect(popup.closed).toBe(true)

    // Cancelled during load (Esc) closes it too.
    const popup2 = fakePopup()
    const host2 = freshHost()
    const run2 = startPresent({ doc: PPTX, host: host2, opts: { presenterView: true }, platform: platform(log, popup2), generation: 1, callbacks: callbacks(), allow: true, motion: 'auto' })
    host2.dispatchEvent(new KeyboardEvent('keydown', { key: 'Escape' }))
    expect(await run2.then(() => 'resolved', (x: SlideShowError) => x.code)).toBe('cancelled')
    expect(popup2.closed).toBe(true)
  })
})

describe('defaultPlatform.screenDetails (review I1)', () => {
  const nav = navigator as Navigator & { permissions?: Permissions }
  const win = window as Window & { getScreenDetails?: () => Promise<unknown> }
  let saved: PropertyDescriptor | undefined
  beforeEach(() => { saved = Object.getOwnPropertyDescriptor(nav, 'permissions') })
  afterEach(() => {
    if (saved) Object.defineProperty(nav, 'permissions', saved)
    else delete (nav as { permissions?: unknown }).permissions
    delete win.getScreenDetails
  })
  const permissions = (query: () => unknown) => Object.defineProperty(nav, 'permissions', { value: { query }, configurable: true })

  it('resolves undefined, never rejects, where Window Management is missing, unknown, not granted or failing', async () => {
    Object.defineProperty(nav, 'permissions', { value: undefined, configurable: true })
    await expect(defaultPlatform.screenDetails!()).resolves.toBeUndefined()
    // Firefox and Safari: the permission name is not a PermissionName (a synchronous or asynchronous TypeError).
    permissions(() => { throw new TypeError("'window-management' is not a valid enum value") })
    await expect(defaultPlatform.screenDetails!()).resolves.toBeUndefined()
    permissions(() => Promise.reject(new TypeError('unknown permission')))
    await expect(defaultPlatform.screenDetails!()).resolves.toBeUndefined()
    permissions(() => Promise.resolve({ state: 'prompt' }))
    win.getScreenDetails = vi.fn()
    await expect(defaultPlatform.screenDetails!()).resolves.toBeUndefined()
    expect(win.getScreenDetails).not.toHaveBeenCalled()
    // Granted, but the API itself is missing or fails.
    permissions(() => Promise.resolve({ state: 'granted' }))
    delete win.getScreenDetails
    await expect(defaultPlatform.screenDetails!()).resolves.toBeUndefined()
    win.getScreenDetails = () => Promise.reject(new Error('NotAllowedError'))
    await expect(defaultPlatform.screenDetails!()).resolves.toBeUndefined()
  })

  it('returns the screen details when the permission is already granted', async () => {
    const details = { screens: [{ availWidth: 1, availHeight: 1 }] }
    permissions(() => Promise.resolve({ state: 'granted' }))
    win.getScreenDetails = () => Promise.resolve(details)
    await expect(defaultPlatform.screenDetails!()).resolves.toBe(details)
  })
})

// ---------------------------------------------------------------------------
// The session and the presenter window
// ---------------------------------------------------------------------------

describe('presenter popup: the session', () => {
  it('null or already-closed popup is blocked: single-screen with popup-blocked', async () => {
    const doc = await notesDeck()
    const blocked = await show(doc, null)
    expect(blocked.log).toContain('diag:popup-blocked')
    expect(blocked.log.indexOf('diag:popup-blocked')).toBeLessThan(blocked.log.indexOf('onStart'))
    expect(blocked.s.presenter).toBe('swap')
    blocked.s.exit()

    const closed = fakePopup()
    closed.close()
    const r = await show(doc, closed)
    expect(r.log).toContain('diag:popup-blocked')
    expect(r.s.presenter).toBe('swap')
    expect(closed.document.body.childElementCount).toBe(0)
    r.s.exit()

    // Not wanted: no popup, no diagnostic.
    const none = await show(doc, undefined)
    expect(none.s.presenter).toBe('none')
    expect(none.log.filter((l) => l.startsWith('diag:'))).toEqual([])
    none.s.exit()
  })

  it('closing the presenter window keeps the show in single-screen', async () => {
    const popup = fakePopup()
    const r = await show(await notesDeck(), popup)
    expect(r.s.presenter).toBe('window')
    let ended = false
    void r.s.ended.then(() => { ended = true })
    popup.close()
    await flush()
    expect(r.s.presenter).toBe('swap')
    expect(ended).toBe(false)
    expect(r.cleanup).not.toHaveBeenCalled()
    // The audience still navigates.
    r.s.next()
    expect(r.s.slideIndex).toBe(1)
    r.s.exit()
    await flush()
    expect(ended).toBe(true)
  })

  it('exit and opener pagehide close the popup', async () => {
    const popup = fakePopup()
    const r = await show(await notesDeck(), popup)
    expect(popup.closed).toBe(false)
    r.s.exit()
    expect(popup.closed).toBe(true)

    // The window on its own: the opener's pagehide closes it; close() is idempotent and onClosed fires once.
    const p2 = fakePopup()
    const opener = new EventTarget() as unknown as Window
    Object.assign(opener, { document })
    const pw = adoptPresenterWindow(p2, opener, createDiagnosticSink())
    const closedCalls = vi.fn()
    pw.onClosed(closedCalls)
    opener.dispatchEvent(new Event('pagehide'))
    expect(p2.closed).toBe(true)
    pw.close()
    expect(closedCalls).toHaveBeenCalledTimes(1)
  })

  it('notes and labels are text nodes; deck text with <script> renders literally', async () => {
    const popup = fakePopup()
    const r = await show(await notesDeck(), popup)
    const pdoc = popup.document
    expect(pdoc.querySelector('[data-ov-theme-root]')).not.toBeNull()
    expect(pdoc.querySelector('script, b')).toBeNull()
    const notes = pdoc.querySelector('[data-officeview-presenter-notes]')!
    expect(notes.textContent).toContain(SCRIPT)
    // Every note paragraph holds a single text node.
    for (const p of notes.querySelectorAll('p')) {
      expect([...p.childNodes].every((n) => n.nodeType === Node.TEXT_NODE)).toBe(true)
    }
    expect((window as { __pwned?: number }).__pwned).toBeUndefined()
    // Labelled controls, and they work: both views move together.
    const button = (name: string) => pdoc.querySelector<HTMLButtonElement>(`button[aria-label="${name}"]`)!
    expect(button('Previous slide')).not.toBeNull()
    expect(button('End show')).not.toBeNull()
    expect(pdoc.querySelector('[data-officeview-presenter-status]')!.textContent).toBe('Slide 1 of 3')
    button('Next slide').click()
    await flush()
    expect(r.s.slideIndex).toBe(1)
    expect(pdoc.querySelector('[data-officeview-presenter-status]')!.textContent).toBe('Slide 2 of 3')
    // Keys on the popup's own document: the same table as the audience.
    pdoc.body.dispatchEvent(new KeyboardEvent('keydown', { key: 'ArrowLeft', bubbles: true, cancelable: true }))
    await flush()
    expect(r.s.slideIndex).toBe(0)
    pdoc.body.dispatchEvent(new KeyboardEvent('keydown', { key: 'Escape', bubbles: true, cancelable: true }))
    expect(r.log).toContain('onEnd')
    expect(popup.closed).toBe(true)
  })

  it('keys typed or scrolled in the speaker notes stay there, except Esc (review I2)', async () => {
    const popup = fakePopup()
    const r = await show(await notesDeck(), popup)
    const notes = popup.document.querySelector<HTMLElement>('[data-officeview-presenter-notes]')!
    const key = (target: Element, k: string) => {
      const e = new KeyboardEvent('keydown', { key: k, bubbles: true, cancelable: true })
      target.dispatchEvent(e)
      return e.defaultPrevented
    }
    for (const k of ['ArrowDown', 'ArrowRight', 'PageDown', ' ', 'Enter', 'n', 'b']) expect(key(notes, k)).toBe(false)
    // A paragraph inside the notes counts as the notes.
    expect(key(notes.querySelector('p')!, 'ArrowUp')).toBe(false)
    await flush()
    expect(r.s.slideIndex).toBe(0)
    expect(r.s.state.screen).toBe('slide')
    // Outside the notes the audience keys still apply.
    key(popup.document.body, 'ArrowRight')
    await flush()
    expect(r.s.slideIndex).toBe(1)
    key(notes, 'Escape')
    expect(r.log).toContain('onEnd')
  })

  it('a popup that cannot be adopted is closed and the show runs single-screen (best effort)', async () => {
    const popup = fakePopup()
    Object.defineProperty(popup, 'document', { get() { throw new Error('denied') } })
    const r = await show(await notesDeck(), popup)
    expect(r.log).toContain('diag:presenter-window-failed')
    expect(r.log).toContain('onStart')
    expect(r.s.presenter).toBe('swap')
    expect(popup.closed).toBe(true)
    r.s.exit()
  })

  it('a refused presenter stylesheet is diagnosed once and the controls stay usable', async () => {
    const style = document.createElement('style')
    style.textContent = '[data-ov-theme-root] { color: red }'
    document.head.append(style)
    const popup = fakePopup()
    const r = await show(await notesDeck(), popup)
    const copy = popup.document.head.querySelector('style')
    expect(copy?.textContent).toContain('color: red')
    const violation = new Event('securitypolicyviolation', { bubbles: true }) as Event & { effectiveDirective: string }
    Object.defineProperty(violation, 'effectiveDirective', { value: 'style-src-elem' })
    copy!.dispatchEvent(violation)
    copy!.dispatchEvent(new Event('error'))
    expect(r.log.filter((l) => l === 'diag:presenter-styles-blocked')).toHaveLength(1)
    popup.document.querySelector<HTMLButtonElement>('button[aria-label="Next slide"]')!.click()
    await flush()
    expect(r.s.slideIndex).toBe(1)
    r.s.exit()
  })
})
