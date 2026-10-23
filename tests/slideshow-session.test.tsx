/**
 * Slideshow plan S1-19: the session controller, the navigation shell and the
 * audience surface (spec 4.1 exit and end-of-show, 4.2 session and starting
 * slide, 4.3 keys and precedence, 4.4 hidden/loop/kiosk/pause, 6.2 hidden
 * tab, 7 slide size, 8 lease fallback; checkpoint 1: static slides and cuts).
 *
 * jsdom. `runShow` is driven with a hand-built `ShowStart` (the eager start
 * is covered by slideshow-present.test.tsx), a ManualClock and a recording
 * painter injected through `sessionDeps`. Nothing waits on wall time.
 * Builder decks are "Office-shaped synthetic".
 */
import { describe, it, expect, vi, beforeEach, afterEach, type MockInstance } from 'vitest'
import JSZip from 'jszip'
import { runShow, viewerCallbacks } from '../src/slideshow'
import { firstSlide, sessionDeps, type SessionController } from '../src/slideshow/player/session'
import { createGestureState, gestureToInput, keyToInput } from '../src/slideshow/player/keys'
import { visibleOrder } from '../src/slideshow/player/order'
import { ManualClock } from '../src/slideshow/clock'
import type { Stage } from '../src/slideshow/player/stage'
import { createSlideShowWorkspace } from '../src/slideshow/workspace'
import { SlideShowError, type ShowStart, type SlideShowCallbacks } from '../src/components/present'
import { buildPptxDeck, mainSeqXml, timingXmlOf, type PptxDeckSpec, type PptxShapeSpec } from '../src/testdata/ooxml-builders'
import { OfficePackage } from '../src/core/zip'
import { parsePptx, pptxPackageHandle } from '../src/pptx/parse'
import { openShowPackage } from '../src/slideshow/read/deck'
import type { PptxDocument } from '../src/pptx/types'
import type { ShowIndex } from '../src/slideshow/types'

const W = 9144000, H = 6858000
function staticDoc(n: number, size: [number, number] = [W, H]): PptxDocument {
  return {
    slideWidthEmu: size[0], slideHeightEmu: size[1], images: [],
    slides: Array.from({ length: n }, (_, index) => ({ index, widthEmu: size[0], heightEmu: size[1], shapes: [] })),
  } as unknown as PptxDocument
}
const props = (body: string, attrs = '') => `<p:presentationPr xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main"><p:showPr ${attrs}>${body}</p:showPr></p:presentationPr>`

/** One click step with a 1 ms entrance on shape `spid`; ids from `base`. */
const click = (base: number, spid = '2') =>
  `<p:par><p:cTn id="${base}" fill="hold"><p:stCondLst><p:cond delay="indefinite"/></p:stCondLst><p:childTnLst>` +
  `<p:par><p:cTn id="${base + 1}" fill="hold"><p:stCondLst><p:cond delay="0"/></p:stCondLst><p:childTnLst>` +
  `<p:par><p:cTn id="${base + 2}" presetID="1" presetClass="entr" presetSubtype="0" fill="hold" grpId="0" nodeType="clickEffect"><p:stCondLst><p:cond delay="0"/></p:stCondLst><p:childTnLst>` +
  `<p:set><p:cBhvr><p:cTn id="${base + 3}" dur="1" fill="hold"><p:stCondLst><p:cond delay="0"/></p:stCondLst></p:cTn><p:tgtEl><p:spTgt spid="${spid}"/></p:tgtEl><p:attrNameLst><p:attrName>style.visibility</p:attrName></p:attrNameLst></p:cBhvr><p:to><p:strVal val="visible"/></p:to></p:set>` +
  `</p:childTnLst></p:cTn></p:par></p:childTnLst></p:cTn></p:par></p:childTnLst></p:cTn></p:par>`
const twoSteps = timingXmlOf(mainSeqXml([click(3), click(10)]))
const box: PptxShapeSpec = { id: 2, prst: 'rect', off: ['0', '0'], ext: ['914400', '914400'] }

async function deck(spec: PptxDeckSpec): Promise<{ doc: PptxDocument; bytes: Uint8Array }> {
  const bytes = await buildPptxDeck(spec)
  return { doc: await parsePptx(await OfficePackage.load(bytes)), bytes }
}

interface Rig {
  s: SessionController
  log: string[]
  changes: Array<[number, number]>
  host: HTMLElement
  clock: ManualClock
  cleanup: ReturnType<typeof vi.fn>
  ac: AbortController
  key(key: string, init?: KeyboardEventInit): void
}

let painted: number[]
/** S1-20: the engine seam. A recording stage stands in for it here (the engine has its own suite). */
const recordingStage = (): Stage => ({
  leave() {}, async enter(m) { painted.push(m.index) }, slideTime: (t) => t, frame: () => undefined, stateAt: () => new Map(),
  plan: undefined, toSlide: () => undefined, runAt: () => undefined, resize() {}, level: 'full', backend: 'canvas2d', backingStore: { w: 1, h: 1, dpr: 1 }, transition: undefined, inspect: () => ({ current: undefined, slides: [], lastDraws: [] }), dispose() {},
})
let clock: ManualClock
beforeEach(() => {
  painted = []
  clock = new ManualClock()
  vi.spyOn(sessionDeps, 'clock').mockImplementation(() => clock)
  vi.spyOn(sessionDeps, 'stage').mockImplementation(() => recordingStage())
  vi.spyOn(sessionDeps, 'frames').mockImplementation(() => () => {})
})
afterEach(() => {
  vi.restoreAllMocks()
  document.body.replaceChildren()
})

const flush = async () => { for (let i = 0; i < 20; i++) await new Promise((r) => setTimeout(r, 0)) }

function hostIn(doc: Document = document): HTMLElement {
  const host = doc.createElement('div')
  host.tabIndex = -1
  host.setAttribute('data-officeview-show', '')
  host.style.visibility = 'visible'
  doc.body.append(host)
  return host
}

async function show(doc: PptxDocument, over: Partial<ShowStart> = {}, cb: Partial<SlideShowCallbacks> = {}): Promise<Rig> {
  const host = over.host ?? hostIn()
  const ac = new AbortController()
  const log: string[] = []
  const changes: Array<[number, number]> = []
  const cleanup = vi.fn(() => { log.push('cleanup'); ac.abort(); host.style.visibility = 'hidden' })
  const start: ShowStart = {
    doc, handle: pptxPackageHandle(doc), host,
    token: { release() {}, released: false, owns: () => false },
    platform: { hasActivation: () => true, requestFullscreen: async () => {}, exitFullscreen: async () => {}, fullscreenElement: () => null, open: () => null },
    fullscreen: Promise.resolve('entered'),
    retryFullscreen: async () => 'entered',
    from: 'beginning', motion: 'auto',
    callbacks: {
      onStart: () => log.push('onStart'),
      onEnd: () => log.push('onEnd'),
      onSlideChange: (i, b) => changes.push([i, b]),
      onDiagnostic: (d) => log.push(`diag:${d.code}`),
      currentSlide: () => 0,
      restoreFocus: () => log.push('restoreFocus'),
      scrollToSlide: (i) => log.push(`scroll:${i}`),
      ...cb,
    },
    signal: ac.signal, generation: 1, isCurrent: () => !ac.signal.aborted, cleanup,
    ...over,
  }
  const s = await runShow(start)
  await flush()
  const key = (k: string, init: KeyboardEventInit = {}) => { host.dispatchEvent(new KeyboardEvent('keydown', { key: k, bubbles: true, cancelable: true, ...init })) }
  return { s, log, changes, host, clock, cleanup, ac, key }
}

async function press(r: Rig, ...keys: string[]) {
  for (const k of keys) { r.key(k); await flush() }
}

const k = (key: string, init: KeyboardEventInit = {}) => new KeyboardEvent('keydown', { key, ...init })

describe('keys and precedence', () => {
  it('key table: every next and previous key; Home/End; 12 Enter goes to slide 12; B and . black, W and , white; any key resumes; Esc exits; S and + pause', async () => {
    for (const key of [' ', 'ArrowRight', 'ArrowDown', 'PageDown', 'n', 'N']) expect(keyToInput(k(key), '')).toEqual({ kind: 'next' })
    expect(keyToInput(k('Enter'), '')).toEqual({ kind: 'enter' })
    for (const key of ['ArrowLeft', 'ArrowUp', 'PageUp', 'p', 'P', 'Backspace']) expect(keyToInput(k(key), '')).toEqual({ kind: 'previous' })
    expect(keyToInput(k('Home'), '')).toEqual({ kind: 'first' })
    expect(keyToInput(k('End'), '')).toEqual({ kind: 'last' })
    for (const key of ['b', 'B', '.']) expect(keyToInput(k(key), '')).toEqual({ kind: 'black' })
    for (const key of ['w', 'W', ',']) expect(keyToInput(k(key), '')).toEqual({ kind: 'white' })
    for (const key of ['s', 'S', '+']) expect(keyToInput(k(key), '')).toEqual({ kind: 'pause' })
    for (const key of ['Escape', '-']) expect(keyToInput(k(key), '12')).toEqual({ kind: 'exit' })
    expect(keyToInput(k('7'), '')).toEqual({ kind: 'digit', d: 7 })
    expect(keyToInput(k('Enter'), '12')).toEqual({ kind: 'goto', index: 11 })
    expect(keyToInput(k('x'), '')).toBeUndefined()
    expect(keyToInput(k('ArrowRight', { ctrlKey: true }), '')).toBeUndefined()

    const r = await show(staticDoc(14))
    expect(r.s.state.slide).toBe(0)
    for (const key of [' ', 'Enter', 'ArrowRight', 'ArrowDown', 'PageDown', 'n', 'N']) {
      const before = r.s.state.slide as number
      await press(r, key)
      expect(r.s.state.slide, key).toBe(before + 1)
    }
    for (const key of ['ArrowLeft', 'ArrowUp', 'PageUp', 'p', 'P', 'Backspace']) {
      const before = r.s.state.slide as number
      await press(r, key)
      expect(r.s.state.slide, key).toBe(before - 1)
    }
    await press(r, 'End')
    expect(r.s.state.slide).toBe(13)
    await press(r, 'Home')
    expect(r.s.state.slide).toBe(0)
    await press(r, '1', '2')
    expect(r.s.state.buffer).toBe('12')
    expect(r.s.state.slide).toBe(0)
    await press(r, 'Enter')
    expect(r.s.state.slide).toBe(11)
    expect(r.s.state.buffer).toBe('')
    expect(r.host.querySelector('[data-officeview-show-surface]')!.getAttribute('data-slide')).toBe('11')
    // black and white, and any key resumes without navigating
    for (const [key, screen] of [['b', 'black'], ['B', 'black'], ['.', 'black'], ['w', 'white'], ['W', 'white'], [',', 'white']] as const) {
      await press(r, key)
      expect(r.s.state.screen, key).toBe(screen)
      expect(r.host.querySelector(`[data-officeview-show-blank="${screen}"]`)).not.toBeNull()
      await press(r, key === 'b' ? 'x' : 'ArrowRight')
      expect(r.s.state.screen).toBe('slide')
      expect(r.s.state.slide).toBe(11)
    }
    // pause and resume
    for (const key of ['s', 'S', '+']) {
      const was = r.s.state.paused
      await press(r, key)
      expect(r.s.state.paused, key).toBe(!was)
      expect(r.clock.paused).toBe(!was)
    }
    await press(r, 's')
    expect(r.s.state.paused).toBe(false)
    // '-' exits
    await press(r, '-')
    await r.s.ended
    expect(r.cleanup).toHaveBeenCalledTimes(1)
    expect(r.log.slice(-4)).toEqual(['cleanup', 'scroll:11', 'restoreFocus', 'onEnd'])
  })

  it('Esc beats the number buffer and black screen', async () => {
    const r = await show(staticDoc(14))
    await press(r, 'b', '1', '2')
    expect(r.s.state).toMatchObject({ screen: 'black', buffer: '12', slide: 0 })
    await press(r, 'Escape')
    await r.s.ended
    expect(r.s.state.exited).toBe(true)
    expect(r.cleanup).toHaveBeenCalledTimes(1)
    expect(r.changes.some(([i]) => i === 11)).toBe(false)
    expect(r.log).toContain('scroll:0')
  })

  it('wheel down and swipe left go next; wheel up and swipe right go previous; a 30 px swipe does nothing; kiosk ignores both', async () => {
    const t = { now: 0 }
    const g = createGestureState(() => t.now)
    const wheel = (deltaY: number) => new WheelEvent('wheel', { deltaY })
    expect(gestureToInput(wheel(120), g)).toEqual({ kind: 'next' })
    // debounced: a second notch in the same instant is dropped
    expect(gestureToInput(wheel(120), g)).toBeUndefined()
    t.now += 1000
    expect(gestureToInput(wheel(-120), g)).toEqual({ kind: 'previous' })
    t.now += 1000
    // one step per 120 units: half a notch twice
    expect(gestureToInput(wheel(60), g)).toBeUndefined()
    t.now += 10
    expect(gestureToInput(wheel(60), g)).toEqual({ kind: 'next' })
    const touch = (type: string, x: number, y = 100) => Object.assign(new Event(type), { touches: type === 'touchend' ? [] : [{ clientX: x, clientY: y }], changedTouches: [{ clientX: x, clientY: y }] }) as unknown as TouchEvent
    const swipe = (from: number, to: number) => { gestureToInput(touch('touchstart', from), g); return gestureToInput(touch('touchend', to), g) }
    expect(swipe(300, 200)).toEqual({ kind: 'next' })
    expect(swipe(200, 300)).toEqual({ kind: 'previous' })
    expect(swipe(200, 170)).toBeUndefined()
    expect(swipe(200, 150)).toEqual({ kind: 'next' })

    // through the session
    const r = await show(staticDoc(5))
    r.host.dispatchEvent(new WheelEvent('wheel', { deltaY: 120, bubbles: true, cancelable: true }))
    await flush()
    expect(r.s.state.slide).toBe(1)
    r.clock.advance(1000)
    const sw = (from: number, to: number) => { r.host.dispatchEvent(touch('touchstart', from)); r.host.dispatchEvent(touch('touchend', to)) }
    sw(300, 200)
    await flush()
    expect(r.s.state.slide).toBe(2)
    sw(200, 300)
    await flush()
    expect(r.s.state.slide).toBe(1)
    r.clock.advance(1000)
    r.host.dispatchEvent(new WheelEvent('wheel', { deltaY: -120, bubbles: true, cancelable: true }))
    await flush()
    expect(r.s.state.slide).toBe(0)
    r.s.exit()

    const kiosk = await deck({ slides: [{ shapes: [] }, { shapes: [] }], presProps: props('<p:kiosk/>') })
    const rk = await show(kiosk.doc)
    expect(rk.s.state.kiosk).toBe(true)
    rk.host.dispatchEvent(new WheelEvent('wheel', { deltaY: 120, bubbles: true, cancelable: true }))
    rk.host.dispatchEvent(touch('touchstart', 300)); rk.host.dispatchEvent(touch('touchend', 100))
    await flush()
    expect(rk.s.state.slide).toBe(0)
    rk.s.exit()
  })
})

describe('order, hidden slides and loop', () => {
  it('next/previous skip hidden slides; goTo shows a hidden slide; current and number starts on a hidden slide', async () => {
    const index = (hidden: boolean[], range?: { st: number; end: number }): ShowIndex => ({
      slides: hidden.map((h, i) => ({ index: i, hidden: h })), diagnostics: [],
      settings: { loop: false, showAnimation: true, showNarration: false, useTimings: true, type: 'present', kioskRestartMs: 300000, ...(range ? { range } : {}) },
    })
    expect(visibleOrder(index([false, true, false, true, false]))).toEqual([0, 2, 4])
    expect(visibleOrder(index([false, false, true, false, false], { st: 1, end: 3 }))).toEqual([1, 3])
    expect(firstSlide({ index: index([true, false, false]) }, 'beginning')).toBe(1)
    expect(firstSlide({ index: index([false, false, false, false], { st: 2, end: 3 }) }, 'beginning')).toBe(2)
    expect(firstSlide({ index: index([false, true]) }, 1)).toBe(1)

    const { doc } = await deck({ slides: [{ shapes: [] }, { shapes: [], hidden: true }, { shapes: [] }, { shapes: [], hidden: true }, { shapes: [] }] })
    const r = await show(doc)
    expect(r.s.state.order).toEqual([0, 2, 4])
    await press(r, 'ArrowRight')
    expect(r.s.slideIndex).toBe(2)
    await press(r, 'ArrowRight')
    expect(r.s.slideIndex).toBe(4)
    await press(r, 'ArrowLeft')
    expect(r.s.slideIndex).toBe(2)
    r.s.goTo(1)
    await flush()
    expect(r.s.slideIndex).toBe(1)
    await press(r, '4', 'Enter')
    expect(r.s.slideIndex).toBe(3)
    await press(r, 'ArrowRight')
    expect(r.s.slideIndex).toBe(4)
    r.s.exit()

    const cur = await show(doc, { from: 'current' }, { currentSlide: () => 3 })
    expect(cur.s.slideIndex).toBe(3)
    await press(cur, 'ArrowLeft')
    expect(cur.s.slideIndex).toBe(2)
    cur.s.exit()
    const num = await show(doc, { from: 1 })
    expect(num.s.slideIndex).toBe(1)
    await press(num, 'ArrowRight')
    expect(num.s.slideIndex).toBe(2)
    num.s.exit()
  })

  it('PR #18 review 4: an empty deck started from the current slide opens on end-of-show', async () => {
    const { doc } = await deck({ slides: [] })
    expect(doc.slides).toHaveLength(0)
    const r = await show(doc, { from: 'current' }, { currentSlide: () => 0 })
    expect(r.s.state.slide).toBe('end')
    expect(r.log).toContain('onStart')
    r.s.exit()
  })

  it('PR #18 review 8: removing the full-screen notice (Dismiss, or a retry that enters) returns focus to the show host', async () => {
    const { doc } = await deck({ slides: [{ shapes: [] }, { shapes: [] }] })
    for (const which of ['Dismiss', 'Click for full screen']) {
      const r = await show(doc, { fullscreen: Promise.resolve('refused'), retryFullscreen: async () => 'entered' })
      await flush()
      const button = [...r.host.querySelectorAll<HTMLButtonElement>('[data-officeview-show-notice] button')].find(b => b.textContent === which)!
      button.focus()
      button.click()
      await flush()
      expect(r.host.querySelector('[data-officeview-show-notice]')).toBeNull()
      expect(document.activeElement).toBe(r.host)
      // Keys reach the show again.
      r.key('ArrowRight')
      await flush()
      expect(r.s.slideIndex).toBe(1)
      r.s.exit()
    }
  })

  it('a deck with every slide hidden opens on end-of-show', async () => {
    const { doc } = await deck({ slides: [{ shapes: [], hidden: true }, { shapes: [], hidden: true }] })
    const r = await show(doc)
    expect(r.s.state.slide).toBe('end')
    expect(r.log).toContain('onStart')
    const end = r.host.querySelector('[data-officeview-show-end]')!
    expect(end.textContent).toBe('End of slide show, click to exit.')
    expect((end as HTMLElement).hidden).toBe(false)
    expect(r.s.state.exited).toBe(false)
    // the next input exits
    r.host.querySelector('[data-officeview-show-surface]')!.dispatchEvent(new MouseEvent('click', { bubbles: true, clientX: 5, clientY: 5 }))
    await r.s.ended
    expect(r.cleanup).toHaveBeenCalledTimes(1)
    expect(r.log).toContain('scroll:0')
  })

  it('Previous from slide n+1 enters slide n at its last checkpoint; Next from n-1 enters it at initial', async () => {
    const { doc } = await deck({ slides: [{ shapes: [] }, { shapes: [box], timingXml: twoSteps }, { shapes: [] }] })
    const r = await show(doc)
    await press(r, 'ArrowRight')
    expect([r.s.slideIndex, r.s.buildIndex]).toEqual([1, 0])
    await press(r, 'ArrowRight')
    r.clock.advance(100)
    await press(r, 'ArrowRight')
    expect([r.s.slideIndex, r.s.buildIndex]).toEqual([1, 2])
    r.clock.advance(100)
    await press(r, 'ArrowRight')
    expect([r.s.slideIndex, r.s.buildIndex]).toEqual([2, 0])
    await press(r, 'ArrowLeft')
    expect([r.s.slideIndex, r.s.buildIndex]).toEqual([1, 2])
    await press(r, 'ArrowLeft')
    expect([r.s.slideIndex, r.s.buildIndex]).toEqual([1, 1])
    await press(r, 'ArrowLeft')
    expect([r.s.slideIndex, r.s.buildIndex]).toEqual([1, 0])
    await press(r, 'ArrowLeft')
    expect([r.s.slideIndex, r.s.buildIndex]).toEqual([0, 0])
    // inputs typed while a slide's metadata is still loading are queued, not lost
    r.key('ArrowRight'); r.key('ArrowRight')
    await flush()
    expect([r.s.slideIndex, r.s.buildIndex]).toEqual([1, 1])
    r.s.exit()
  })

  it('loop continues past the last visible slide without end-of-show', async () => {
    const { doc } = await deck({ slides: [{ shapes: [] }, { shapes: [] }, { shapes: [], hidden: true }], presProps: props('<p:present/>', 'loop="1"') })
    const r = await show(doc)
    expect(r.s.state.loop).toBe(true)
    await press(r, 'ArrowRight')
    expect(r.s.slideIndex).toBe(1)
    await press(r, 'ArrowRight')
    expect(r.s.state.slide).toBe(0)
    await press(r, 'ArrowRight', 'ArrowRight')
    expect(r.s.state.slide).toBe(0)
    expect(r.host.querySelector<HTMLElement>('[data-officeview-show-end]')!.hidden).toBe(true)
    r.s.exit()
    // without loop the same deck ends, and the end-of-show click exits
    const plain = await deck({ slides: [{ shapes: [] }, { shapes: [] }] })
    const e = await show(plain.doc)
    await press(e, 'ArrowRight', 'ArrowRight')
    expect(e.s.state.slide).toBe('end')
    expect(e.s.slideIndex).toBe(1)
    await press(e, 'ArrowLeft')
    expect(e.s.state.slide).toBe(1)
    await press(e, 'ArrowRight', 'ArrowRight')
    await e.s.ended
    expect(e.log).toContain('scroll:1')
  })

  it('kiosk ignores clicks and keys except Esc', async () => {
    const { doc } = await deck({ slides: [{ shapes: [] }, { shapes: [] }, { shapes: [] }], presProps: props('<p:kiosk/>') })
    const r = await show(doc)
    expect(r.s.state.kiosk).toBe(true)
    r.host.querySelector('[data-officeview-show-surface]')!.dispatchEvent(new MouseEvent('click', { bubbles: true }))
    await press(r, 'ArrowRight', ' ', 'End', 'b', 's', '2', 'Enter')
    expect(r.s.state).toMatchObject({ slide: 0, screen: 'slide', paused: false, buffer: '' })
    await press(r, 'Escape')
    await r.s.ended
    expect(r.cleanup).toHaveBeenCalledTimes(1)
  })
})

describe('clock, slide size and the package lease', () => {
  it('a hidden document pauses the clock', async () => {
    let state: DocumentVisibilityState = 'visible'
    const prop = Object.getOwnPropertyDescriptor(Document.prototype, 'visibilityState')
    Object.defineProperty(document, 'visibilityState', { configurable: true, get: () => state })
    try {
      const r = await show(staticDoc(2))
      r.clock.advance(10)
      expect(r.clock.now()).toBe(10)
      state = 'hidden'
      document.dispatchEvent(new Event('visibilitychange'))
      expect(r.s.state.clockPaused).toBe(true)
      r.clock.advance(1000)
      expect(r.clock.now()).toBe(10)
      state = 'visible'
      document.dispatchEvent(new Event('visibilitychange'))
      expect(r.s.state.clockPaused).toBe(false)
      r.clock.advance(5)
      expect(r.clock.now()).toBe(15)
      // a manual pause survives the document becoming visible again
      await press(r, 's')
      state = 'hidden'; document.dispatchEvent(new Event('visibilitychange'))
      state = 'visible'; document.dispatchEvent(new Event('visibilitychange'))
      expect(r.s.state.clockPaused).toBe(true)
      r.s.exit()
      // after exit the listener is gone
      state = 'hidden'; document.dispatchEvent(new Event('visibilitychange'))
    } finally {
      delete (document as { visibilityState?: unknown }).visibilityState
      if (prop) expect(Object.getOwnPropertyDescriptor(document, 'visibilityState')).toBeUndefined()
    }
  })

  it('zero, negative, non-finite and 1:21 slide sizes, including caller-constructed documents, are not-presentable', async () => {
    const parsed = (await deck({ slideSize: [914400 * 21, 914400], slides: [{ shapes: [] }] })).doc
    const docs: PptxDocument[] = [
      staticDoc(2, [0, H]), staticDoc(2, [W, 0]), staticDoc(2, [-W, H]), staticDoc(2, [NaN, H]), staticDoc(2, [W, Infinity]),
      staticDoc(2, [914400 * 21, 914400]), parsed,
      // a caller-constructed document whose slide disagrees with the deck
      { ...staticDoc(2), slides: [{ ...staticDoc(1).slides[0] }, { ...staticDoc(1, [0, 0]).slides[0], index: 1 }] } as PptxDocument,
    ]
    const workspace = vi.spyOn(sessionDeps, 'workspace')
    for (const doc of docs) {
      const host = hostIn()
      const log: string[] = []
      const cleanup = vi.fn()
      const err = await show(doc, { host, cleanup }, { onDiagnostic: (d) => log.push(d.code) }).then(() => undefined, (e: unknown) => e)
      expect(err).toBeInstanceOf(SlideShowError)
      expect((err as SlideShowError).code).toBe('not-presentable')
      expect(log).toEqual(['slide-size'])
      expect(host.childElementCount).toBe(0)
      expect(painted).toEqual([])
    }
    expect(workspace).not.toHaveBeenCalled()
  })

  it('a handle-less document presents static slides with no-package', async () => {
    const { doc } = await deck({ slides: [{ shapes: [box], timingXml: twoSteps }, { shapes: [] }] })
    expect(pptxPackageHandle(doc)).toBeDefined()
    const clone = structuredClone(doc)
    expect(pptxPackageHandle(clone)).toBeUndefined()
    const r = await show(clone)
    expect(r.log).toContain('diag:no-package')
    expect(r.log.indexOf('diag:no-package')).toBeLessThan(r.log.indexOf('onStart'))
    expect(painted).toContain(0)
    // static: the first Next leaves the slide (no click steps), as a cut
    await press(r, 'ArrowRight')
    expect([r.s.slideIndex, r.s.buildIndex]).toEqual([1, 0])
    expect(painted).toEqual([0, 1])
    r.s.exit()
  })

  it('a re-readable source is reopened when the handle is missing; a one-shot source is not', async () => {
    const { doc, bytes } = await deck({ slides: [{ shapes: [] }, { shapes: [], hidden: true }, { shapes: [] }] })
    const clone = structuredClone(doc)
    const reopen = vi.spyOn(sessionDeps, 'openShowPackage')
    const r = await show(clone, { handle: undefined, source: bytes })
    expect(reopen).toHaveBeenCalledTimes(1)
    expect(r.log).not.toContain('diag:no-package')
    expect(r.s.state.order).toEqual([0, 2]) // read from the reopened package
    r.s.exit()

    reopen.mockClear()
    const oneShot = await show(clone, { handle: undefined, source: new Response(bytes as BodyInit) as unknown as Blob })
    expect(reopen).not.toHaveBeenCalled()
    expect(oneShot.log).toContain('diag:no-package')
    expect(oneShot.s.state.order).toEqual([0, 1, 2])
    oneShot.s.exit()

    // a reopen that fails (a cookie-authenticated URL, say) falls back to no-package
    reopen.mockRejectedValueOnce(new Error('HTTP 403'))
    const failed = await show(clone, { handle: undefined, source: 'https://example.invalid/deck.pptx' })
    expect(failed.log).toEqual(expect.arrayContaining(['diag:package-reopen', 'diag:no-package']))
    failed.s.exit()
  })

  it('a slide nested deeper than the XML parser allows degrades to a static slide with a diagnostic', async () => {
    // The eager parser refuses such a deck outright, so the show reads a deeper
    // slide than the document was parsed from: the same deck with one slide's
    // timing nested 40 time-node levels deep.
    let inner = '<p:par><p:cTn id="900" dur="100"/></p:par>'
    for (let i = 0; i < 40; i++) inner = `<p:par><p:cTn id="${100 + i}" fill="hold"><p:childTnLst>${inner}</p:childTnLst></p:cTn></p:par>`
    const shallow = await deck({ slides: [{ shapes: [box], timingXml: twoSteps }, { shapes: [box], timingXml: twoSteps }] })
    const deep = await buildPptxDeck({ slides: [{ shapes: [box], timingXml: twoSteps }, { shapes: [box], timingXml: timingXmlOf(mainSeqXml([inner])) }] })
    await expect(OfficePackage.load(deep).then(parsePptx)).rejects.toThrow(/nested/i)
    const r = await show(shallow.doc, { handle: await openShowPackage(deep) })
    await press(r, 'ArrowRight', 'ArrowRight')
    r.clock.advance(100)
    await press(r, 'ArrowRight')
    expect([r.s.slideIndex, r.s.buildIndex]).toEqual([1, 0])
    expect(r.log).toContain('diag:slide-refused')
    // static: one Next leaves it
    await press(r, 'ArrowRight')
    expect(r.s.state.slide).toBe('end')
    expect(r.s.state.exited).toBe(false)
    r.s.exit()
  })
})

describe('model assembly and the static surface', () => {
  it('layout and master artwork reach the resolver, and a transition on a layout or master is diagnosed as not inherited', async () => {
    const master = { id: 9, prst: 'rect', off: ['0', '0'], ext: ['914400', '914400'] } as PptxShapeSpec
    const bytes = await buildPptxDeck({
      masters: [{ shapes: [master], layouts: [{}] }],
      slides: [{ layout: [0, 0], shapes: [box], timingXml: timingXmlOf(mainSeqXml([click(3, '9'), click(10, '2')])) }],
    })
    const zip = await JSZip.loadAsync(bytes)
    const layout = 'ppt/slideLayouts/slideLayout1.xml'
    const xml = await zip.file(layout)!.async('string')
    expect(xml).toMatch(/<\/p:sldLayout>$/)
    zip.file(layout, xml.replace(/<\/p:sldLayout>$/, '<p:transition><p:fade/></p:transition></p:sldLayout>'))
    const patched = await zip.generateAsync({ type: 'uint8array' })
    const doc = await parsePptx(await OfficePackage.load(patched))
    const r = await show(doc)
    expect(r.log).toContain('diag:transition-inherited')
    // spid 9 lives only on the master: master-only, not missing
    expect(r.log).toContain('diag:master-target')
    expect(r.log).not.toContain('diag:unresolved-target')
    expect(r.log).not.toContain('diag:underlay')
    // the dropped behaviour's step still counts; the other step plays
    await press(r, 'ArrowRight')
    r.clock.advance(100)
    await press(r, 'ArrowRight')
    expect(r.s.buildIndex).toBe(2)
    r.s.exit()
  })

  it('the static painter paints the slide letterboxed into the output canvas (terminal path)', async () => {
    // S1-20: the engine's terminal rung, reached with a workspace that holds the permanent reserve and little else.
    ;(sessionDeps.stage as unknown as MockInstance).mockRestore()
    vi.spyOn(sessionDeps, 'workspace').mockImplementation(() => createSlideShowWorkspace('desktop', { limit: 1024 * 768 * 4 + 8 * 2 ** 20 + 4096, scratch: 8 * 2 ** 20 }))
    const { doc } = await deck({ slides: [{ shapes: [{ off: ['914400', '914400'], ext: ['7315200', '1828800'], fill: '336699' }] }, { shapes: [{ off: ['0', '0'], ext: ['9144000', '6858000'], fill: 'CC3300' }] }] })
    const r = await show(doc)
    const canvas = r.host.querySelector<HTMLCanvasElement>('[data-officeview-show-canvas]')!
    const ctx = canvas.getContext('2d')!
    // 960 x 720 slide px fitted to the output backing store
    const scale = Math.min(canvas.width / 960, canvas.height / 720)
    const ox = (canvas.width - 960 * scale) / 2, oy = (canvas.height - 720 * scale) / 2
    const at = (x: number, y: number) => [...ctx.getImageData(Math.round(ox + x * scale), Math.round(oy + y * scale), 1, 1).data]
    expect(at(480, 192)).toEqual([0x33, 0x66, 0x99, 255])
    expect(at(480, 600)).toEqual([255, 255, 255, 255])
    // a cut to the next slide repaints it
    await press(r, 'ArrowRight')
    r.s.tick()
    expect(at(480, 600)).toEqual([0xcc, 0x33, 0x00, 255])
    expect(r.log).not.toContain('diag:paint-failed')
    r.s.exit()
  })
})

describe('exit, focus and events', () => {
  it('focus returns to the start element or the viewer root; the viewer scrolls to the last slide', async () => {
    const root = document.createElement('div')
    const region = document.createElement('div')
    region.setAttribute('data-officeview-root', '')
    region.tabIndex = 0
    let scrollTop = 0
    Object.defineProperty(region, 'scrollTop', { configurable: true, get: () => scrollTop, set: (v: number) => { scrollTop = v } })
    region.getBoundingClientRect = () => ({ top: 100, height: 800, bottom: 900 }) as DOMRect
    for (let i = 0; i < 3; i++) {
      const c = document.createElement('canvas')
      c.getBoundingClientRect = () => ({ top: i * 316 - scrollTop, bottom: i * 316 + 300 - scrollTop }) as DOMRect
      region.append(c)
    }
    const host = document.createElement('div')
    host.setAttribute('data-officeview-show', '')
    host.tabIndex = -1
    const back = document.createElement('button')
    root.append(region, host)
    document.body.append(back, root)
    back.focus()

    const r = await show(staticDoc(3), { host }, viewerCallbacks(host, back))
    host.focus()
    r.s.goTo(2)
    r.s.exit()
    await r.s.ended
    expect(document.activeElement).toBe(back)
    expect(scrollTop).toBe(632 - 100 - 16)

    // the start element is gone (an API-started show): focus goes to the viewer root
    scrollTop = 0
    back.remove()
    const r2 = await show(staticDoc(3), { host }, viewerCallbacks(host, back))
    await press(r2, 'ArrowRight')
    await press(r2, 'Escape')
    await r2.s.ended
    expect(document.activeElement).toBe(region)
    expect(scrollTop).toBe(316 - 100 - 16)

    // nothing was focused (the body): the viewer root again
    region.blur()
    const r3 = await show(staticDoc(3), { host }, viewerCallbacks(host, document.body))
    r3.s.exit()
    await r3.s.ended
    expect(document.activeElement).toBe(region)
  })

  it('onSlideShowSlideChange reports slide and build index', async () => {
    const { doc } = await deck({ slides: [{ shapes: [box], timingXml: twoSteps }, { shapes: [] }] })
    const r = await show(doc)
    await press(r, 'ArrowRight')
    r.clock.advance(100)
    await press(r, 'ArrowRight')
    r.clock.advance(100)
    await press(r, 'ArrowRight', 'ArrowLeft')
    expect(r.changes).toEqual([[0, 0], [0, 1], [0, 2], [1, 0], [0, 2]])
    // abort (the eager start's signal) exits exactly once
    r.ac.abort()
    await r.s.ended
    r.s.exit()
    expect(r.log.filter((l) => l === 'onEnd')).toHaveLength(1)
    expect(r.cleanup).toHaveBeenCalledTimes(1)
  })

  it('PR #18 review 2: ended settles even when the host\'s end callback throws', async () => {
    const { doc } = await deck({ slides: [{ shapes: [] }] })
    const r = await show(doc, {}, { onEnd: () => { throw new Error('host onSlideShowEnd failed') } })
    expect(() => r.s.exit()).toThrow('host onSlideShowEnd failed')
    let settled = false
    void r.s.ended.then(() => { settled = true })
    await flush()
    expect(settled).toBe(true)
    expect(r.cleanup).toHaveBeenCalledTimes(1)
  })
})

describe('PR #19 review 2', () => {
  it('Presenter View primes previews with the stage\'s expanded text timing', async () => {
    const textTiming = vi.fn(() => undefined)
    vi.spyOn(sessionDeps, 'stage').mockImplementation(() => ({ ...recordingStage(), textTiming }))
    const r = await show(staticDoc(2))
    const opts = r.s.presenterSource.previewOptions
    expect(typeof opts.textTiming).toBe('function')
    const meta = { index: 1, hidden: false, timing: undefined, notes: [], links: [], media: [], builds: [] }
    const charge = () => true
    opts.textTiming!(1, meta, charge)
    expect(textTiming).toHaveBeenCalledWith(1, meta, charge)
    r.s.exit()
  })

  it('the transition sound starts when the stage starts the transition, not before the slide is drawable', async () => {
    // Slide 2's transition sound names the audio part on its shape (jsdom decodes no audio, so starting it is observed
    // as its media-codec diagnostic, emitted by the transition-sound controller at start).
    const sound = '<p:transition><p:fade/><p:sndAc><p:stSnd><p:snd r:embed="rIdShow1" name="a.wav"/></p:stSnd></p:sndAc></p:transition>'
    const { doc } = await deck({ slides: [{ shapes: [box] }, { shapes: [box], media: [{ shapeId: 2, kind: 'audio', ext: 'wav', data: new Uint8Array([1, 2, 3]) }], transition: { xml: sound } }] })
    let ready!: () => void
    const entered: number[] = []
    vi.spyOn(sessionDeps, 'stage').mockImplementation(() => ({
      ...recordingStage(),
      async enter(m) {
        entered.push(m.index)
        // Slide 2 is not drawable yet (images decoding, layers building): the transition has not started.
        if (m.index === 1) await new Promise<void>((res) => { ready = res })
      },
    }))
    const r = await show(doc)
    r.s.next()
    await flush()
    expect(entered).toEqual([0, 1])
    expect(r.log.filter((l) => l === 'diag:media-codec')).toEqual([])
    ready()
    await flush()
    expect(r.log.filter((l) => l === 'diag:media-codec')).toEqual(['diag:media-codec'])
    r.s.exit()
  })
})
