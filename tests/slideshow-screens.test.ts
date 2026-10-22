/**
 * Slideshow Task S2-3: Window Management placement for Presenter View.
 * Spec: §5.1 placement in Chromium, other browsers, denial diagnostics.
 *
 * Placement writes S2-2's eager `preparedPlacement` slot (present.ts), which
 * the synchronous start reads; "Use both screens" and "Swap displays" act on
 * a `PlacementTarget` (the provisional popup plus `ShowStart.retryFullscreen`).
 */
import { afterEach, describe, expect, it, vi } from 'vitest'
import {
  applyPlacement,
  createDragHint,
  diagnosePlacement,
  grantedScreenDetails,
  DRAG_HINT_STORAGE_KEY,
  planPlacement,
  preparePlacement,
  prepareScreens,
  requestScreens,
  shouldShowDragHint,
  showDragHint,
  swapDisplays,
  useBothScreens,
  type PlacementTarget,
} from '../src/slideshow/presenter/screens'
import { preparedPlacement, type ScreenDetails, type ScreenLike } from '../src/components/present'
import { createDiagnosticSink } from '../src/slideshow/diagnostics'

const screen = (s: Partial<ScreenLike> & { availLeft: number; availWidth: number }): ScreenLike => ({ availTop: 0, availHeight: 1040, ...s })
const PRIMARY = screen({ isPrimary: true, isInternal: true, availLeft: 0, availWidth: 1920, label: 'primary' })
const EXTERNAL = screen({ isPrimary: false, isInternal: false, availLeft: 1920, availWidth: 2560, availHeight: 1440, label: 'external' })
const TWO: ScreenDetails = { screens: [PRIMARY, EXTERNAL], currentScreen: PRIMARY }

function fakeStorage(): Storage {
  const data = new Map<string, string>()
  return {
    getItem: (k) => data.get(k) ?? null,
    setItem: (k, v) => void data.set(k, String(v)),
    removeItem: (k) => void data.delete(k),
    clear: () => data.clear(),
    key: (i) => [...data.keys()][i] ?? null,
    get length() { return data.size },
  }
}

/** A popup stand-in recording moves; `retry` resolves to the given outcome. */
function target(outcome: 'entered' | 'refused' | 'reject' | 'throw' = 'entered', state?: PlacementTarget['state']) {
  const popup = { closed: false, moveTo: vi.fn(), resizeTo: vi.fn() }
  const retryFullscreen = vi.fn((_o?: { screen?: ScreenLike }) => {
    if (outcome == 'throw') throw new Error('sync throw')
    return outcome == 'reject' ? Promise.reject(new Error('x')) : Promise.resolve(outcome)
  })
  const t: PlacementTarget = { popup: popup as unknown as Window, retryFullscreen, state }
  return { t, popup, retryFullscreen }
}

afterEach(() => {
  vi.restoreAllMocks()
})

describe('Slideshow Screen Placement (Task S2-3)', () => {
  it('granted with two screens: popup on primary, audience requestFullscreen({screen}) on the external screen', () => {
    const plan = planPlacement(TWO)
    expect(plan).toEqual({ popupFeatures: 'popup=yes,left=0,top=0,width=1920,height=1040', presenterScreen: PRIMARY, audienceScreen: EXTERNAL })
    // The plan lands in S2-2's eager slot, which the synchronous start reads:
    // the popup opens with these features and full screen goes to `audienceScreen`.
    const doc = document.implementation.createHTMLDocument('opener')
    preparePlacement(doc, TWO)
    expect(preparedPlacement.get(doc)).toEqual({ popupFeatures: 'popup=yes,left=0,top=0,width=1920,height=1040', audienceScreen: EXTERNAL })
    // A later hint with no usable second screen (unplugged) drops the stale plan.
    preparePlacement(doc, { screens: [PRIMARY] })
    expect(preparedPlacement.has(doc)).toBe(false)
    preparePlacement(doc, TWO)
    preparePlacement(doc, undefined)
    expect(preparedPlacement.has(doc)).toBe(false)
  })

  it('primary not first, isPrimary absent, three screens and negative offsets', () => {
    // [external, primary]: the primary is found by isPrimary, not position.
    expect(planPlacement({ screens: [EXTERNAL, PRIMARY] })).toMatchObject({ presenterScreen: PRIMARY, audienceScreen: EXTERNAL })
    // No isPrimary: the screen the opener is on presents; else the first.
    const a = screen({ availLeft: 0, availWidth: 1000 }), b = screen({ availLeft: 1000, availWidth: 1000 })
    expect(planPlacement({ screens: [a, b], currentScreen: b })).toMatchObject({ presenterScreen: b, audienceScreen: a })
    expect(planPlacement({ screens: [a, b] })).toMatchObject({ presenterScreen: a, audienceScreen: b })
    // Three screens: an external (isInternal false) beats an internal second panel, whatever the order.
    const internal2 = screen({ isInternal: true, availLeft: 1920, availWidth: 1280 })
    expect(planPlacement({ screens: [PRIMARY, internal2, EXTERNAL] })).toMatchObject({ audienceScreen: EXTERNAL })
    // Negative offsets pass through unchanged.
    const left = screen({ isPrimary: true, availLeft: -2560, availTop: -200, availWidth: 2560, availHeight: 1400 })
    expect(planPlacement({ screens: [PRIMARY, left] })).toMatchObject({ popupFeatures: 'popup=yes,left=0,top=0,width=1920,height=1040' })
    expect(planPlacement({ screens: [left, PRIMARY] })).toMatchObject({ popupFeatures: 'popup=yes,left=-2560,top=-200,width=2560,height=1400', audienceScreen: PRIMARY })
    // A screen with no usable size is not a screen: no fabricated geometry.
    expect(planPlacement({ screens: [PRIMARY, {} as ScreenLike] })).toEqual({ manual: true, reason: 'single-screen' })
    expect(planPlacement({ screens: [PRIMARY, screen({ availLeft: 5, availWidth: NaN })] })).toEqual({ manual: true, reason: 'single-screen' })
  })

  it('not yet granted: start proceeds and Use both screens asks on click', async () => {
    // The preload hint only reads granted details: ungranted means no plan, so
    // 'auto' starts single-screen (controller ruling) and nothing is diagnosed.
    const doc = document.implementation.createHTMLDocument('opener')
    await prepareScreens(doc, () => Promise.resolve(undefined))
    expect(preparedPlacement.has(doc)).toBe(false)
    // A throwing details reader degrades the same way.
    await prepareScreens(doc, () => Promise.reject(new Error('no permissions API')))
    expect(preparedPlacement.has(doc)).toBe(false)
    await prepareScreens(doc, () => Promise.resolve(TWO))
    expect(preparedPlacement.get(doc)?.audienceScreen).toBe(EXTERNAL)

    // The click: getScreenDetails runs synchronously, inside the click's activation.
    const opener = document.implementation.createHTMLDocument('viewer')
    const getScreenDetails = vi.fn(() => Promise.resolve(TWO))
    const win = { getScreenDetails, document: opener } as unknown as Window
    const { t, popup, retryFullscreen } = target('entered')
    const pending = useBothScreens(win, t)
    expect(getScreenDetails).toHaveBeenCalledTimes(1)
    const result = await pending
    // ...and then places the windows: audience full screen on the external
    // screen first, the popup moved to the primary only once that entered.
    expect(result).toMatchObject({ outcome: 'placed' })
    expect(retryFullscreen).toHaveBeenCalledWith({ screen: EXTERNAL })
    expect(popup.moveTo).toHaveBeenCalledWith(0, 0)
    expect(popup.resizeTo).toHaveBeenCalledWith(1920, 1040)
    expect(t.state).toEqual({ presenter: PRIMARY, audience: EXTERNAL })
    // The grant prepares the next start too.
    expect(preparedPlacement.get(opener)?.audienceScreen).toBe(EXTERNAL)

    // Refused (the click was in the popup, the opener had no activation): nothing moves.
    const refused = target('refused')
    expect(await useBothScreens(win, refused.t)).toMatchObject({ outcome: 'refused' })
    expect(refused.popup.moveTo).not.toHaveBeenCalled()
    expect(refused.t.state).toBeUndefined()
    // A later gesture retries with the same plan.
    const plan = planPlacement(TWO)
    if ('manual' in plan) throw new Error('expected a plan')
    const later = target('entered')
    expect(await applyPlacement(later.t, plan.presenterScreen, plan.audienceScreen)).toBe('placed')
    expect(later.popup.moveTo).toHaveBeenCalledWith(0, 0)
  })

  it('the preload hint reads granted details from the document\'s own window and never prompts', async () => {
    const getScreenDetails = vi.fn(() => Promise.resolve(TWO))
    const win = (state: string | Error) => ({
      navigator: { permissions: { query: vi.fn(() => state instanceof Error ? Promise.reject(state) : Promise.resolve({ state })) } },
      getScreenDetails,
    }) as unknown as Window
    expect(await grantedScreenDetails(win('granted'))).toBe(TWO)
    expect(await grantedScreenDetails(win('prompt'))).toBeUndefined()
    expect(await grantedScreenDetails(win('denied'))).toBeUndefined()
    expect(getScreenDetails).toHaveBeenCalledTimes(1)
    // Firefox/Safari: unknown permission name rejects; no navigator.permissions at all; no window.
    expect(await grantedScreenDetails(win(new TypeError('bad name')))).toBeUndefined()
    expect(await grantedScreenDetails({ navigator: {} } as unknown as Window)).toBeUndefined()
    expect(await grantedScreenDetails(null)).toBeUndefined()
    // prepareScreens defaults to the document's own window (realm-safe).
    const doc = document.implementation.createHTMLDocument('popup-realm')
    Object.defineProperty(doc, 'defaultView', { value: win('granted') })
    await prepareScreens(doc)
    expect(preparedPlacement.get(doc)?.audienceScreen).toBe(EXTERNAL)
  })

  it('denied, missing API and single screen are three distinct diagnostics, none is popup-blocked', async () => {
    const sink = createDiagnosticSink()
    const denied = Object.assign(new Error('denied'), { name: 'NotAllowedError' })
    const cases: Array<[string, Window]> = [
      ['permission-denied', { getScreenDetails: () => Promise.reject(denied) } as unknown as Window],
      ['api-missing', {} as Window],
      ['single-screen', { getScreenDetails: () => Promise.resolve({ screens: [PRIMARY] }) } as unknown as Window],
      ['failed', { getScreenDetails: () => Promise.reject(Object.assign(new Error('frame'), { name: 'SecurityError' })) } as unknown as Window],
      ['failed', { get getScreenDetails(): never { throw new Error('getter') } } as unknown as Window],
    ]
    for (const [reason, win] of cases) {
      const { t, popup, retryFullscreen } = target('entered')
      expect(await useBothScreens(win, t, sink)).toEqual({ reason })
      expect(retryFullscreen).not.toHaveBeenCalled()
      expect(popup.moveTo).not.toHaveBeenCalled()
    }
    expect(await requestScreens({} as Window)).toEqual({ reason: 'api-missing' })
    const codes = sink.list().map((d) => d.code)
    expect(codes).toEqual(['placement-permission-denied', 'placement-api-missing', 'placement-single-screen', 'placement-failed'])
    expect(codes).not.toContain('popup-blocked')
    expect(sink.list().every((d) => d.message.length > 0)).toBe(true)
    // diagnosePlacement maps a manual plan as well.
    const other = createDiagnosticSink()
    diagnosePlacement(planPlacement({ screens: [] }), other)
    diagnosePlacement(planPlacement(TWO), other)
    expect(other.list().map((d) => d.code)).toEqual(['placement-single-screen'])
  })

  it('swap displays exchanges the screens', async () => {
    const { t, popup, retryFullscreen } = target('entered', { presenter: PRIMARY, audience: EXTERNAL })
    expect(await swapDisplays(t)).toBe('placed')
    // Full screen is requested first, on the old presenter screen...
    expect(retryFullscreen).toHaveBeenCalledWith({ screen: PRIMARY })
    // ...and only then the popup moves to the old audience screen (outer size).
    expect(popup.moveTo).toHaveBeenCalledWith(1920, 0)
    expect(popup.resizeTo).toHaveBeenCalledWith(2560, 1440)
    expect(t.state).toEqual({ presenter: EXTERNAL, audience: PRIMARY })
    expect(retryFullscreen.mock.invocationCallOrder[0]).toBeLessThan(popup.moveTo.mock.invocationCallOrder[0])

    // Negative-offset target screen.
    const left = screen({ availLeft: -2560, availTop: -200, availWidth: 2560, availHeight: 1400 })
    const neg = target('entered', { presenter: left, audience: PRIMARY })
    await swapDisplays(neg.t)
    expect(neg.popup.moveTo).toHaveBeenCalledWith(0, 0)
    await swapDisplays(neg.t)
    expect(neg.popup.moveTo).toHaveBeenLastCalledWith(-2560, -200)
  })

  it('a refused, rejecting or throwing swap leaves everything in place and is diagnosed', async () => {
    const unhandled = vi.fn()
    process.on('unhandledRejection', unhandled)
    try {
      for (const outcome of ['refused', 'reject', 'throw'] as const) {
        const sink = createDiagnosticSink()
        const state = { presenter: PRIMARY, audience: EXTERNAL }
        const { t, popup } = target(outcome, state)
        expect(await swapDisplays(t, sink)).toBe('refused')
        expect(t.state).toBe(state)
        expect(t.state).toEqual({ presenter: PRIMARY, audience: EXTERNAL })
        expect(popup.moveTo).not.toHaveBeenCalled()
        expect(sink.list().map((d) => d.code)).toEqual(['placement-fullscreen-refused'])
      }
      // No state (never placed) and a closed popup: nothing to do, no throw.
      expect(await swapDisplays(target('entered').t)).toBe('refused')
      const closed = target('entered', { presenter: PRIMARY, audience: EXTERNAL })
      ;(closed.popup as { closed: boolean }).closed = true
      expect(await swapDisplays(closed.t)).toBe('placed')
      expect(closed.popup.moveTo).not.toHaveBeenCalled()
      // A popup whose moveTo throws still counts as placed (the popup is best effort).
      const throwing = target('entered', { presenter: PRIMARY, audience: EXTERNAL })
      throwing.popup.moveTo.mockImplementation(() => { throw new Error('blocked') })
      expect(await swapDisplays(throwing.t)).toBe('placed')
      await new Promise((r) => setTimeout(r, 0))
      expect(unhandled).not.toHaveBeenCalled()
    } finally {
      process.off('unhandledRejection', unhandled)
    }
  })

  it('other browsers show the one-time drag hint', () => {
    const storage = fakeStorage()
    const doc = document.implementation.createHTMLDocument('Presenter')
    expect(shouldShowDragHint(storage)).toBe(true)
    const hint = createDragHint(doc, { storage })!
    expect(hint.ownerDocument).toBe(doc)
    expect(hint.getAttribute('role')).toBe('status')
    expect(hint.textContent).toContain('Drag this presenter window to your other display')
    hint.querySelector('button')!.click()
    expect(storage.getItem(DRAG_HINT_STORAGE_KEY)).toBe('1')
    expect(shouldShowDragHint(storage)).toBe(false)
    expect(showDragHint(doc.createElement('div'), { storage })).toBeUndefined()
  })

  it('the drag hint works on a popup (other-realm) document and stays dismissed', () => {
    // An iframe document is another realm: `instanceof Document` is false there.
    const frame = document.createElement('iframe')
    document.body.append(frame)
    try {
      const popupDoc = frame.contentDocument!
      expect(popupDoc instanceof Document).toBe(false)
      localStorage.removeItem(DRAG_HINT_STORAGE_KEY)
      const hint = showDragHint(popupDoc)!
      expect(hint).toBeDefined()
      expect(hint.parentNode).toBe(popupDoc.body)
      hint.querySelector('button')!.click()
      expect(popupDoc.querySelector('[data-officeview-drag-hint]')).toBeNull()
      // Persisted in the origin's localStorage (a popup's sessionStorage copy dies with it).
      expect(localStorage.getItem(DRAG_HINT_STORAGE_KEY)).toBe('1')
      expect(showDragHint(popupDoc)).toBeUndefined()
      expect(showDragHint(document)).toBeUndefined()
    } finally {
      frame.remove()
      localStorage.removeItem(DRAG_HINT_STORAGE_KEY)
    }
  })

  it('blocked storage degrades: the hint still shows and dismissal does not throw', () => {
    const doc = document.implementation.createHTMLDocument('Sandboxed')
    const win = {}
    Object.defineProperty(win, 'localStorage', { get() { throw Object.assign(new Error('blocked'), { name: 'SecurityError' }) } })
    Object.defineProperty(doc, 'defaultView', { value: win })
    const hint = showDragHint(doc)!
    expect(hint).toBeDefined()
    expect(() => hint.querySelector('button')!.click()).not.toThrow()
    const throwing = { getItem() { throw new Error('x') }, setItem() { throw new Error('x') } } as unknown as Storage
    expect(shouldShowDragHint(throwing)).toBe(true)
    expect(() => createDragHint(doc, { storage: throwing })!.querySelector('button')!.click()).not.toThrow()
    expect(showDragHint(doc.createElement('div').ownerDocument.createElement('div'), { storage: null })).toBeDefined()
  })
})
