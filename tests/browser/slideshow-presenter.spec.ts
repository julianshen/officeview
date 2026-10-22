/**
 * Slideshow plan S2-2: the Presenter View popup lifecycle in Chromium (the
 * `slideshow` Playwright project), driven through src/debug/slideshow-start.tsx
 * (`presenter=1`, `frame=sandbox`) and the CSP-hardened
 * src/debug/slideshow-csp.html.
 *
 * Asserts behaviour only (popups, full-screen element, diagnostics, slide
 * changes), never pixels or timings. Waits poll observable state.
 */
import { expect, test, type Frame, type Page } from '@playwright/test'
import { fileURLToPath } from 'node:url'
import { checkIdentity, gitIdentity } from '../../scripts/harness/identity'

const ROOT = fileURLToPath(new URL('../..', import.meta.url))
const BASE = `http://127.0.0.1:${process.env.IV_HARNESS_PORT ?? 4173}`
const PAGE = '/src/debug/slideshow-start.html'
const CSP_PAGE = '/src/debug/slideshow-csp.html'

const errors: string[] = []

test.beforeAll(async () => {
  const text = await (await fetch(`${BASE}/__harness-id`)).text()
  let body: unknown = text
  try { body = JSON.parse(text) } catch { /* checkIdentity rejects raw text */ }
  expect(checkIdentity(body, { commit: gitIdentity(ROOT).commit, root: ROOT }), 'harness server identity').toBeNull()
})

test.describe('S2-2 presenter window lifecycle', () => {
  test.beforeEach(({ page }) => {
    errors.length = 0
    page.on('pageerror', (e) => errors.push(e.message))
  })
  test.afterEach(() => {
    expect(errors).toEqual([])
  })

  async function open(page: Page, path: string): Promise<void> {
    await page.goto(path)
    await page.waitForFunction(() => window.__slideshowStart?.ready === true)
    await expect(page.locator('[data-officeview-show]').first()).toBeAttached()
  }
  const log = (page: Page | Frame) => page.evaluate(() => [...window.__slideshowStart!.log])
  const fullscreen = (page: Page | Frame) => page.evaluate(() => window.__slideshowStart!.fullscreenElement())
  const presenter = (page: Page) => page.evaluate(() => window.__slideshowStart!.sessions.a?.presenter)

  /** Clicks Present with Presenter View and returns the popup. */
  async function startWithPopup(page: Page): Promise<Page> {
    const popupEvent = page.waitForEvent('popup')
    await page.click('#present-a')
    const popup = await popupEvent
    await expect.poll(() => log(page)).toContain('a:started')
    await expect(popup.locator('[data-officeview-presenter-view]')).toBeVisible()
    return popup
  }

  test('popup-then-full-screen in one activation; a refused full screen shows Click for full screen and retries on the next click', async ({ page }) => {
    await open(page, `${PAGE}?presenter=1&deck=nav`)
    const popup = await startWithPopup(page)
    expect(await presenter(page)).toBe('window')
    expect(await popup.evaluate(() => location.href)).toBe('about:blank')
    // Engine-dependent (spec 4.1): the popup consumed or kept the activation. Either outcome is legal; a refusal must offer the retry.
    await expect.poll(async () => (await log(page)).some((l) => l === 'a:diag:fullscreen-refused') || (await fullscreen(page)) === 'a').toBe(true)
    const refused = (await log(page)).includes('a:diag:fullscreen-refused')
    test.info().annotations.push({ type: 'engine', description: refused ? 'Chromium: window.open consumed the activation; full screen refused, retry offered' : 'Chromium: popup and full screen both succeeded in one activation' })
    if (refused) {
      expect(await page.evaluate(() => window.__slideshowStart!.sessions.a!.surface)).toBe('in-window')
      const retry = page.locator('[data-officeview-show-notice] button', { hasText: 'Click for full screen' })
      await expect(retry).toBeVisible()
      await retry.click()
      await expect.poll(() => fullscreen(page)).toBe('a')
      await expect.poll(() => page.evaluate(() => window.__slideshowStart!.sessions.a!.surface)).toBe('fullscreen')
    }
    // The popup survives the full-screen change; exit closes it.
    expect(popup.isClosed()).toBe(false)
    await page.evaluate(() => window.__slideshowStart!.sessions.a!.exit())
    await expect.poll(() => popup.isClosed()).toBe(true)
    await expect.poll(() => fullscreen(page)).toBeNull()
  })

  test('iframe refusing both fullscreen and popups runs in-window single-screen', async ({ page, context }) => {
    await page.goto(`${PAGE}?frame=sandbox&presenter=1`)
    const frameEl = page.locator('iframe[title="viewer frame"]')
    await expect(frameEl).toBeAttached()
    expect(await frameEl.getAttribute('sandbox')).toBe('allow-scripts allow-same-origin')
    expect(await frameEl.getAttribute('allow')).toBeNull()
    const inner = (await (await frameEl.elementHandle())!.contentFrame())!
    await inner.waitForFunction(() => window.__slideshowStart?.ready === true)
    expect(await inner.evaluate(() => location.origin)).not.toBe(await page.evaluate(() => location.origin))
    const pages = context.pages().length
    await inner.click('#present-a')
    // The promise resolved (never rejected), with two separate diagnostics.
    await expect.poll(() => log(inner)).toEqual(expect.arrayContaining(['a:started', 'a:diag:fullscreen-refused', 'a:diag:popup-blocked']))
    expect((await log(inner)).some((l) => l.startsWith('a:rejected'))).toBe(false)
    expect(await inner.evaluate(() => {
      const s = window.__slideshowStart!.sessions.a!
      return { surface: s.surface, presenter: s.presenter }
    })).toEqual({ surface: 'in-window', presenter: 'swap' })
    expect(context.pages().length).toBe(pages)
    await inner.locator('[data-viewer=a] [data-officeview-show]').press('Escape')
    await expect.poll(() => log(inner)).toContain('a:ended')
  })

  test('CSP-hardened host: presenter controls render and work; the styles diagnostic appears when blocked', async ({ page }) => {
    // Allowed: officeview's stylesheet is a same-origin link, so the popup's copy loads under the inherited policy.
    await open(page, `${CSP_PAGE}?presenter=1&deck=nav`)
    expect(await page.evaluate(() => document.querySelector('meta[http-equiv="Content-Security-Policy"]')?.getAttribute('content'))).toContain("style-src 'self'")
    let popup = await startWithPopup(page)
    await expect.poll(() => popup.evaluate(() => [...document.styleSheets].some((s) => (s.href ?? '').endsWith('/src/styles/officeview.css') && s.cssRules.length > 0))).toBe(true)
    await popup.click('button[aria-label="Next slide"]')
    await expect.poll(() => log(page)).toContain('a:change:1:0')
    await expect(popup.locator('[data-officeview-presenter-status]')).toHaveText('Slide 2 of 5')
    expect(await log(page)).not.toContain('a:diag:presenter-styles-blocked')
    // Esc in the popup exits the show, which closes the popup under the key press.
    await popup.keyboard.press('Escape').catch(() => {})
    await expect.poll(() => log(page)).toContain('a:ended')
    await expect.poll(() => popup.isClosed()).toBe(true)

    // Blocked: stylesheets need the host's nonce, which officeview never replays into the popup.
    await open(page, `${CSP_PAGE}?presenter=1&deck=nav&styles=nonce`)
    expect(await page.evaluate(() => document.querySelector('meta[http-equiv="Content-Security-Policy"]')?.getAttribute('content'))).toContain("style-src 'nonce-ovcsp'")
    // The host page itself is styled (its own link carries the nonce).
    expect(await page.evaluate(() => [...document.styleSheets].some((s) => (s.href ?? '').endsWith('/src/styles/officeview.css')))).toBe(true)
    popup = await startWithPopup(page)
    await expect.poll(() => log(page)).toContain('a:diag:presenter-styles-blocked')
    expect((await log(page)).filter((l) => l === 'a:diag:presenter-styles-blocked')).toHaveLength(1)
    // Unstyled but usable.
    await popup.click('button[aria-label="Next slide"]')
    await expect.poll(() => log(page)).toContain('a:change:1:0')
    await popup.click('button[aria-label="End show"]').catch(() => {})
    await expect.poll(() => log(page)).toContain('a:ended')
    await expect.poll(() => popup.isClosed()).toBe(true)
  })

  test('both views stay in sync on next and previous', async ({ page }) => {
    await open(page, `${PAGE}?presenter=1&deck=nav`)
    const popup = await startWithPopup(page)
    const status = popup.locator('[data-officeview-presenter-status]')
    await expect(status).toHaveText('Slide 1 of 5')
    // Audience input moves the presenter view.
    await page.locator('[data-viewer=a] [data-officeview-show]').press('ArrowRight')
    await expect.poll(() => log(page)).toContain('a:change:1:0')
    await expect(status).toHaveText('Slide 2 of 5')
    // Presenter controls and keys move the audience (the hidden third slide is skipped).
    await popup.click('button[aria-label="Next slide"]')
    await expect.poll(() => page.evaluate(() => window.__slideshowStart!.sessions.a!.slideIndex)).toBe(3)
    await expect(status).toHaveText('Slide 4 of 5')
    await popup.locator('[data-officeview-presenter-view]').press('ArrowLeft')
    await expect.poll(() => page.evaluate(() => window.__slideshowStart!.sessions.a!.slideIndex)).toBe(1)
    await popup.click('button[aria-label="Previous slide"]')
    await expect.poll(() => page.evaluate(() => window.__slideshowStart!.sessions.a!.slideIndex)).toBe(0)
    await expect(status).toHaveText('Slide 1 of 5')
    // Closing the presenter window keeps the show, single-screen.
    await popup.close()
    await expect.poll(() => presenter(page)).toBe('swap')
    expect(await log(page)).not.toContain('a:ended')
    await page.locator('[data-viewer=a] [data-officeview-show]').press('ArrowRight')
    await expect.poll(() => page.evaluate(() => window.__slideshowStart!.sessions.a!.slideIndex)).toBe(1)
    await page.evaluate(() => window.__slideshowStart!.sessions.a!.exit())
    await expect.poll(() => log(page)).toContain('a:ended')
  })
})

/**
 * S2-3 screen placement. Window Management is feature-detection mocked in an
 * init script: `navigator.permissions.query({name:'window-management'})`,
 * `getScreenDetails` (an own property of the global, Window being [Global],
 * so "missing" deletes it and really is missing),
 * and recorders around `window.open` and `requestFullscreen`. The recorder
 * strips `screen` before delegating: a plain mock object is not a
 * ScreenDetailed, and Chromium would refuse the request on it.
 */
interface WmLog { queries: number; calls: number; opens: string[]; popups: Array<Window | null>; fullscreen: Array<string | null>; moves: number[][]; resizes: number[][] }
type ScreensModule = typeof import('../../src/slideshow/presenter/screens')
type Target = Parameters<ScreensModule['useBothScreens']>[1]
declare global { interface Window { __wm?: WmLog; __screens?: ScreensModule; __target?: (refuse: boolean) => Target } }

type Mock = { permission: 'granted' | 'prompt' | 'denied'; screens?: number; missing?: boolean }

async function mockWindowManagement(page: Page, mock: Mock): Promise<void> {
  await page.addInitScript((m: Mock) => {
    const wm: WmLog = { queries: 0, calls: 0, opens: [], popups: [], fullscreen: [], moves: [], resizes: [] }
    window.__wm = wm
    let state: string = m.permission
    const all = [
      { isPrimary: true, isInternal: true, label: 'primary', left: 0, top: 0, width: 800, height: 600, availLeft: 0, availTop: 0, availWidth: 800, availHeight: 600 },
      { isPrimary: false, isInternal: false, label: 'external', left: 800, top: 0, width: 1024, height: 768, availLeft: 800, availTop: 0, availWidth: 1024, availHeight: 768 },
    ]
    const details = { screens: all.slice(0, m.screens ?? 2), currentScreen: all[0] }
    const global = window as unknown as Record<string, unknown>
    if (m.missing) delete global.getScreenDetails
    else {
      global.getScreenDetails = function () {
        wm.calls++
        if (state == 'denied') return Promise.reject(Object.assign(new Error('denied'), { name: 'NotAllowedError' }))
        state = 'granted' // a prompt answered "allow"
        return Promise.resolve(details)
      }
    }
    const query = navigator.permissions.query.bind(navigator.permissions)
    navigator.permissions.query = ((d: { name: string }) => {
      if (d?.name != 'window-management') return query(d as PermissionDescriptor)
      wm.queries++
      return m.missing ? Promise.reject(new TypeError('unknown permission')) : Promise.resolve({ state } as PermissionStatus)
    }) as typeof navigator.permissions.query
    const open = window.open
    window.open = function (this: Window, ...a: Parameters<typeof window.open>) {
      const w = open.apply(this, a)
      wm.opens.push(String(a[2] ?? ''))
      wm.popups.push(w)
      if (w) {
        const mv = w.moveTo.bind(w), rs = w.resizeTo.bind(w)
        w.moveTo = (x: number, y: number) => { wm.moves.push([x, y]); mv(x, y) }
        w.resizeTo = (x: number, y: number) => { wm.resizes.push([x, y]); rs(x, y) }
      }
      return w
    }
    const rf = Element.prototype.requestFullscreen
    Element.prototype.requestFullscreen = function (this: Element, o?: FullscreenOptions & { screen?: { label?: string } }) {
      wm.fullscreen.push(o?.screen?.label ?? null)
      const { screen: _s, ...rest } = o ?? {}
      return rf.call(this, rest)
    }
  }, mock)
}

test.describe('S2-3 screen placement', () => {
  test.beforeEach(({ page }) => {
    errors.length = 0
    page.on('pageerror', (e) => errors.push(e.message))
  })
  test.afterEach(() => {
    expect(errors).toEqual([])
  })

  const SCREENS = '/src/slideshow/presenter/screens.ts'
  const wm = (page: Page) => page.evaluate(() => ({ ...window.__wm!, popups: [] }))
  const log = (page: Page) => page.evaluate(() => [...window.__slideshowStart!.log])
  async function open(page: Page, query: string): Promise<void> {
    await page.goto(`${PAGE}?${query}`)
    await page.waitForFunction(() => window.__slideshowStart?.ready === true)
    await expect(page.locator('[data-officeview-show]').first()).toBeAttached()
    // The real module instance (same URL as the app's), with its extension so Vite serves it.
    await page.evaluate(async (u) => {
      window.__screens = await import(/* @vite-ignore */ u)
      // A PlacementTarget on the audience host and the recorded popup; `refuse` makes the request refuse.
      window.__target = (refuse) => {
        const host = document.querySelector('[data-viewer=a] [data-officeview-show]')!
        return { popup: window.__wm!.popups.at(-1) ?? null, retryFullscreen: (o) => refuse ? Promise.resolve('refused') : host.requestFullscreen(o).then(() => 'entered' as const, () => 'refused' as const) }
      }
    }, SCREENS)
  }
  /** The preload hint: hover the toolbar's Present button. */
  async function hint(page: Page): Promise<void> {
    await page.locator('[data-viewer=a] [data-testid=officeview-present]').hover()
    await expect.poll(async () => (await wm(page)).queries).toBeGreaterThan(0)
  }
  async function exit(page: Page): Promise<void> {
    await page.evaluate(() => window.__slideshowStart!.sessions.a!.exit())
    await expect.poll(async () => (await log(page)).filter((l) => l == 'a:ended').length).toBeGreaterThan(0)
  }
  test('granted with two screens: popup on primary, audience requestFullscreen({screen}) on the external screen', async ({ page }) => {
    await mockWindowManagement(page, { permission: 'granted' })
    await open(page, 'toolbar=1&deck=nav')
    await hint(page)
    await expect.poll(async () => (await wm(page)).calls).toBe(1)
    // A plain Present ('auto'): the prepared plan opens the popup on the primary
    // screen and full screen goes to the external screen, in one synchronous prologue.
    const popupEvent = page.waitForEvent('popup')
    await page.click('#present-a')
    const popup = await popupEvent
    await expect.poll(() => log(page)).toContain('a:started')
    const w = await wm(page)
    expect(w.opens).toEqual(['popup=yes,left=0,top=0,width=800,height=600'])
    expect(w.fullscreen).toEqual(['external'])
    expect(await page.evaluate(() => window.__slideshowStart!.sessions.a!.presenter)).toBe('window')
    await expect(popup.locator('[data-officeview-presenter-view]')).toBeVisible()
    expect((await log(page)).filter((l) => l.includes(':diag:placement-') || l.includes(':diag:popup-blocked'))).toEqual([])
    await exit(page)
    await expect.poll(() => popup.isClosed()).toBe(true)
  })

  test('not yet granted: start proceeds and Use both screens asks on click', async ({ page }) => {
    await mockWindowManagement(page, { permission: 'prompt' })
    await open(page, 'toolbar=1&deck=nav')
    await hint(page)
    // The hint never prompts: no getScreenDetails while the permission is "prompt".
    expect((await wm(page)).calls).toBe(0)
    // 'auto' with nothing prepared: single-screen, no popup, full screen on the opener's screen.
    await page.click('#present-a')
    await expect.poll(() => log(page)).toContain('a:started')
    expect(await page.evaluate(() => window.__slideshowStart!.sessions.a!.presenter)).toBe('none')
    expect((await wm(page)).opens).toEqual([])
    expect((await wm(page)).fullscreen).toEqual([null])
    await exit(page)
    // Presenter View on request (Alt+F5): a default-size popup.
    const popupEvent = page.waitForEvent('popup')
    expect(await page.evaluate(() => window.__slideshowStart!.present('a', { presenterView: true }))).toBe('started')
    await popupEvent
    expect((await wm(page)).opens).toEqual(['popup'])
    // "Use both screens": getScreenDetails is called synchronously in the click, then both windows are placed.
    const r = await page.evaluate(async () => {
      const target = window.__target!(false)
      const pending = window.__screens!.useBothScreens(window, target)
      const sync = window.__wm!.calls
      return { sync, result: await pending, state: target.state && [target.state.presenter.label, target.state.audience.label] }
    })
    expect(r.sync).toBe(1)
    expect(r.result).toMatchObject({ outcome: 'placed' })
    expect(r.state).toEqual(['primary', 'external'])
    const w = await wm(page)
    expect(w.fullscreen.at(-1)).toBe('external')
    expect(w.moves).toEqual([[0, 0]])
    expect(w.resizes).toEqual([[800, 600]])
    await exit(page)
    // The grant prepared the next start: a plain Present now opens the placed popup.
    const next = page.waitForEvent('popup')
    await page.click('#present-a')
    await next
    expect((await wm(page)).opens.at(-1)).toBe('popup=yes,left=0,top=0,width=800,height=600')
    await exit(page)
  })

  test('denied, missing API and single screen are three distinct diagnostics, none is popup-blocked', async ({ context }) => {
    const seen: string[] = []
    for (const [mock, reason] of [
      [{ permission: 'denied' }, 'permission-denied'],
      [{ permission: 'prompt', missing: true }, 'api-missing'],
      [{ permission: 'granted', screens: 1 }, 'single-screen'],
    ] as Array<[Mock, string]>) {
      const page = await context.newPage()
      page.on('pageerror', (e) => errors.push(e.message))
      await mockWindowManagement(page, mock)
      await open(page, 'toolbar=1&deck=nav')
      await hint(page)
      // Nothing prepared, so a plain Present stays single-screen.
      await page.click('#present-a')
      await expect.poll(() => log(page)).toContain('a:started')
      expect((await wm(page)).opens).toEqual([])
      const r = await page.evaluate(async () => {
        const codes: string[] = []
        const sink = { add: (d: { code: string }) => void codes.push(d.code), list: () => [] }
        const result = await window.__screens!.useBothScreens(window, window.__target!(false), sink)
        return { result, codes }
      })
      expect(r.result).toEqual({ reason })
      expect(r.codes).toEqual([`placement-${reason}`])
      seen.push(...r.codes)
      expect((await log(page)).some((l) => l.endsWith(':diag:popup-blocked'))).toBe(false)
      await exit(page)
      await page.close()
    }
    expect(new Set(seen).size).toBe(3)
    expect(seen).not.toContain('popup-blocked')
  })

  test('swap displays exchanges the screens', async ({ page }) => {
    await mockWindowManagement(page, { permission: 'granted' })
    await open(page, 'toolbar=1&deck=nav')
    await hint(page)
    await expect.poll(async () => (await wm(page)).calls).toBe(1)
    const popupEvent = page.waitForEvent('popup')
    await page.click('#present-a')
    await popupEvent
    await expect.poll(() => log(page)).toContain('a:started')
    const swap = (refuse: boolean) => page.evaluate(async (refuse) => {
      const d = await (window as unknown as { getScreenDetails(): Promise<{ screens: Array<{ availWidth: number; availHeight: number; label?: string }> }> }).getScreenDetails()
      const target = Object.assign(window.__target!(refuse), { state: { presenter: d.screens[0], audience: d.screens[1] } })
      const codes: string[] = []
      const outcome = await window.__screens!.swapDisplays(target, { add: (x: { code: string }) => void codes.push(x.code), list: () => [] })
      return { outcome, codes, state: [target.state.presenter.label, target.state.audience.label] }
    }, refuse)
    // Refused (no activation in the opener): nothing moves, nothing flips, diagnosed, no page error.
    expect(await swap(true)).toEqual({ outcome: 'refused', codes: ['placement-fullscreen-refused'], state: ['primary', 'external'] })
    expect((await wm(page)).moves).toEqual([])
    // Entered: full screen on the old presenter screen, then the popup moves to the external screen.
    expect(await swap(false)).toEqual({ outcome: 'placed', codes: [], state: ['external', 'primary'] })
    const w = await wm(page)
    expect(w.fullscreen.at(-1)).toBe('primary')
    expect(w.moves).toEqual([[800, 0]])
    expect(w.resizes).toEqual([[1024, 768]])
    await exit(page)
  })

  test('other browsers show the one-time drag hint', async ({ page }) => {
    await mockWindowManagement(page, { permission: 'prompt', missing: true })
    await open(page, 'deck=nav')
    expect(await page.evaluate(() => 'getScreenDetails' in window)).toBe(false)
    await page.evaluate(() => localStorage.removeItem('officeview:drag-hint-dismissed'))
    const startPopup = async () => {
      const popupEvent = page.waitForEvent('popup')
      expect(await page.evaluate(() => window.__slideshowStart!.present('a', { presenterView: true }))).toBe('started')
      const popup = await popupEvent
      await expect(popup.locator('[data-officeview-presenter-view]')).toBeVisible()
      return popup
    }
    let popup = await startPopup()
    expect((await wm(page)).opens).toEqual(['popup'])
    // The opener's module shows the hint in the popup's document (another realm).
    expect(await page.evaluate(() => !!window.__screens!.showDragHint(window.__wm!.popups.at(-1)!.document))).toBe(true)
    const hintEl = popup.locator('[data-officeview-drag-hint]')
    await expect(hintEl).toContainText('Drag this presenter window to your other display')
    await popup.click('[data-officeview-drag-hint] button')
    await expect(hintEl).toHaveCount(0)
    await exit(page)
    await expect.poll(() => popup.isClosed()).toBe(true)
    // One-time: the dismissal outlives the popup.
    popup = await startPopup()
    expect(await page.evaluate(() => window.__screens!.showDragHint(window.__wm!.popups.at(-1)!.document) === undefined)).toBe(true)
    await expect(popup.locator('[data-officeview-drag-hint]')).toHaveCount(0)
    expect(await page.evaluate(() => localStorage.getItem('officeview:drag-hint-dismissed'))).toBe('1')
    await exit(page)
    await page.evaluate(() => localStorage.removeItem('officeview:drag-hint-dismissed'))
  })
})
