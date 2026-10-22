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
