/**
 * Slideshow plan S1-19: the session controller and audience navigation in
 * Chromium (the `slideshow` Playwright project), driven through
 * src/debug/slideshow-start.tsx with ?deck=nav (five slides, the third
 * hidden) or ?deck=loop (three slides, the second hidden, looping). Decks
 * are "Office-shaped synthetic".
 *
 * Asserts behaviour only (the surface's slide, the log, focus, full screen),
 * never pixels or timings; only observable state is polled.
 */
import { expect, test, type Page } from '@playwright/test'
import { fileURLToPath } from 'node:url'
import { checkIdentity, gitIdentity } from '../../scripts/harness/identity'
import type { SessionController } from '../../src/slideshow/player/session'

export type { SlideshowStart } from '../../src/debug/slideshow-start'

const ROOT = fileURLToPath(new URL('../..', import.meta.url))
const BASE = `http://127.0.0.1:${process.env.IV_HARNESS_PORT ?? 4173}`
const PAGE = '/src/debug/slideshow-start.html'

const errors: string[] = []

test.beforeAll(async () => {
  const text = await (await fetch(`${BASE}/__harness-id`)).text()
  let body: unknown = text
  try { body = JSON.parse(text) } catch { /* checkIdentity rejects raw text */ }
  expect(checkIdentity(body, { commit: gitIdentity(ROOT).commit, root: ROOT }), 'harness server identity').toBeNull()
})
test.beforeEach(({ page }) => {
  errors.length = 0
  page.on('pageerror', (e) => errors.push(e.message))
})
test.afterEach(() => {
  expect(errors).toEqual([])
})

async function open(page: Page, deck: 'nav' | 'loop'): Promise<void> {
  await page.goto(`${PAGE}?deck=${deck}`)
  await page.waitForFunction(() => window.__slideshowStart?.ready === true)
  await expect(page.locator('[data-officeview-show]').first()).toBeAttached()
}
const log = (page: Page) => page.evaluate(() => [...window.__slideshowStart!.log])
const fullscreen = (page: Page) => page.evaluate(() => window.__slideshowStart!.fullscreenElement())
const shown = (page: Page) => page.locator('[data-officeview-show-surface]').getAttribute('data-slide')
const state = (page: Page) => page.evaluate(() => (window.__slideshowStart!.sessions.a as unknown as SessionController).state)

async function present(page: Page): Promise<void> {
  await page.click('#present-a')
  await expect.poll(() => log(page)).toContain('a:started')
  await expect.poll(() => shown(page)).not.toBeNull()
}
async function ended(page: Page): Promise<void> {
  await expect.poll(() => log(page)).toEqual(expect.arrayContaining(['a:onEnd', 'a:ended']))
  await expect.poll(() => fullscreen(page)).toBeNull()
  expect(await page.locator('[data-officeview-show]').evaluate((el) => (el as HTMLElement).style.visibility)).toBe('hidden')
}

test('keyboard navigation and key precedence', async ({ page }) => {
  await open(page, 'nav')
  await present(page)
  expect(await shown(page)).toBe('0')
  await page.keyboard.press('ArrowRight')
  await expect.poll(() => shown(page)).toBe('1')
  // the hidden third slide is skipped
  await page.keyboard.press('Space')
  await expect.poll(() => shown(page)).toBe('3')
  await page.keyboard.press('ArrowLeft')
  await expect.poll(() => shown(page)).toBe('1')
  // number entry shows a hidden slide
  await page.keyboard.press('3')
  await page.keyboard.press('Enter')
  await expect.poll(() => shown(page)).toBe('2')
  await page.keyboard.press('End')
  await expect.poll(() => shown(page)).toBe('4')
  await page.keyboard.press('Home')
  await expect.poll(() => shown(page)).toBe('0')
  // black screen: any key resumes without navigating
  await page.keyboard.press('b')
  await expect(page.locator('[data-officeview-show-blank="black"]')).toBeVisible()
  await page.keyboard.press('x')
  await expect(page.locator('[data-officeview-show-blank]')).toHaveCount(0)
  expect(await shown(page)).toBe('0')
  await page.keyboard.press('w')
  await expect(page.locator('[data-officeview-show-blank="white"]')).toBeVisible()
  await page.keyboard.press('ArrowRight')
  await expect(page.locator('[data-officeview-show-blank]')).toHaveCount(0)
  expect(await shown(page)).toBe('0')
  // pause
  await page.keyboard.press('s')
  await expect.poll(async () => (await state(page)).paused).toBe(true)
  await page.keyboard.press('s')
  await expect.poll(async () => (await state(page)).paused).toBe(false)
  expect(await log(page)).toEqual(expect.arrayContaining(['a:change:0:0', 'a:change:1:0', 'a:change:3:0', 'a:change:2:0', 'a:change:4:0']))
  // Esc beats the number buffer and the black screen
  const changes = (await log(page)).filter((l) => l.startsWith('a:change:')).length
  await page.keyboard.press('b')
  await page.keyboard.press('4')
  await expect.poll(async () => (await state(page)).buffer).toBe('4')
  await page.keyboard.press('Escape')
  await ended(page)
  expect((await log(page)).filter((l) => l.startsWith('a:change:')).length).toBe(changes)
})

test('end-of-show click exits', async ({ page }) => {
  await open(page, 'nav')
  await present(page)
  await page.keyboard.press('End')
  await expect.poll(() => shown(page)).toBe('4')
  await page.keyboard.press('ArrowRight')
  await expect.poll(() => shown(page)).toBe('end')
  await expect(page.locator('[data-officeview-show-end]')).toHaveText('End of slide show, click to exit.')
  await expect(page.locator('[data-officeview-show-end]')).toBeVisible()
  await page.locator('[data-officeview-show-end]').click()
  await ended(page)
})

test('hidden slides and loop', async ({ page }) => {
  await open(page, 'loop')
  await present(page)
  expect((await state(page)).order).toEqual([0, 2])
  await page.keyboard.press('ArrowRight')
  await expect.poll(() => shown(page)).toBe('2')
  // past the last visible slide the show starts again, with no end-of-show slide
  await page.keyboard.press('ArrowRight')
  await expect.poll(() => shown(page)).toBe('0')
  await expect(page.locator('[data-officeview-show-end]')).toBeHidden()
  await page.keyboard.press('ArrowLeft')
  await expect.poll(() => shown(page)).toBe('0')
  await page.evaluate(() => window.__slideshowStart!.sessions.a!.exit())
  await ended(page)
})

test('focus restore after an API-started show', async ({ page }) => {
  await open(page, 'nav')
  // started from the button: focus goes back to it
  await page.focus('#present-a')
  await present(page)
  await page.keyboard.press('ArrowRight')
  await expect.poll(() => shown(page)).toBe('1')
  await page.evaluate(() => window.__slideshowStart!.sessions.a!.exit())
  await ended(page)
  // focus moves back once the browser has left full screen (until then the page outside the host is inert)
  await expect.poll(() => page.evaluate(() => document.activeElement?.id)).toBe('present-a')

  // started through the API with nothing focused: focus goes to the viewer root
  await page.evaluate(() => (document.activeElement as HTMLElement | null)?.blur())
  expect(await page.evaluate(() => document.activeElement === document.body)).toBe(true)
  const before = await page.evaluate(() => window.__slideshowStart!.log.length)
  expect(await page.evaluate(() => window.__slideshowStart!.present('a'))).toBe('started')
  await expect.poll(() => page.evaluate((n) => window.__slideshowStart!.log.slice(n), before)).toContain('a:started')
  await page.keyboard.press('End')
  await expect.poll(() => shown(page)).toBe('4')
  await page.evaluate(() => window.__slideshowStart!.sessions.a!.exit())
  await expect.poll(() => page.evaluate((n) => window.__slideshowStart!.log.slice(n), before)).toContain('a:ended')
  await expect.poll(() => page.evaluate(() => document.activeElement?.hasAttribute('data-officeview-root'))).toBe(true)
  // the viewer scrolled to the last slide shown: its page is in view
  const inView = await page.evaluate(() => {
    const region = document.querySelector('[data-officeview-root]')!
    const pages = region.querySelectorAll('canvas')
    const r = region.getBoundingClientRect(), p = pages[4].getBoundingClientRect()
    return p.top < r.bottom && p.bottom > r.top
  })
  expect(inView).toBe(true)
})

test('leaving full screen through the browser exits the show', async ({ page }) => {
  await open(page, 'nav')
  await present(page)
  await expect.poll(() => fullscreen(page)).toBe('a')
  await page.keyboard.press('ArrowRight')
  await expect.poll(() => shown(page)).toBe('1')
  await page.evaluate(() => document.exitFullscreen())
  await ended(page)
  expect(await page.evaluate(() => window.__slideshowStart!.sessions.a!.slideIndex)).toBe(1)
})
