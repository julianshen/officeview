/**
 * Slideshow plan S1-20: the engine in Chromium (the `slideshow` Playwright
 * project, SwiftShader WebGL2), driven through src/debug/slideshow-start.tsx
 * with ?deck=play: slide 1 builds "First build" (with "Animated link") on
 * the first click and "Second build" on the second; "Press" is a trigger
 * revealing "Triggered"; "Jump" links to slide 3; "Animated link" jumps to
 * the last slide. Slide 2 pushes in. Decks are "Office-shaped synthetic".
 *
 * Asserts behaviour only (the backend in use, the ladder rung, the
 * offscreen text list, the buttons, the slide shown, the draws of the last
 * frame by plan layer), never GPU pixels or frame times.
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

async function present(page: Page, query: string): Promise<void> {
  await page.goto(`${PAGE}?deck=play${query}`)
  await page.waitForFunction(() => window.__slideshowStart?.ready === true)
  await page.click('#present-a')
  await expect.poll(() => page.evaluate(() => [...window.__slideshowStart!.log])).toContain('a:started')
  await expect.poll(() => page.locator('[data-officeview-show-surface]').getAttribute('data-slide')).toBe('0')
}
const session = (page: Page) => page.evaluate(() => {
  const s = window.__slideshowStart!.sessions.a as unknown as SessionController
  const i = s.stage!.inspect()
  return { slide: s.state.slide, build: s.state.build, backend: s.stage!.backend, level: s.stage!.level, draws: i.lastDraws, transition: s.stage!.transition }
})
const text = (page: Page) => page.locator('[data-officeview-show-text] li').allTextContents()
const buttons = (page: Page) => page.locator('[data-officeview-show-targets] button').allTextContents()
const log = (page: Page) => page.evaluate(() => [...window.__slideshowStart!.log])

for (const path of ['webgl2', 'canvas2d'] as const) {
  test(`animations play in build order on the WebGL2 and forced Canvas 2D paths (${path})`, async ({ page }) => {
    await present(page, path === 'canvas2d' ? '&force=canvas2d' : '')
    await expect.poll(async () => (await session(page)).level).toBe('full')
    expect((await session(page)).backend).toBe(path)
    if (path === 'canvas2d') expect(await log(page)).toContain('a:diag:backend-fallback')
    await expect.poll(() => text(page)).toEqual(['Press', 'Jump'])
    await page.keyboard.press('Space')
    await expect.poll(() => text(page)).toEqual(['First build', 'Press', 'Jump', 'Animated link'])
    await expect.poll(async () => (await session(page)).draws).toEqual(expect.arrayContaining(['target:sp:2', 'target:sp:7']))
    expect((await session(page)).draws).not.toContain('target:sp:3')
    await page.keyboard.press('Space')
    await expect.poll(() => text(page)).toEqual(['First build', 'Second build', 'Press', 'Jump', 'Animated link'])
    await expect.poll(async () => (await session(page)).draws).toContain('target:sp:3')
    expect((await session(page)).build).toBe(2)
    // Next leaves the slide through its push transition (a Next during the 300 ms fade settles it first)
    await expect(async () => {
      await page.keyboard.press('Space')
      expect((await session(page)).slide).toBe(1)
    }).toPass()
    await expect.poll(() => text(page)).toEqual(['Pushed in'])
    await page.keyboard.press('Escape')
    await expect.poll(() => log(page)).toContain('a:ended')
  })
}

test('keyboard activation of triggers and links; an animated link', async ({ page }) => {
  await present(page, '')
  await expect.poll(() => buttons(page)).toEqual(['Trigger: Press', 'Link: Jump'])
  // Tab reaches the buttons (focus moves through them; nothing traps it)
  await page.locator('[data-officeview-show-targets] button', { hasText: 'Trigger: Press' }).focus()
  await page.keyboard.press('Enter')
  await expect.poll(() => text(page)).toContain('Triggered')
  expect((await session(page)).slide).toBe(0)
  expect((await session(page)).build).toBe(0)
  await page.keyboard.press('Tab')
  expect(await page.evaluate(() => document.activeElement?.textContent)).toBe('Link: Jump')
  // the animated link has no button until its entrance has run
  await page.locator('[data-officeview-show]').focus()
  await page.keyboard.press('ArrowRight')
  await expect.poll(() => buttons(page)).toEqual(['Trigger: Press', 'Link: Jump', 'Link: Animated link'])
  await page.locator('[data-officeview-show-targets] button', { hasText: 'Link: Animated link' }).focus()
  await page.keyboard.press('Enter')
  await expect.poll(async () => (await session(page)).slide).toBe(3)
  // a slide-jump link by keyboard from the first slide
  await page.locator('[data-officeview-show]').focus()
  await page.keyboard.press('Home')
  await expect.poll(async () => (await session(page)).slide).toBe(0)
  await page.locator('[data-officeview-show-targets] button', { hasText: 'Link: Jump' }).focus()
  await page.keyboard.press(' ')
  await expect.poll(async () => (await session(page)).slide).toBe(2)
  await page.keyboard.press('Escape')
  await expect.poll(() => log(page)).toContain('a:ended')
})

test('no premature text disclosure', async ({ page }) => {
  // Watch the text list from before the show starts: no build text may ever appear before its click.
  await page.goto(`${PAGE}?deck=play`)
  await page.waitForFunction(() => window.__slideshowStart?.ready === true)
  await page.evaluate(() => {
    const seen: string[] = []
    ;(window as unknown as { __seen: string[] }).__seen = seen
    new MutationObserver(() => {
      for (const li of document.querySelectorAll('[data-officeview-show-text] li')) seen.push(li.textContent ?? '')
    }).observe(document.body, { subtree: true, childList: true, characterData: true })
  })
  await page.click('#present-a')
  await expect.poll(() => text(page)).toEqual(['Press', 'Jump'])
  // let frames run: nothing builds on its own
  for (let i = 0; i < 10; i++) await page.evaluate(() => new Promise(requestAnimationFrame))
  expect(await text(page)).toEqual(['Press', 'Jump'])
  const s = await session(page)
  for (const hidden of ['target:sp:2', 'target:sp:3', 'target:sp:5', 'target:sp:7']) expect(s.draws).not.toContain(hidden)
  const seen = await page.evaluate(() => (window as unknown as { __seen: string[] }).__seen)
  for (const t of ['First build', 'Second build', 'Triggered', 'Animated link']) expect(seen).not.toContain(t)
  await page.keyboard.press('Escape')
})

test('reduced motion', async ({ page }) => {
  await present(page, '&motion=reduced')
  await expect.poll(async () => (await session(page)).level).toBe('full')
  // steps jump to their end states: the fading second build is listed at once
  await page.keyboard.press('Space')
  await page.keyboard.press('Space')
  await expect.poll(() => text(page)).toContain('Second build')
  // the push becomes a crossfade of at most 200 ms
  await page.keyboard.press('Space')
  await expect.poll(async () => (await session(page)).slide).toBe(1)
  expect(await log(page)).toContain('a:change:0:2')
  const t = (await session(page)).transition
  if (t) expect(t.effect === 'fade' && t.duration <= 200).toBe(true)
  await expect.poll(async () => (await session(page)).transition).toBeUndefined()
  await page.keyboard.press('Escape')
})

test('terminal fallback under a tight budget reaches direct-paint cuts without stalling', async ({ page }) => {
  await present(page, '&budget=tight')
  await expect.poll(async () => (await session(page)).level).toBe('terminal')
  await expect.poll(() => log(page)).toContain('a:diag:ladder-terminal')
  // builds still step (as cuts) and every slide change is a cut
  await page.keyboard.press('Space')
  await expect.poll(() => text(page)).toContain('First build')
  await page.keyboard.press('Space')
  await expect.poll(() => text(page)).toContain('Second build')
  for (const want of [1, 2, 3]) {
    // (a Next during the second build's fade settles it first)
    await expect(async () => {
      await page.keyboard.press('ArrowRight')
      expect((await session(page)).slide).toBe(want)
    }).toPass({ timeout: 5000 })
    // each slide settles on the terminal rung (its ensure runs right after the cut) and changes are cuts
    await expect.poll(async () => (await session(page)).level).toBe('terminal')
    expect((await session(page)).transition).toBeUndefined()
  }
  await page.keyboard.press('Escape')
  await expect.poll(() => log(page)).toContain('a:ended')
})
