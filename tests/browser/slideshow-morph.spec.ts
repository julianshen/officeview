/**
 * Slideshow plan S3-6: the morph transition in Chromium (the `slideshow`
 * Playwright project, SwiftShader WebGL2), driven through
 * src/debug/slideshow-start.tsx with ?deck=morph: two slides whose shapes
 * pair (a moved square, a moved title) and slide 2's `p159:morph byObject`
 * (1.2 s). The show ENTERS slide 2 through the stage's dispatch.
 *
 * Asserts behaviour only: on both backends the running transition is the
 * morph route (the stage's `lastDraws` probe shows the matched square's
 * morph pair; the dispatch's `transition-approximate` diagnostic explains
 * the approximation: matched objects interpolate, content crossfades) and no
 * `morph-fallback` is diagnosed. The composited morph frames are covered by
 * the pixel goldens and the morphFrame unit tests.
 */
import { expect, test, type Page } from '@playwright/test'
import { fileURLToPath } from 'node:url'
import { checkIdentity, gitIdentity } from '../../scripts/harness/identity'
import type { SessionController } from '../../src/slideshow/player/session'

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
  await page.goto(`${PAGE}?deck=morph${query}`)
  await page.waitForFunction(() => window.__slideshowStart?.ready === true)
  await page.click('#present-a')
  await expect.poll(() => page.evaluate(() => [...window.__slideshowStart!.log])).toContain('a:started')
  await expect.poll(() => page.locator('[data-officeview-show-surface]').getAttribute('data-slide')).toBe('0')
}
const session = (page: Page) => page.evaluate(() => {
  const s = window.__slideshowStart!.sessions.a as unknown as SessionController
  return { slide: s.state.slide, backend: s.stage!.backend, level: s.stage!.level, transition: s.stage!.transition, draws: s.stage!.inspect().lastDraws }
})
const diagnostics = (page: Page) => page.evaluate(() => {
  const s = window.__slideshowStart!.sessions.a as unknown as SessionController
  return s.presenterSource.diagnostics().map(d => ({ code: d.code, effect: d.effect, message: d.message }))
})
const text = (page: Page) => page.locator('[data-officeview-show-text] li').allTextContents()

for (const path of ['webgl2', 'canvas2d'] as const) {
  test(`the morph runs on the ${path} path with its diagnostic`, async ({ page }) => {
    await present(page, path === 'canvas2d' ? '&force=canvas2d' : '')
    await expect.poll(async () => (await session(page)).level).toBe('full')
    expect((await session(page)).backend).toBe(path)
    if (path === 'canvas2d') expect(await diagnostics(page)).toEqual(expect.arrayContaining([expect.objectContaining({ code: 'backend-fallback' })]))
    expect(await text(page)).toEqual(['Morph me'])
    // Next enters slide 2 through its morph transition (the 1.2 s window; the running record is stable across it).
    await expect(async () => {
      await page.keyboard.press('Space')
      expect((await session(page)).slide).toBe(1)
      const now = await session(page)
      expect(now.transition).toEqual({ effect: 'morph', duration: 1200 })
      // The morph route draws the matched square's pair (not the fade look-alike).
      expect(now.draws).toContain('morph:pair:sp:2')
    }).toPass()
    // The diagnostic says what played: the approximation (matched objects interpolate, content crossfades), on either backend; no fallback.
    const list = (await diagnostics(page)).filter(d => d.code === 'transition-approximate')
    expect(list).toEqual([
      expect.objectContaining({ code: 'transition-approximate', effect: 'morph', message: expect.stringContaining('p159:morph') }),
    ])
    expect(list[0]!.message.includes('approximation')).toBe(true)
    await page.keyboard.press('Escape')
    await expect.poll(() => page.evaluate(() => [...window.__slideshowStart!.log])).toContain('a:ended')
  })
}
