/**
 * Slideshow plan S1-22: corpus conformance in Chromium (the `slideshow`
 * Playwright project, SwiftShader WebGL2), on both the WebGL2 path and the
 * forced Canvas 2D path. CI's slideshow job fetches the corpus first, so
 * this runs there; WebGL2 is required (controller ruling): a context that
 * cannot be created fails the WebGL2 test, never a silent Canvas 2D pass.
 *
 * The decks (LibreOffice core sd/qa/unit/data/pptx, MPL-2.0) are never
 * committed: the spec serves corpus/lo-<deck>.pptx to the harness page
 * (src/debug/slideshow-start.tsx ?deck=corpus) with page.route.
 *
 * Asserts behaviour only (the backend, the slide reached, page errors, the
 * frame-work recorder producing samples), never GPU pixels or frame-time
 * numbers: those belong to slideshow-perf.spec.ts (OV_SLIDESHOW_BENCH=1).
 */
import { expect, test, type Page } from '@playwright/test'
import { existsSync, readFileSync } from 'node:fs'
import { join } from 'node:path'
import { fileURLToPath } from 'node:url'
import { checkIdentity, gitIdentity } from '../../scripts/harness/identity'
import type { SessionController } from '../../src/slideshow/player/session'

export type { SlideshowStart } from '../../src/debug/slideshow-start'

const ROOT = fileURLToPath(new URL('../..', import.meta.url))
const BASE = `http://127.0.0.1:${process.env.IV_HARNESS_PORT ?? 4173}`
const PAGE = '/src/debug/slideshow-start.html'
const CORPUS = join(ROOT, 'corpus')
/** Every LibreOffice deck in docs/validation/slideshow-corpus.md. */
const DECKS = [
  'tdf104788', 'tdf115394', 'tdf142915', 'tdf44223', 'bnc591147', 'connector-shape-animations',
  'tdf104792-smart-art-animation', 'narration', 'airplane-transition', 'tdf156808', 'tdf149588_transparentSolidFill',
  'tdf105739', 'tdf146223', 'tdf123684', 'ooxtheme', 'tdf93868',
]
// Locally the corpus may be absent (bun scripts/fetch-corpus.ts fetches it); in CI it must be there.
const HAVE = existsSync(join(CORPUS, 'lo-tdf104788.pptx'))
test.skip(!HAVE && !process.env.CI, 'corpus/ not fetched (bun scripts/fetch-corpus.ts)')

const errors: string[] = []
test.beforeAll(async () => {
  const text = await (await fetch(`${BASE}/__harness-id`)).text()
  let body: unknown = text
  try { body = JSON.parse(text) } catch { /* checkIdentity rejects raw text */ }
  expect(checkIdentity(body, { commit: gitIdentity(ROOT).commit, root: ROOT }), 'harness server identity').toBeNull()
})
test.beforeEach(async ({ page }) => {
  errors.length = 0
  page.on('pageerror', (e) => errors.push(e.message))
  await page.route('**/__corpus/*.pptx', (route) => {
    const name = new URL(route.request().url()).pathname.split('/').pop()!
    return route.fulfill({ body: readFileSync(join(CORPUS, name)), contentType: 'application/vnd.openxmlformats-officedocument.presentationml.presentation' })
  })
})
test.afterEach(() => {
  expect(errors).toEqual([])
})

async function present(page: Page, deck: string, query = ''): Promise<void> {
  await page.goto(`${PAGE}?deck=corpus&name=${deck}${query}`)
  await page.waitForFunction(() => window.__slideshowStart?.ready === true)
  await page.click('#present-a')
  await expect.poll(() => page.evaluate(() => [...window.__slideshowStart!.log])).toContain('a:started')
}
const session = (page: Page) => page.evaluate(() => {
  const s = window.__slideshowStart!.sessions.a as unknown as SessionController
  return { slide: s.state.slide, build: s.state.build, backend: s.stage?.backend, level: s.stage?.level, exited: s.state.exited }
})

for (const path of ['webgl2', 'canvas2d'] as const) {
  test(`every corpus deck presents and plays to the end on the ${path === 'webgl2' ? 'WebGL2' : 'forced Canvas 2D'} path (${path})`, async ({ page }) => {
    test.setTimeout(300_000)
    for (const deck of DECKS) {
      await present(page, deck, path === 'canvas2d' ? '&force=canvas2d' : '')
      // WebGL2 is required on the WebGL2 path (SwiftShader in CI): no silent fallback.
      await expect.poll(async () => (await session(page)).backend, { message: `${deck} backend` }).toBe(path)
      // Next through every step and slide (a Next during an effect settles it) until the end-of-show slide,
      // pressing again only while it has not arrived (a slow slide load on 2-CPU SwiftShader just takes more tries).
      await expect(async () => {
        if ((await session(page)).slide !== 'end') await page.keyboard.press('ArrowRight')
        expect((await session(page)).slide, deck).toBe('end')
      }).toPass({ timeout: 60_000, intervals: [40] })
      expect((await session(page)).exited).toBe(false)
      await page.keyboard.press('Escape')
      await expect.poll(() => page.evaluate(() => [...window.__slideshowStart!.log])).toContain('a:ended')
      expect(errors, deck).toEqual([])
    }
  })
}

test('the frame benchmark runs on WebGL2 here: tdf104788 slide 6 records frame work for all six steps (no thresholds in CI)', async ({ page }) => {
  await present(page, 'tdf104788', '&bench=1')
  await expect.poll(async () => (await session(page)).backend).toBe('webgl2')
  const renderer = await page.evaluate(() => window.__slideshowStart!.bench!.renderer)
  expect(renderer).not.toBe('no webgl2')
  await page.evaluate(() => (window.__slideshowStart!.sessions.a as unknown as SessionController).goTo(5))
  await expect.poll(async () => (await session(page)).slide).toBe(5)
  await expect.poll(async () => (await session(page)).level).toBe('full')
  for (let step = 1; step <= 6; step++) {
    await page.evaluate(() => window.__slideshowStart!.bench!.reset())
    await page.keyboard.press('ArrowRight')
    await expect.poll(async () => (await session(page)).build).toBe(step)
    await expect.poll(() => page.evaluate(() => window.__slideshowStart!.bench!.work.length)).toBeGreaterThan(5)
    const work = await page.evaluate(() => [...window.__slideshowStart!.bench!.work])
    expect(work.every((ms) => Number.isFinite(ms) && ms >= 0)).toBe(true)
  }
  await page.keyboard.press('Escape')
  await expect.poll(() => page.evaluate(() => [...window.__slideshowStart!.log])).toContain('a:ended')
})
