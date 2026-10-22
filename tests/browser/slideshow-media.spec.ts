/**
 * Slideshow plan S2-7: media playback in Chromium (the `slideshow`
 * Playwright project), driven through src/debug/slideshow-start.tsx with
 * ?deck=media ("Office-shaped synthetic": silent PCM WAV parts, which every
 * Chromium decodes) and ?deck=narration (the LibreOffice corpus deck
 * lo-narration.pptx, MPL-2.0, skipped when the corpus is not fetched).
 *
 * Asserts behaviour only - elements, their play state, the log of media
 * events, the visible-text list and diagnostics - never audio output.
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

async function present(page: Page, deck: string): Promise<void> {
  await page.goto(`${PAGE}?deck=${deck}`)
  await page.waitForFunction(() => window.__slideshowStart?.ready === true)
  await page.click('#present-a')
  await expect.poll(() => log(page)).toContain('a:started')
}
const log = (page: Page) => page.evaluate(() => [...window.__slideshowStart!.log])
const text = (page: Page) => page.locator('[data-officeview-show-text] li').allTextContents()
const media = (page: Page, spid: string) => page.locator(`[data-officeview-show-media-spid="${spid}"]`)
const playState = (page: Page, spid: string) => media(page, spid).evaluateAll((els) => els.map((e) => {
  const m = e as HTMLMediaElement
  return { paused: m.paused, time: m.currentTime, blob: m.src.startsWith('blob:') }
}))
const call = (page: Page, what: 'next' | 'goTo1' | 'goTo2') => page.evaluate((w) => {
  const s = window.__slideshowStart!.sessions.a as unknown as SessionController
  if (w === 'next') s.next()
  else s.goTo(w === 'goTo1' ? 1 : 2)
}, what)

test('media playback of a supported format', async ({ page }) => {
  await present(page, 'media')
  expect(await text(page)).not.toContain('After clip')
  await call(page, 'next')
  // The clip plays from its Blob URL, and its real end releases "After clip" (the successor waited for it).
  await expect.poll(() => log(page)).toContain('media:playing:2')
  expect((await playState(page, '2'))[0]?.blob).toBe(true)
  await expect.poll(() => log(page), { timeout: 10_000 }).toContain('media:ended:2')
  await expect.poll(() => text(page)).toContain('After clip')
  // One element per node, on the audience surface; nothing failed.
  expect(await page.locator('audio, video').count()).toBeGreaterThanOrEqual(1)
  expect((await log(page)).filter((l) => /a:diag:media-(autoplay|error|codec|refused)/.test(l))).toEqual([])
})

test('audio spanning slides', async ({ page }) => {
  await present(page, 'media')
  // "Spanning" (numSld=2) starts as slide 1 begins.
  await expect.poll(() => log(page)).toContain('media:playing:4')
  await call(page, 'goTo1')
  await expect.poll(() => page.locator('[data-officeview-show-surface]').getAttribute('data-slide')).toBe('1')
  const kept = await playState(page, '4')
  expect(kept).toHaveLength(1)
  expect(kept[0].paused).toBe(false)
  expect(kept[0].blob).toBe(true)
  await call(page, 'goTo2')
  await expect.poll(() => page.locator('[data-officeview-show-surface]').getAttribute('data-slide')).toBe('2')
  // Its span ran out: the element and its URL are gone.
  await expect.poll(() => media(page, '4').count()).toBe(0)
})

test('a rejected later play()', async ({ page }) => {
  await present(page, 'media')
  await call(page, 'next')
  await expect.poll(() => log(page)).toContain('media:playing:2')
  // The browser refuses the next play() (an autoplay policy): stand it in for one call.
  await page.evaluate(() => {
    const proto = HTMLMediaElement.prototype as HTMLMediaElement & { __play?: () => Promise<void> }
    proto.__play = proto.play
    proto.play = function () { return Promise.reject(new DOMException('refused by policy', 'NotAllowedError')) }
  })
  await call(page, 'next')
  const button = page.locator('[data-officeview-show-play]')
  await expect(button).toHaveText('Click to play')
  // The show continues: the synthetic end releases "After second" at once, diagnosed.
  await expect.poll(() => text(page)).toContain('After second')
  expect(await log(page)).toContain('a:diag:media-autoplay')
  expect((await playState(page, '5'))[0]?.paused).toBe(true)
  await page.evaluate(() => {
    const proto = HTMLMediaElement.prototype as HTMLMediaElement & { __play?: () => Promise<void> }
    proto.play = proto.__play!
  })
  const slideBefore = await page.locator('[data-officeview-show-surface]').getAttribute('data-slide')
  await button.click()
  // The click plays the media in its own activation, and is not a slide advance.
  await expect.poll(async () => (await playState(page, '5'))[0]?.paused).toBe(false)
  await expect(button).toHaveCount(0)
  expect(await page.locator('[data-officeview-show-surface]').getAttribute('data-slide')).toBe(slideBefore)
  const s = await page.evaluate(() => (window.__slideshowStart!.sessions.a as unknown as SessionController).state.build)
  expect(s).toBe(2)
})

test('narration.pptx plays its m4a where the browser decodes it', async ({ page }) => {
  const file = join(ROOT, 'corpus', 'lo-narration.pptx')
  test.skip(!existsSync(file), 'corpus/lo-narration.pptx is not fetched (bun scripts/fetch-corpus.ts)')
  // The harness server serves only src/debug; the corpus file is handed to the page as is.
  await page.route('**/corpus/lo-narration.pptx', (route) => route.fulfill({ body: readFileSync(file), contentType: 'application/vnd.openxmlformats-officedocument.presentationml.presentation' }))
  await page.goto(`${PAGE}?deck=narration`)
  await page.waitForFunction(() => window.__slideshowStart?.ready === true)
  const decodes = await page.evaluate(() => document.createElement('audio').canPlayType('audio/mp4'))
  await page.click('#present-a')
  await expect.poll(() => log(page)).toContain('a:started')
  test.info().annotations.push({ type: 'audio/mp4', description: decodes || 'not decodable' })
  if (decodes) {
    // The narration (shape 6) plays from its slide's start: the media call in the main sequence's first group.
    await expect.poll(() => log(page), { timeout: 10_000 }).toContain('media:playing:6')
  } else {
    // Not decodable here: the poster stays, nothing is read, and the show says why.
    await expect.poll(() => log(page), { timeout: 10_000 }).toContain('a:diag:media-codec')
    expect(await page.locator('audio, video').count()).toBe(0)
  }
})
