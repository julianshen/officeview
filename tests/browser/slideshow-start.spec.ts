/**
 * Slideshow plan S1-17: the eager start sequence in Chromium (the `slideshow`
 * Playwright project), driven through src/debug/slideshow-start.tsx.
 *
 * Asserts behaviour only (full-screen element, outcomes, event order), never
 * pixels or timings. The only thing awaited is the browser's own full-screen
 * and chunk-load events, polled on observable state.
 */
import { expect, test, type Frame, type Page, type Route } from '@playwright/test'
import { fileURLToPath } from 'node:url'
import { checkIdentity, gitIdentity } from '../../scripts/harness/identity'

const ROOT = fileURLToPath(new URL('../..', import.meta.url))
const BASE = `http://127.0.0.1:${process.env.IV_HARNESS_PORT ?? 4173}`
const PAGE = '/src/debug/slideshow-start.html'
const CHUNK = /\/src\/slideshow\/index\.ts(\?.*)?$/

export type { SlideshowStart } from '../../src/debug/slideshow-start'

const errors: string[] = []

test.beforeAll(async () => {
  // Never test a stale or foreign server: it must be THIS harness, commit and checkout.
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

async function open(page: Page, query = ''): Promise<void> {
  await page.goto(PAGE + query)
  await page.waitForFunction(() => window.__slideshowStart?.ready === true)
  // the viewer has mounted its host
  await expect(page.locator('[data-officeview-show]').first()).toBeAttached()
}
const log = (page: Page | Frame) => page.evaluate(() => [...window.__slideshowStart!.log])
const fullscreen = (page: Page | Frame) => page.evaluate(() => window.__slideshowStart!.fullscreenElement())
const hostVisibility = (page: Page, viewer = 'a') => page.locator(`[data-viewer=${viewer}] [data-officeview-show]`).evaluate((el) => (el as HTMLElement).style.visibility)

/** Holds the slideshow chunk until released, or fails it. */
async function holdChunk(page: Page, outcome: 'delay' | 'fail'): Promise<{ release(): void; requested: Promise<void> }> {
  let release!: () => void
  const gate = new Promise<void>((r) => { release = r })
  let requested!: () => void
  const seen = new Promise<void>((r) => { requested = r })
  await page.route(CHUNK, async (route: Route) => {
    requested()
    if (outcome === 'fail') return route.abort('failed')
    await gate
    return route.continue()
  })
  return { release, requested: seen }
}

test('present from the API inside a click enters full screen', async ({ page }) => {
  await open(page)
  await page.click('#present-a')
  await expect.poll(() => fullscreen(page)).toBe('a')
  await expect.poll(() => log(page)).toEqual(expect.arrayContaining(['a:started', 'a:onStart']))
  expect(await page.evaluate(() => window.__slideshowStart!.sessions.a!.surface)).toBe('fullscreen')
  expect(await page.evaluate(() => document.activeElement?.hasAttribute('data-officeview-show'))).toBe(true)
  await page.evaluate(() => window.__slideshowStart!.sessions.a!.exit())
  await expect.poll(() => fullscreen(page)).toBeNull()
  await expect.poll(() => log(page)).toEqual(expect.arrayContaining(['a:onEnd', 'a:ended']))
  expect(await hostVisibility(page)).toBe('hidden')
})

test('present() from page.evaluate without a gesture rejects gesture-required', async ({ page }) => {
  // Playwright's evaluate carries a user gesture to its frame, its ancestors and
  // same-origin descendants, so the call runs in a cross-origin inner frame that
  // never had one, asked through postMessage (which carries no activation).
  await page.goto(`${PAGE}?frame=cross`)
  await page.waitForFunction(() => window.__slideshowFrame?.ready === true)
  const inner = (await (await page.locator('iframe[title="viewer frame"]').elementHandle())!.contentFrame())!
  await page.evaluate(() => document.querySelector('iframe')!.contentWindow!.postMessage({ cmd: 'present' }, '*'))
  await expect.poll(() => page.evaluate(() => window.__slideshowFrame!.results)).toEqual([{ result: 'rejected:gesture-required', active: false }])
  // nothing was opened: no full screen, the host stays hidden
  expect(await inner.evaluate(() => document.fullscreenElement)).toBeNull()
  expect(await inner.locator('[data-officeview-show]').evaluate((el) => (el as HTMLElement).style.visibility)).toBe('hidden')
  expect(await log(inner)).toEqual(['a:rejected:gesture-required'])
})

test('cold start: a delayed chunk still enters full screen first; a failed chunk closes everything', async ({ page }) => {
  await open(page)
  const held = await holdChunk(page, 'delay')
  await page.click('#present-a')
  await held.requested
  // full screen was entered inside the gesture, before the chunk arrived
  await expect.poll(() => fullscreen(page)).toBe('a')
  expect(await log(page)).not.toContain('a:started')
  expect(await page.locator('[data-viewer=a] [data-officeview-show]').evaluate((el) => el.matches(':empty'))).toBe(true)
  held.release()
  await expect.poll(() => log(page)).toContain('a:started')
  expect(await page.locator('[data-viewer=a] [data-officeview-show]').evaluate((el) => el.matches(':empty'))).toBe(false)
  await page.evaluate(() => window.__slideshowStart!.sessions.a!.exit())
  await expect.poll(() => fullscreen(page)).toBeNull()

  // a failed chunk: fresh page, so the module is not cached
  await page.unroute(CHUNK)
  await open(page)
  await holdChunk(page, 'fail')
  await page.click('#present-a')
  await expect.poll(() => log(page)).toContain('a:rejected:load-failed')
  await expect.poll(() => fullscreen(page)).toBeNull()
  expect(await hostVisibility(page)).toBe('hidden')
  await expect(page.getByRole('alert')).toHaveText('Slideshow could not load. Reload to retry.')
  // the viewer stays usable
  await expect(page.getByRole('region', { name: 'Document viewer' })).toBeVisible()
  errors.length = 0 // the aborted module request reports itself as a page error
})

test('iframe without allow=fullscreen runs in-window with a notice', async ({ page }) => {
  await page.goto(`${PAGE}?frame=cross`)
  const frameEl = page.locator('iframe[title="viewer frame"]')
  await expect(frameEl).toBeAttached()
  expect(await frameEl.getAttribute('allow')).toBeNull()
  const frame = (await frameEl.elementHandle())!.contentFrame()
  const inner = (await frame)!
  await inner.waitForFunction(() => window.__slideshowStart?.ready === true)
  expect(await inner.evaluate(() => location.origin)).not.toBe(await page.evaluate(() => location.origin))
  await inner.click('#present-a')
  await expect.poll(() => log(inner)).toEqual(expect.arrayContaining(['a:started', 'a:diag:fullscreen-refused']))
  expect(await inner.evaluate(() => window.__slideshowStart!.sessions.a!.surface)).toBe('in-window')
  expect(await fullscreen(inner)).toBeNull()
  // the host is fixed over the frame's viewport, with a dismissible notice
  const host = inner.locator('[data-viewer=a] [data-officeview-show]')
  expect(await host.evaluate((el) => ({ v: (el as HTMLElement).style.visibility, p: getComputedStyle(el).position }))).toEqual({ v: 'visible', p: 'fixed' })
  await expect(inner.locator('[data-officeview-show-notice]')).toContainText(/full screen/i)
  await inner.click('[data-officeview-show-notice] >> text=Dismiss')
  await expect(inner.locator('[data-officeview-show-notice]')).toHaveCount(0)
  // Esc exits an in-window show
  await inner.locator('[data-viewer=a] [data-officeview-show]').press('Escape')
  await expect.poll(() => log(inner)).toEqual(expect.arrayContaining(['a:onEnd', 'a:ended']))
})

test('two viewers on one page: the second present is rejected; browser Esc exits only the owner', async ({ page }) => {
  await open(page, '?viewers=2')
  // Both in one click: A's requestFullscreen consumes the activation, so
  // Chromium refuses B before it reaches the coordinator.
  await page.click('#present-both')
  await expect.poll(() => fullscreen(page)).toBe('a')
  await expect.poll(() => log(page)).toEqual(expect.arrayContaining(['b:rejected:gesture-required', 'a:started']))
  // B from its own gesture while A's show runs: the coordinator refuses it.
  await page.keyboard.press('b')
  await expect.poll(() => log(page)).toContain('b:rejected:already-presenting')
  expect(await fullscreen(page)).toBe('a')
  expect(await hostVisibility(page, 'b')).toBe('hidden')
  // the browser leaves full screen (as its Esc does): only A's show ends
  await page.evaluate(() => document.exitFullscreen())
  await expect.poll(() => log(page)).toEqual(expect.arrayContaining(['a:onEnd', 'a:ended']))
  const events = await log(page)
  expect(events.filter((e) => e.startsWith('b:'))).toEqual(['b:rejected:gesture-required', 'b:rejected:already-presenting'])
  expect(events.filter((e) => e === 'a:onEnd')).toHaveLength(1)
  expect(await hostVisibility(page, 'a')).toBe('hidden')
  // B can present now that A released the token; Esc ends it
  await page.click('#present-b')
  await expect.poll(() => fullscreen(page)).toBe('b')
  // Full screen comes first in the start sequence; wait for the show itself before Esc, or Esc cancels the
  // start (b:rejected:cancelled, no onEnd) on a slow runner (PR #19 round 3: CI raced exactly that).
  await expect.poll(() => log(page)).toEqual(expect.arrayContaining(['b:onStart', 'b:started']))
  await page.keyboard.press('Escape')
  await expect.poll(() => fullscreen(page)).toBeNull()
  await expect.poll(() => log(page)).toEqual(expect.arrayContaining(['b:onEnd', 'b:ended']))
  expect((await log(page)).filter((e) => e.startsWith('a:'))).toEqual(['a:onStart', 'a:started', 'a:onEnd', 'a:ended'])
})

test('unmount during load and document change during a show', async ({ page }) => {
  await open(page)
  const held = await holdChunk(page, 'delay')
  await page.click('#present-a')
  await held.requested
  await expect.poll(() => fullscreen(page)).toBe('a')
  await page.evaluate(() => window.__slideshowStart!.unmount())
  await expect.poll(() => log(page)).toContain('a:rejected:cancelled')
  await expect.poll(() => fullscreen(page)).toBeNull()
  const arrived = page.waitForResponse(CHUNK)
  held.release()
  await arrived
  // the late chunk never starts a show (one page task after it arrived)
  await page.evaluate(() => new Promise((r) => setTimeout(r, 0)))
  expect(await log(page)).not.toContain('a:onStart')

  await page.unroute(CHUNK)
  await open(page)
  await page.click('#present-a')
  await expect.poll(() => log(page)).toContain('a:started')
  await expect.poll(() => fullscreen(page)).toBe('a')
  await page.evaluate(() => window.__slideshowStart!.changeDocument())
  await expect.poll(() => log(page)).toEqual(expect.arrayContaining(['a:onEnd', 'a:ended']))
  await expect.poll(() => fullscreen(page)).toBeNull()
  expect((await log(page)).filter((e) => e === 'a:onEnd')).toHaveLength(1)
  expect(await hostVisibility(page)).toBe('hidden')
})
