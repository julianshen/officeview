/**
 * Slideshow plan S1-14: the WebGL2 compositor and context-loss recovery in
 * Chromium (the `slideshow` Playwright project: SwiftShader WebGL2), driven
 * through src/debug/slideshow-gpu.ts.
 *
 * Controller ruling: CI must exercise WebGL2 through SwiftShader. A failure
 * to create the WebGL2 compositor FAILS here (never skips); the stage PR is
 * blocked and escalated to the owner. There is no Canvas-2D-only CI path.
 *
 * Clocks are manual: the spec moves the show clock (`tick`) and the
 * monotonic restore-deadline clock (`advance`) itself. The only thing
 * awaited is the browser's own asynchronous context-loss and restore events,
 * polled on observable state, never on elapsed time.
 */
import { expect, test, type Page } from '@playwright/test'
import { fileURLToPath } from 'node:url'
import { checkIdentity, gitIdentity } from '../../scripts/harness/identity'

const ROOT = fileURLToPath(new URL('../..', import.meta.url))
const BASE = `http://127.0.0.1:${process.env.IV_HARNESS_PORT ?? 4173}`
const PAGE = '/src/debug/slideshow-gpu.html'

// The page's API type; its `declare global` gives `window.__slideshowGpu` (type-only: nothing is loaded in Node).
export type { SlideshowGpu } from '../../src/debug/slideshow-gpu'

const errors: string[] = []

test.beforeAll(async () => {
  // Never test a stale or foreign server: it must be THIS harness, commit and checkout.
  const text = await (await fetch(`${BASE}/__harness-id`)).text()
  let body: unknown = text
  try { body = JSON.parse(text) } catch { /* checkIdentity rejects raw text */ }
  expect(checkIdentity(body, { commit: gitIdentity(ROOT).commit, root: ROOT }), 'harness server identity (a failure here is NOT a WebGL2 failure)').toBeNull()
})

test.beforeEach(async ({ page }) => {
  errors.length = 0
  page.on('pageerror', e => errors.push(e.message))
  await page.goto(PAGE)
  await page.waitForFunction(() => window.__slideshowGpu !== undefined)
})

test.afterEach(() => {
  expect(errors, 'uncaught page errors').toEqual([])
})

const probe = (page: Page) => page.evaluate(() => window.__slideshowGpu!.probe())
const tick = (page: Page, ms: number) => page.evaluate(n => window.__slideshowGpu!.tick(n), ms)
const advance = (page: Page, ms: number) => page.evaluate(n => window.__slideshowGpu!.advance(n), ms)

async function startOnWebGL2(page: Page) {
  const info = await page.evaluate(() => window.__slideshowGpu!.start())
  expect(info.backend, 'the WebGL2 compositor (blocking; escalate to the owner if this fails in CI)').toBe('webgl2')
  expect(info.hasLoseContext, 'WEBGL_lose_context').toBe(true)
  return info
}

async function loseAndWait(page: Page) {
  await page.evaluate(() => window.__slideshowGpu!.lose())
  await expect.poll(async () => (await probe(page)).state).toBe('lost')
}

async function restoreAndWait(page: Page) {
  await page.evaluate(() => window.__slideshowGpu!.restore())
  await expect.poll(async () => (await probe(page)).state).toBe('webgl2')
}

test('the compositor runs on WebGL2 in CI', async ({ page }) => {
  const info = await startOnWebGL2(page)
  console.log(`[slideshow-gpu] ${JSON.stringify(info)}`)
  expect(info.state).toBe('webgl2')
  // The launch args select SwiftShader; a hardware GPU here would mean the flags were not applied.
  expect(info.renderer).toMatch(/SwiftShader/i)
  expect(info.maxTextureSize).toBeGreaterThanOrEqual(2048)
  const p = await probe(page)
  expect(p.canvases).toEqual(['webgl2'])
  // The GPU frame really shows the slide: the square where the show clock puts it, the grey band below.
  expect(p.redX).toBe(p.expectedX)
  expect(p.grey.slice(0, 3).every(v => Math.abs(v - 128) <= 2)).toBe(true)
  await tick(page, 500)
  const q = await probe(page)
  expect(q.redX).toBe(q.expectedX)
  expect(q.redX).toBeGreaterThan(p.redX)
})

test('the forced Canvas 2D path shows the same frame', async ({ page }) => {
  const info = await page.evaluate(() => window.__slideshowGpu!.start({ force: 'canvas2d' }))
  expect(info.backend).toBe('canvas2d')
  await tick(page, 500)
  const p = await probe(page)
  expect(p.canvases).toEqual(['canvas2d'])
  expect(p.redX).toBe(p.expectedX)
  expect(await page.evaluate(() => window.__slideshowGpu!.diagnostics())).toContain('backend-fallback:forced')
})

test('forced context loss during a transition keeps a changing animation advancing on the 2-D canvas; the frame after recovery matches the show clock', async ({ page }) => {
  await startOnWebGL2(page)
  await page.evaluate(() => window.__slideshowGpu!.beginTransition())
  await tick(page, 200)
  expect((await probe(page)).backend).toBe('webgl2')

  await loseAndWait(page)
  const lost = await probe(page)
  // The output canvas under the hidden lost WebGL canvas (PR #18 review P2: no second 2-D canvas).
  expect(lost.canvases).toEqual(['canvas2d', 'webgl2'])
  // The clock keeps running; the 2-D canvas follows it frame by frame.
  const seen: number[] = []
  for (let i = 0; i < 3; i++) {
    await tick(page, 400)
    const p = await probe(page)
    expect(p.state).toBe('lost')
    expect(p.redX).toBe(p.expectedX)
    seen.push(p.redX)
  }
  expect(seen[1]).toBeGreaterThan(seen[0])
  expect(seen[2]).toBeGreaterThan(seen[1])

  await restoreAndWait(page)
  await tick(page, 300)
  const after = await probe(page)
  expect(after.backend).toBe('webgl2')
  expect(after.canvases).toEqual(['webgl2'])
  expect(after.showTime).toBe(1700)
  expect(after.redX).toBe(after.expectedX)
})

test('never restored switches at the deadline', async ({ page }) => {
  await startOnWebGL2(page)
  await loseAndWait(page)
  await advance(page, 2999)
  expect((await probe(page)).state).toBe('lost')
  await advance(page, 1)
  await expect.poll(async () => (await probe(page)).backend).toBe('canvas2d')
  await tick(page, 600)
  const p = await probe(page)
  expect(p.state).toBe('canvas2d')
  expect(p.canvases).toEqual(['canvas2d'])
  expect(p.redX).toBe(p.expectedX)
  expect(await page.evaluate(() => window.__slideshowGpu!.diagnostics())).toContain('backend-fallback:restore-deadline')
  // A late restore is ignored.
  await page.evaluate(() => window.__slideshowGpu!.restore())
  await expect.poll(async () => (await probe(page)).events.restored).toBe(1)
  const q = await probe(page)
  expect(q.state).toBe('canvas2d')
  expect(q.canvases).toEqual(['canvas2d'])
})

test('repeated loss', async ({ page }) => {
  await startOnWebGL2(page)
  for (let i = 0; i < 2; i++) {
    await loseAndWait(page)
    await tick(page, 250)
    const p = await probe(page)
    expect(p.redX).toBe(p.expectedX)
    await restoreAndWait(page)
    const q = await probe(page)
    expect(q.canvases).toEqual(['webgl2'])
    expect(q.redX).toBe(q.expectedX)
  }
  // The third loss pins Canvas 2D for the rest of the show.
  await page.evaluate(() => window.__slideshowGpu!.lose())
  await expect.poll(async () => (await probe(page)).backend).toBe('canvas2d')
  const pinned = await probe(page)
  expect(pinned.state).toBe('canvas2d')
  expect(pinned.canvases).toEqual(['canvas2d'])
  expect(await page.evaluate(() => window.__slideshowGpu!.diagnostics())).toContain('backend-fallback:repeated-loss')
  await page.evaluate(() => window.__slideshowGpu!.restore())
  await expect.poll(async () => (await probe(page)).events.restored).toBe(3)
  await tick(page, 250)
  const later = await probe(page)
  expect(later.state).toBe('canvas2d')
  expect(later.redX).toBe(later.expectedX)
})

test('exit during restore', async ({ page }) => {
  await startOnWebGL2(page)
  await loseAndWait(page)
  const exited = await page.evaluate(() => window.__slideshowGpu!.exit())
  expect(exited).toEqual({ state: 'disposed', canvases: 0 })
  await page.evaluate(() => window.__slideshowGpu!.restore())
  await expect.poll(async () => (await probe(page)).events.restored).toBe(1)
  await advance(page, 5000)
  const p = await probe(page)
  expect(p.state).toBe('disposed')
  expect(p.canvases).toEqual([])
})

test('WebGL2 and Canvas 2D frames agree within the row tolerance on a static slide', async ({ page }) => {
  const r = await page.evaluate(() => window.__slideshowGpu!.staticSlide())
  console.log(`[slideshow-gpu] static slide: ${JSON.stringify(r)}`)
  expect(r.opaque).toBe(true)
  expect(r.blue.slice(0, 3)).toEqual([0x33, 0x66, 0xcc])
  expect(r.ok, `meanAbs ${r.meanAbs}, rowMax ${r.rowMax}`).toBe(true)
})

// One shared scene list, rendered on both backends: every mask kind (as a
// layer mask and as a transition mask), transforms, opacity, colour matrices
// and transitions.
test.describe('WebGL2 ↔ Canvas 2D parity', () => {
  test('the comparison catches a flipped frame', async ({ page }) => {
    const r = await page.evaluate(() => window.__slideshowGpu!.flippedControl())
    expect(r.ok, `meanAbs ${r.meanAbs}, rowMax ${r.rowMax}`).toBe(false)
  })
  for (const group of ['static', 'transform', 'mask-layer', 'mask-transition', 'transition'] as const) {
    test(`${group} scenes agree within tolerance`, async ({ page }) => {
      const results = await page.evaluate(g => window.__slideshowGpu!.parity(g), group)
      for (const r of results) console.log(`[slideshow-gpu] ${group} ${r.name}: meanAbs ${r.meanAbs.toFixed(4)} rowMax ${r.rowMax.toFixed(4)} ${r.ok ? 'ok' : 'FAIL'} ${r.diagnostics.join(' | ')}`)
      expect(results.length).toBeGreaterThan(0)
      expect(results.filter(r => !r.ok).map(r => `${r.name} (meanAbs ${r.meanAbs.toFixed(4)}, rowMax ${r.rowMax.toFixed(4)}, ${r.diagnostics.join(' | ')})`)).toEqual([])
    })
  }
  // S1-15: every 2-D catalogue row through transitionDraw, at the row's tolerance.
  test('every 2-D transition row agrees on WebGL2 and Canvas 2D at 0/25/50/75/100 %', async ({ page }) => {
    const results = await page.evaluate(() => window.__slideshowGpu!.parity('transition-2d'))
    // 26 stage-1 rows plus 9 attribute variants, five progress points each.
    expect(results).toHaveLength((26 + 9) * 5)
    const worst = [...results].sort((a, b) => b.rowMax - a.rowMax).slice(0, 5)
    for (const r of worst) console.log(`[slideshow-gpu] transition-2d worst ${r.name}: meanAbs ${r.meanAbs.toFixed(4)} rowMax ${r.rowMax.toFixed(4)}`)
    expect(results.filter(r => !r.ok).map(r => `${r.name} (meanAbs ${r.meanAbs.toFixed(4)}, rowMax ${r.rowMax.toFixed(4)}, ${r.diagnostics.join(' | ')})`)).toEqual([])
  })
  test('the mask groups cover every mask kind', async ({ page }) => {
    const layer = (await page.evaluate(() => window.__slideshowGpu!.parity('mask-layer'))).map(r => r.name)
    const transition = (await page.evaluate(() => window.__slideshowGpu!.parity('mask-transition'))).map(r => r.name)
    expect(layer).toHaveLength(15)
    expect(transition).toHaveLength(15)
  })
})

// PR #20 review 2: a same-outline Morph pair sums alpha-correctly on both backends (the WebGL2 sum target too).
for (const backend of ['webgl2', 'canvas2d'] as const) {
  test(`a translucent Morph pair keeps its alpha on ${backend}`, async ({ page }) => {
    const r = await page.evaluate(b => window.__slideshowGpu!.morphSum(b), backend)
    expect(r.diagnostics).toEqual([])
    expect(r.results).toHaveLength(6)
    for (const x of r.results) {
      expect(Math.abs(x.inside - 128), JSON.stringify(x)).toBeLessThanOrEqual(2)
      // Below the shape nothing is painted: the sum target lands the right way up.
      expect(x.below, JSON.stringify(x)).toBe(0)
    }
  })
}
