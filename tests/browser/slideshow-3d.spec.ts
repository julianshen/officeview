/**
 * Slideshow plan S3-1: the 3-D transition framework in Chromium (the
 * `slideshow` Playwright project: SwiftShader WebGL2), driven through
 * `window.__slideshowGpu.three` (src/debug/slideshow-gpu.ts) with a SAMPLE
 * shader registered for `p14:vortex` in that page only.
 *
 * It asserts behaviour on both paths, never GPU frame times: the shader's
 * ends match the flat frames, its middle is a turned card and not the
 * look-alike, Canvas 2D draws the catalogued look-alike with a diagnostic,
 * and a real context loss recovers. A failure to create the WebGL2
 * compositor FAILS here (controller ruling: CI must exercise WebGL2).
 */
import { expect, test } from '@playwright/test'
import { fileURLToPath } from 'node:url'
import { checkIdentity, gitIdentity } from '../../scripts/harness/identity'

const ROOT = fileURLToPath(new URL('../..', import.meta.url))
const BASE = `http://127.0.0.1:${process.env.IV_HARNESS_PORT ?? 4173}`
const PAGE = '/src/debug/slideshow-gpu.html'

export type { SlideshowGpu } from '../../src/debug/slideshow-gpu'

const errors: string[] = []

test.beforeAll(async () => {
  const text = await (await fetch(`${BASE}/__harness-id`)).text()
  let body: unknown = text
  try { body = JSON.parse(text) } catch { /* checkIdentity rejects raw text */ }
  expect(checkIdentity(body, { commit: gitIdentity(ROOT).commit, root: ROOT }), 'harness server identity').toBeNull()
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

test('the sample 3-D shader reproduces the outgoing and incoming frames at its ends', async ({ page }) => {
  const rows = await page.evaluate(() => window.__slideshowGpu!.three.ends())
  console.log(`[slideshow-3d] ends ${JSON.stringify(rows)}`)
  for (const r of rows) {
    expect(r.ok, `${r.which} @${r.progress}: meanAbs ${r.meanAbs}, rowMax ${r.rowMax}`).toBe(true)
    expect(r.diagnostics).toEqual([])
    // Exactly 0 and 1 draw the flat frame; just inside the ends the shader itself draws.
    expect(r.shaderCalls).toBe(r.progress === 0 || r.progress === 1 ? 0 : 1)
  }
})

test('mid-way the shader draws a turned card, not the look-alike', async ({ page }) => {
  const rows = await page.evaluate(() => window.__slideshowGpu!.three.middle())
  console.log(`[slideshow-3d] middle ${JSON.stringify(rows)}`)
  for (const r of rows) {
    expect(r.shaderCalls).toBe(1)
    expect(r.diagnostics).toEqual([])
    // The card's far edge no longer covers that side of the slide (the right before 50 %, the left after): opaque black there.
    expect(r.progress < 0.5 ? r.right : r.left).toEqual([0, 0, 0, 255])
    expect(r.progress < 0.5 ? r.left : r.right).not.toEqual([0, 0, 0, 255])
    // The centre shows the face of the half: the outgoing frame before 50 %, the incoming one after.
    const face = r.progress < 0.5 ? r.fromCentre : r.toCentre
    const other = r.progress < 0.5 ? r.toCentre : r.fromCentre
    const dist = (a: number[], b: number[]) => Math.max(...a.slice(0, 3).map((v, i) => Math.abs(v - b[i])))
    expect(dist(r.centre, face)).toBeLessThan(dist(r.centre, other))
    // Drawn inside the slide only: the letterbox keeps the readback backdrop.
    expect(r.letterbox).toEqual([32, 32, 32, 255])
    expect(r.vsLookalike).toBeGreaterThan(0.05)
  }
})

test('Canvas 2D draws the catalogued look-alike with a diagnostic', async ({ page }) => {
  const r = await page.evaluate(() => window.__slideshowGpu!.three.canvas2d())
  console.log(`[slideshow-3d] canvas2d ${JSON.stringify(r)}`)
  expect(r.route).toBe('lookalike')
  expect(r.glRoute, 'the WebGL2 compositor (blocking; escalate to the owner if this fails in CI)').toBe('shader')
  expect(r.glDiagnostics).toBe(0)
  expect(r.diagnostics).toEqual([expect.objectContaining({ code: 'transition-approximate', slideIndex: 4, effect: 'vortex' })])
  expect(r.diagnostics[0].message).toContain('dissolve')
  for (const x of r.results) {
    expect(x.shaderCalls).toBe(0)
    expect(x.ok, `@${x.progress}: meanAbs ${x.meanAbs}, rowMax ${x.rowMax}`).toBe(true)
  }
})

test('a context loss during a 3-D transition drops its meshes and recovers the same frame', async ({ page }) => {
  const r = await page.evaluate(() => window.__slideshowGpu!.three.contextLoss())
  console.log(`[slideshow-3d] contextLoss ${JSON.stringify(r)}`)
  expect(r.routeBefore).toBe('shader')
  expect(r.heldLost).toBe(0)
  expect(r.routeLost).toBe('lookalike')
  expect(r.rebuilt).toBe(true)
  expect(r.routeAfter).toBe('shader')
  expect(r.same, `meanAbs ${r.meanAbs}, rowMax ${r.rowMax}`).toBe(true)
  expect(r.heldEnd).toBe(0)
})
