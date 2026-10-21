/**
 * Slideshow plan S1-2: the CI smoke for the `slideshow` Playwright project
 * (Chromium, SwiftShader WebGL2). Both tests are blocking.
 *
 * Controller ruling: CI must exercise WebGL2 through SwiftShader. If the
 * context cannot be created, this test FAILS (never skips): the stage PR is
 * blocked and escalated to the owner. There is no Canvas-2D-only CI path.
 */
import { expect, test } from '@playwright/test'
import { fileURLToPath } from 'node:url'
import { checkIdentity, gitIdentity } from '../../scripts/harness/identity'

const ROOT = fileURLToPath(new URL('../..', import.meta.url))
const BASE = `http://127.0.0.1:${process.env.IV_HARNESS_PORT ?? 4173}`
const PAGE = '/src/debug/slideshow-smoke.html'

interface Smoke { log: string[]; webgl2(): { ok: false } | { ok: true; maxTextureSize: number; version: string; renderer: string } }
declare global { interface Window { __slideshowSmoke?: Smoke } }

test.beforeAll(async () => {
  // Never test a stale or foreign server: it must be THIS harness, commit and checkout.
  const text = await (await fetch(`${BASE}/__harness-id`)).text()
  let body: unknown = text
  try { body = JSON.parse(text) } catch { /* checkIdentity rejects raw text */ }
  expect(checkIdentity(body, { commit: gitIdentity(ROOT).commit, root: ROOT }), 'harness server identity (a failure here is NOT a WebGL2 failure)').toBeNull()
})

test.beforeEach(async ({ page }) => {
  await page.goto(PAGE)
  await page.waitForFunction(() => window.__slideshowSmoke !== undefined)
})

test('the slideshow project creates a WebGL2 context under SwiftShader', async ({ page }) => {
  const result = await page.evaluate(() => window.__slideshowSmoke!.webgl2())
  console.log(`[slideshow-smoke] webgl2: ${JSON.stringify(result)}`)
  expect(result.ok, 'WebGL2 context under SwiftShader (blocking; escalate to the owner if this fails in CI)').toBe(true)
  if (!result.ok) return
  expect(result.version).toMatch(/^WebGL 2\.0/)
  // The launch args select SwiftShader; a hardware GPU here would mean the flags were not applied.
  expect(result.renderer).toMatch(/SwiftShader/i)
  // The compositor sizes layers against this (spec 6.3); WebGL2 guarantees at least 2048.
  expect(result.maxTextureSize).toBeGreaterThanOrEqual(2048)
})

test('headless Chromium enters and exits element full screen on a click', async ({ page }) => {
  expect(await page.evaluate(() => document.fullscreenEnabled), 'document.fullscreenEnabled').toBe(true)
  await page.click('#enter')
  await expect.poll(() => page.evaluate(() => document.fullscreenElement?.id ?? null)).toBe('stage')
  await page.click('#exit')
  await expect.poll(() => page.evaluate(() => document.fullscreenElement?.id ?? null)).toBeNull()
  expect(await page.evaluate(() => window.__slideshowSmoke!.log)).toEqual(['enter', 'exit'])
})
