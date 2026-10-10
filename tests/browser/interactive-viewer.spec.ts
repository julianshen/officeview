/**
 * Real-browser smoke test for the S0.3 interactive-viewer harness scaffold.
 *
 * Runs against the vite dev server (see playwright.viewer.config.ts webServer).
 * Asserts harness API presence, an input→present marker round-trip and result
 * schema/hash integrity — never viewer SLA values (the viewer does not exist
 * yet; this scaffold proves no SLA).
 */
import { expect, test } from '@playwright/test'
import { validateHarnessResult, verifyHarnessResult } from '../../src/debug/harness/result'

test('harness smoke', async ({ page }) => {
  await page.goto('/src/debug/interactive-viewer.html?fixture=F1&profile=desktop&mode=smoke&cache=cold')
  await page.waitForFunction(() => window.__viewerHarness?.booted === true, null, { timeout: 30_000 })

  const api = await page.evaluate(() => {
    const h = window.__viewerHarness
    return h ? Object.keys(h).sort() : []
  })
  expect(api).toEqual(
    expect.arrayContaining(['booted', 'markInput', 'markPresent', 'recordRoundTrip', 'sampleCount', 'startLongTaskCollection', 'stopLongTaskCollection', 'exportResult']),
  )

  // Input → present marker round-trip through the double-rAF presentation path.
  await page.evaluate(() => window.__viewerHarness!.recordRoundTrip('smoke-1'))
  const presentSamples = await page.evaluate(() => window.__viewerHarness!.sampleCount('input-to-present'))
  expect(presentSamples).toBe(1)

  // Long-task collection is exercised; unsupported engines report unavailable.
  const longtaskSupported = await page.evaluate(() => window.__viewerHarness!.startLongTaskCollection())
  await page.evaluate(() => window.__viewerHarness!.stopLongTaskCollection())

  const result = (await page.evaluate(() => window.__viewerHarness!.exportResult())) as unknown as Record<string, unknown>
  const validation = validateHarnessResult(result)
  const validationErrors = validation.ok ? [] : validation.errors
  expect(validation.ok, validationErrors.join('; ')).toBe(true)
  const verification = await verifyHarnessResult(result)
  expect(verification.ok, verification.errors.join('; ')).toBe(true)

  const env = (result.environment ?? {}) as Record<string, unknown>
  expect(env.fixtureId).toBe('F1')
  expect(env.profile).toBe('desktop')
  expect(env.mode).toBe('smoke')
  expect(env.cache).toMatchObject({ state: 'cold' })
  expect(env.source).toMatchObject({ bytes: expect.any(Number) })

  // Measured samples exist (scaffold render one-shot + the round-trip); the
  // one-shot is a real measurement of the placeholder render, not an SLA claim.
  expect(result.samples).toEqual(
    expect.arrayContaining([
      expect.objectContaining({ name: 'input-to-present', unit: 'ms', kind: 'operation', values: [expect.any(Number)] }),
      expect.objectContaining({ name: 'scaffold-render', unit: 'ms', kind: 'operation', values: [expect.any(Number)] }),
    ]),
  )
  const metrics = (result.metrics ?? {}) as Record<string, { status: string }>
  expect(metrics['input-to-present'].status).toBe('insufficient-samples')
  expect(metrics['viewport-frame'].status).toBe('unavailable')
  expect(metrics['cold-import'].status).toBe('unavailable')
  expect(longtaskSupported).toBe(true)
})

// Corner navigation (A1 → XFD1048576) needs the S5 viewport integration; the
// interactive viewer does not exist yet, so this scenario is pinned as fixme.
test.fixme('corner navigation round-trip to XFD1048576', async () => {
  // Requires S5 viewport integration: Go To XFD1048576, verify recenter and
  // header agreement, then export a result containing the navigation sample.
})
