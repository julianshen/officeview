/**
 * Real-browser tests for the S0.3 interactive-viewer harness scaffold.
 *
 * Runs against the vite harness server (see playwright.viewer.config.ts).
 * Asserts harness API presence, a deterministic input→present marker pair,
 * warmup accounting, result schema/hash integrity, engine-aware long-task
 * support and the manual-device path end to end — never viewer SLA values (the
 * viewer does not exist yet; this scaffold proves no SLA).
 */
import { expect, test } from '@playwright/test'
import { spawn, type ChildProcess } from 'node:child_process'
import { mkdtempSync, readdirSync, readFileSync } from 'node:fs'
import { tmpdir } from 'node:os'
import { join } from 'node:path'
import { fileURLToPath } from 'node:url'
import { checkIdentity, commitLabel, gitIdentity } from '../../scripts/harness/identity'
import { validateHarnessResult, verifyHarnessResult } from '../../src/debug/harness/result'

const ROOT = fileURLToPath(new URL('../..', import.meta.url))
const BASE = 'http://127.0.0.1:4173'
const PAGE = '/src/debug/interactive-viewer.html'
const identity = gitIdentity(ROOT)
const label = commitLabel(identity)

// The §12.2 Desktop D profile, so the page conforms and records `conforms: true`.
test.use({ viewport: { width: 1440, height: 900 }, deviceScaleFactor: 2 })

test.beforeAll(async () => {
  // Never measure a stale or foreign server: it must be THIS harness, commit and checkout.
  const response = await fetch(`${BASE}/__harness-id`)
  const text = await response.text()
  let body: unknown = text
  try {
    body = JSON.parse(text)
  } catch {
    // keep the raw text; checkIdentity rejects it
  }
  expect(checkIdentity(body, { commit: identity.commit, root: identity.root })).toBeNull()
})

type Harness = NonNullable<Window['__viewerHarness']>

test('harness smoke', async ({ page, browserName }) => {
  await page.goto(`${PAGE}?fixture=F1&profile=desktop&mode=smoke&cache=cold&commit=${label}`)
  await page.waitForFunction(() => window.__viewerHarness?.booted === true, null, { timeout: 30_000 })

  const api = await page.evaluate(() => Object.keys(window.__viewerHarness ?? {}).sort())
  expect(api).toEqual(
    expect.arrayContaining(['booted', 'markInput', 'markPresent', 'recordRoundTrip', 'sampleCount', 'warmupCount', 'startLongTaskCollection', 'stopLongTaskCollection', 'exportResult', 'ledger', 'deterministicClock', 'fixtureLoader', 'setFixture']),
  )

  // The public markInput/markPresent pair under an injected clock that advances
  // 5 ms per read: input read = 5, present read (after the double rAF) = 10.
  await page.evaluate(async () => {
    const harness = window.__viewerHarness as Harness
    let t = 0
    harness.deterministicClock.inject(() => (t += 5))
    harness.markInput('smoke-1')
    await harness.markPresent('smoke-1')
    await harness.recordRoundTrip('smoke-warmup', true) // discarded, but counted
  })
  expect(await page.evaluate(() => window.__viewerHarness!.sampleCount('input-to-present'))).toBe(1)
  expect(await page.evaluate(() => window.__viewerHarness!.warmupCount('input-to-present'))).toBe(1)

  // Long tasks: WebKit has no longtask entry type (unavailable); Chromium has it
  // and a 120 ms busy loop must be observed as a measured event.
  const longtaskSupported = await page.evaluate(() => window.__viewerHarness!.startLongTaskCollection())
  expect(longtaskSupported).toBe(browserName === 'chromium')
  // Run the busy loop from a timer task: work executed directly inside the
  // CDP evaluate call is not scheduled as a main-thread task, so Blink would not
  // report it as a long task.
  await page.evaluate(
    () =>
      new Promise<void>((resolve) =>
        setTimeout(() => {
          const end = performance.now() + 120
          while (performance.now() < end) {
            /* synthetic long task */
          }
          resolve()
        }, 0),
      ),
  )
  if (browserName === 'chromium') {
    // Entries reach observers asynchronously after the long task ends.
    await expect
      .poll(() => page.evaluate(async () => (await window.__viewerHarness!.exportResult()).samples.find((s) => s.name === 'longtask-duration')?.values.length ?? 0), { timeout: 10_000 })
      .toBeGreaterThanOrEqual(1)
  }

  const result = (await page.evaluate(() => window.__viewerHarness!.exportResult())) as unknown as Record<string, any>
  const validation = validateHarnessResult(result)
  expect(validation.ok, validation.ok ? '' : validation.errors.join('; ')).toBe(true)
  const verification = await verifyHarnessResult(result)
  expect(verification.ok, verification.errors.join('; ')).toBe(true)

  const env = result.environment
  expect(env).toMatchObject({
    buildCommit: label,
    requestedFixtureId: 'F1',
    fixtureId: null, // placeholder content never carries a spec fixture id
    evidence: 'scaffold-placeholder',
    profile: 'desktop',
    mode: 'smoke',
    dpr: 2,
    viewport: { width: 1440, height: 900 },
    profileConformance: { conforms: true, deviations: [] },
    cache: { state: 'cold' },
    source: { adapter: 'xlsx-placeholder', rows: 120, cols: 10 },
    backend: 'canvas2d-placeholder',
    operator: { deviceModel: null, osVersion: null, thermalState: null, batteryState: null },
  })
  expect(env.cache.note).not.toMatch(/fresh/i)
  expect(env.source.bytes).toBeGreaterThan(1000)
  expect(env.os.platform).not.toBe('unknown')
  // Page-level labels: an automated WebKit is plain `webkit`, never `safari`.
  expect(env.browser).toMatchObject(
    browserName === 'chromium'
      ? { name: 'chromium', engine: 'blink', automation: 'webdriver' }
      : { name: 'webkit', engine: 'webkit', automation: 'webdriver' },
  )

  const series = (name: string) => result.samples.find((s: { name: string }) => s.name === name)
  expect(series('input-to-present')).toEqual({ name: 'input-to-present', unit: 'ms', kind: 'operation', values: [5], warmupExcluded: 1 })
  // One real measurement of the placeholder render; its value is not asserted.
  expect(series('scaffold-render')).toMatchObject({ unit: 'ms', kind: 'operation', warmupExcluded: 0, values: [expect.any(Number)] })
  expect(result.metrics['input-to-present']).toMatchObject({ status: 'insufficient-samples', samples: 1, p50: 5, max: 5 })
  expect(result.metrics['input-to-present']).not.toHaveProperty('p95')
  for (const name of ['viewport-frame', 'cold-import', 'progressive-first-viewport']) {
    expect(result.metrics[name]).toEqual({ status: 'unavailable', reason: expect.stringContaining(name) })
  }
  if (browserName === 'chromium') {
    const longtask = series('longtask-duration')
    expect(longtask).toMatchObject({ kind: 'event', unit: 'ms' })
    expect(longtask.values.length).toBeGreaterThanOrEqual(1)
    expect(['insufficient-samples', 'ok']).toContain(result.metrics['longtask-duration'].status)
  } else {
    expect(series('longtask-duration')).toBeUndefined()
    expect(result.metrics['longtask-duration']).toEqual({ status: 'unavailable', reason: expect.stringContaining('not supported') })
  }
})

test('zero long tasks in a supporting engine is a measured empty set', async ({ page, browserName }) => {
  test.skip(browserName !== 'chromium', 'longtask is only supported by Blink')
  await page.goto(`${PAGE}?fixture=F1&profile=desktop&mode=smoke&cache=cold&commit=${label}`)
  await page.waitForFunction(() => window.__viewerHarness?.booted === true, null, { timeout: 30_000 })
  await page.evaluate(() => window.__viewerHarness!.stopLongTaskCollection())
  await page.evaluate(() => window.__viewerHarness!.startLongTaskCollection())
  const result = (await page.evaluate(() => window.__viewerHarness!.exportResult())) as unknown as Record<string, any>
  expect(result.metrics['longtask-duration']).toMatchObject({ status: 'none-observed', samples: 0 })
  expect((await verifyHarnessResult(result)).ok).toBe(true)
})

// ---------------------------------------------------------------------------
// Manual-device path end to end: the real CLI serves the real page on the LAN
// address (an insecure http origin: no crypto.subtle), a headless browser
// exports, and the artifact lands on disk — hash-verified.
// ---------------------------------------------------------------------------

interface ManualServer {
  url: string
  token: string
  origin: string
  output: string
  stop(): Promise<void>
}

async function startManualServer(port: number): Promise<ManualServer> {
  const output = mkdtempSync(join(tmpdir(), 'viewer-manual-'))
  const child: ChildProcess = spawn('bun', ['scripts/bench-interactive-viewer.ts', '--browser', 'safari-manual', '--serve', '--port', String(port), '--output', output, '--device-model', 'e2e-device'], {
    cwd: ROOT,
    stdio: ['ignore', 'pipe', 'pipe'],
  })
  let log = ''
  child.stdout?.on('data', (chunk: Buffer) => (log += chunk.toString()))
  child.stderr?.on('data', (chunk: Buffer) => (log += chunk.toString()))
  const url = await new Promise<string>((resolve, reject) => {
    const timer = setTimeout(() => reject(new Error(`manual server did not print a URL:\n${log}`)), 60_000)
    const poll = setInterval(() => {
      const match = /(http:\/\/[^\s]+interactive-viewer\.html\?[^\s]+)/.exec(log)
      if (match) {
        clearTimeout(timer)
        clearInterval(poll)
        resolve(match[1])
      }
    }, 100)
    child.on('exit', () => reject(new Error(`manual server exited early:\n${log}`)))
  })
  const parsed = new URL(url)
  return {
    url,
    token: parsed.searchParams.get('token') ?? '',
    origin: parsed.origin,
    output,
    async stop() {
      const exited = new Promise<void>((resolve) => child.once('exit', () => resolve()))
      child.kill('SIGINT')
      await Promise.race([exited, new Promise((resolve) => setTimeout(resolve, 10_000))])
      child.kill('SIGKILL')
    },
  }
}

test.describe('manual-device collector over HTTP on the LAN address', () => {
  test('exports with no crypto.subtle on an insecure origin and lands a verified artifact', async ({ page, browserName }) => {
    const server = await startManualServer(browserName === 'chromium' ? 4181 : 4182)
    try {
      test.skip(new URL(server.url).hostname === '127.0.0.1', 'no LAN interface: cannot exercise an insecure (non-loopback) origin')
      // A human-operated device has navigator.webdriver === false; the collector
      // refuses automated pages (asserted in the next test), so model the device.
      await page.addInitScript(() => {
        Object.defineProperty(navigator, 'webdriver', { get: () => false })
        try {
          delete (Crypto.prototype as unknown as Record<string, unknown>).subtle
        } catch {
          /* already absent on insecure origins */
        }
      })
      const foreign: string[] = []
      page.on('request', (request) => {
        if (request.url().includes(':9/')) foreign.push(request.url())
      })
      await page.goto(`${server.url}&collect=${encodeURIComponent('http://127.0.0.1:9/evil')}`)
      await page.waitForFunction(() => window.__viewerHarness?.booted === true, null, { timeout: 30_000 })
      expect(await page.evaluate(() => window.isSecureContext)).toBe(false)
      expect(await page.evaluate(() => typeof crypto.subtle)).toBe('undefined')

      await page.getByRole('button', { name: 'Export result' }).click()
      await expect(page.getByText('posted to collector')).toBeVisible({ timeout: 30_000 })
      expect(foreign).toEqual([]) // a ?collect= URL never receives results

      const files = readdirSync(server.output)
      expect(files).toHaveLength(1)
      expect(files[0]).toMatch(/^viewer-manual-F1-desktop-[0-9T-]+Z-[0-9a-f]{8}\.json$/)
      const artifact = JSON.parse(readFileSync(join(server.output, files[0]), 'utf8'))
      expect(artifact.run).toMatchObject({ collection: 'manual-device-post' })
      const [result] = artifact.results
      expect((await verifyHarnessResult(result)).ok).toBe(true)
      expect(result.environment).toMatchObject({
        evidence: 'scaffold-placeholder',
        browser: { automation: 'none' },
        operator: { deviceModel: 'e2e-device' },
        buildCommit: label,
      })
    } finally {
      await server.stop()
    }
  })

  test('rejects automated pages, forged tokens and traversal payloads; nothing escapes --output', async ({ page, request, browserName }) => {
    const server = await startManualServer(browserName === 'chromium' ? 4183 : 4184)
    try {
      await page.goto(server.url)
      await page.waitForFunction(() => window.__viewerHarness?.booted === true, null, { timeout: 30_000 })
      // Real automation (navigator.webdriver === true): refused by the collector.
      await page.getByRole('button', { name: 'Export result' }).click()
      await expect(page.getByText(/collector rejected the result.*automation/)).toBeVisible({ timeout: 30_000 })

      const exported = (await page.evaluate(() => window.__viewerHarness!.exportResult())) as unknown as Record<string, any>
      const post = (token: string, body: unknown) =>
        request.post(`${server.origin}/__harness/collect`, { headers: { 'x-harness-token': token, 'content-type': 'application/json' }, data: body })
      expect((await post('wrong-token', exported)).status()).toBe(401)
      const traversal = JSON.parse(JSON.stringify(exported)) as Record<string, any>
      traversal.environment.browser.name = '../../../../escape'
      expect((await post(server.token, traversal)).status()).toBe(400)
      const oversized = await post(server.token, { filler: 'x'.repeat(2 * 1024 * 1024) }).then(
        (response) => response.status(),
        () => 'aborted',
      )
      expect([413, 'aborted']).toContain(oversized)
      expect(readdirSync(server.output)).toEqual([])
    } finally {
      await server.stop()
    }
  })
})

// Corner navigation (A1 → XFD1048576) needs the S5 viewport integration; the
// interactive viewer does not exist yet, so this scenario is pinned as fixme.
test.fixme('corner navigation round-trip to XFD1048576', async () => {
  // Requires S5 viewport integration: Go To XFD1048576, verify recenter and
  // header agreement, then export a result containing the navigation sample.
})
