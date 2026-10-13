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
import { activeCell, cellIdOf, counts, ensureFixtures, goTo, openViewer, status } from './viewer-e2e'
import { createFixture } from '../helpers/xlsx-viewer-fixtures'

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
  expect(checkIdentity(body, { commit: identity.commit, root: ROOT })).toBeNull()
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

// ---------------------------------------------------------------------------
// S10.1: the real interactive viewer, end to end. Every scenario mounts the
// shipped <OfficeFile interactiveSheets> on the e2e page (src/debug/viewer-e2e)
// over a true XLSX package, on the resident (object) path and, where the
// fixture crosses an object cap, the worker-backed paged path (S5.4) with a
// real module Worker. WebKit runs are Playwright WebKit (`webkit-playwright`),
// never Safari evidence. No wall-clock assertion appears here (S10.2 owns
// performance); counters come from public boundaries (worker port, clipboard,
// onFrame stats, DOM).
// ---------------------------------------------------------------------------

test.describe('S10.1 real interactive viewer', () => {
  test.beforeAll(async () => { await ensureFixtures() })

  for (const path of [{ name: 'f1-small.xlsx', worker: false }, { name: 'f1-paged.xlsx', worker: true }] as const) {
    test(`F1 ${path.worker ? 'worker-backed' : 'resident'}: Go To XFD1048576, headers, find a never-painted corner, copy`, async ({ page }) => {
      await openViewer(page, path.name, { worker: path.worker })
      if (path.worker) expect((await counts(page)).open).toBe(1)
      else expect(await page.evaluate(() => window.__viewerE2E!.workersCreated)).toBe(0)

      await goTo(page, 'XFD1048576')
      await expect.poll(async () => (await activeCell(page))?.label ?? '').toMatch(/^XFD1048576, value 17179869183/)
      const corner = await activeCell(page)
      // The DOM id carries the exact arithmetic cell id 2^34-1 (no 32-bit packing).
      expect(cellIdOf(corner!.id)).toBe(17179869183)
      await expect(page.getByRole('columnheader', { name: /^Column XFD/ })).toBeVisible()
      await expect(page.getByRole('rowheader', { name: 'Row 1048576' })).toBeVisible()
      expect(await page.getByRole('grid').getAttribute('aria-rowcount')).toBe('1048577')
      expect(await page.getByRole('grid').getAttribute('aria-colcount')).toBe('16385')

      // No full-sheet canvas or CSS scroll element: every surface is viewport-bounded.
      const surfaces = await page.evaluate(() => {
        const dpr = window.devicePixelRatio
        const canvases = [...document.querySelectorAll('canvas')].map(c => ({ w: c.width, h: c.height }))
        const scroller = document.querySelector('[data-testid=officeview-sheet-scroller]')!.firstElementChild as HTMLElement
        return { dpr, vw: window.innerWidth, vh: window.innerHeight, canvases, extent: { w: scroller.offsetWidth, h: scroller.offsetHeight } }
      })
      for (const canvas of surfaces.canvases) {
        expect(canvas.w).toBeLessThanOrEqual(surfaces.vw * surfaces.dpr)
        expect(canvas.h).toBeLessThanOrEqual(surfaces.vh * surfaces.dpr)
      }
      expect(surfaces.extent.h).toBeLessThan(1048576 * 20) // a bounded physical extent, not 1M x row height
      const frame = await page.evaluate(() => window.__viewerE2E!.lastFrame!)
      expect(frame.peakBitmapBytes).toBeLessThanOrEqual(frame.budgetBytes)

      // Find a populated corner that was never painted (A1048576 holds its own
      // address as text; its cell id is 2^34 - 16384).
      await page.getByLabel('Search document').fill('A1048576')
      await expect.poll(async () => (await activeCell(page))?.label ?? '', { timeout: 30_000 }).toMatch(/^A1048576, value A1048576/)
      expect(cellIdOf((await activeCell(page))!.id)).toBe(17179852800)

      // Copy: the clipboard receives exactly the displayed value of the selection.
      await page.getByRole('button', { name: 'Copy selection' }).first().click()
      await expect.poll(() => page.evaluate(() => window.__viewerE2E!.clipboard.at(-1) ?? '')).toBe('A1048576')
      if (path.worker) expect((await counts(page)).open).toBe(1) // one parse/open per document, never per command
    })
  }

  // Reproduced gap: the column popup is anchored to the pane bottom and grew
  // past the pane top, where overflow:hidden clipped its first controls (Sort,
  // Clear column, Clear all filters) under the navigator: unreachable.
  test('column filter popup stays inside a short viewer pane and every control is reachable', async ({ page }) => {
    await openViewer(page, 'f2-1k.xlsx')
    await page.getByRole('button', { name: /^Filter column D/ }).click()
    const dialog = page.getByRole('dialog', { name: 'Column filter' })
    await expect(dialog.getByRole('checkbox', { name: 'alpha', exact: true })).toBeVisible()
    const rects = await page.evaluate(() => ({
      popup: document.querySelector('[data-officeview-ui=grid-popup]')!.getBoundingClientRect().toJSON(),
      pane: document.querySelector('[data-testid=officeview-sheet-pane]')!.getBoundingClientRect().toJSON(),
    }))
    expect(rects.popup.top).toBeGreaterThanOrEqual(rects.pane.top - 0.5)
    expect(rects.popup.bottom).toBeLessThanOrEqual(rects.pane.bottom + 0.5)
    await dialog.getByRole('button', { name: 'Clear all filters' }).click({ timeout: 10_000 })
    await dialog.getByRole('button', { name: 'Sort ascending' }).isVisible()
  })

  for (const path of [{ name: 'f2-1k.xlsx', rows: 1000, worker: false }, { name: 'f2-paged.xlsx', rows: 100_010, worker: true }] as const) {
    test(`F2 ${path.worker ? 'worker-backed' : 'resident'} (${path.rows} rows): header sort, value filter, cancel, clear`, async ({ page }) => {
      const handle = createFixture('F2', { rowCount: path.rows })
      let alpha = 0
      for (let row = 1; row < path.rows; row++) if (handle.cellAt(row, 3)?.value === 'alpha') alpha++
      await openViewer(page, path.name, { worker: path.worker })
      await expect(status(page)).toContainText(`Visible body ${path.rows} / complete body ${path.rows}`)

      // Header sort on G (code = 2*row+1): asc, then desc. The whole table is
      // the range (header protected), so the last body row leads.
      const sortG = page.getByRole('button', { name: 'Cycle sort column G' })
      await sortG.click()
      await expect(status(page)).toContainText('View-only sort applied', { timeout: 60_000 })
      await sortG.click()
      await expect(page.getByRole('columnheader', { name: /^Column G/ })).toHaveAttribute('aria-sort', 'descending', { timeout: 60_000 })
      const last = path.rows - 2
      await goTo(page, 'G2')
      await expect.poll(async () => (await activeCell(page))?.label ?? '').toBe(`G2, value ${(2 * last + 1).toFixed(2)}; original G${last + 1}`)
      expect(cellIdOf((await activeCell(page))!.id)).toBe(last * 16384 + 6) // source identity survives the sort
      await expect(status(page)).toContainText(`complete body ${path.rows - 1}`)

      // Value filter on D: only "alpha".
      await goTo(page, 'A1')
      await page.getByRole('button', { name: /^Filter column D/ }).click()
      const dialog = page.getByRole('dialog', { name: 'Column filter' })
      await dialog.getByRole('button', { name: 'Deselect All', exact: true }).click()
      await dialog.getByRole('checkbox', { name: 'alpha', exact: true }).check()
      await dialog.getByRole('button', { name: 'OK' }).click()
      await expect(status(page)).toContainText(`Visible body ${alpha} / complete body ${path.rows - 1}`, { timeout: 60_000 })
      const afterFilter = await counts(page)

      // Cancel publishes nothing: same records, no worker filter command.
      await page.getByRole('button', { name: /^Filter column D/ }).click()
      await dialog.getByRole('checkbox', { name: 'beta', exact: true }).check()
      await dialog.getByRole('button', { name: 'Cancel' }).click()
      await expect(dialog).toBeHidden()
      await expect(status(page)).toContainText(`Visible body ${alpha} / complete body ${path.rows - 1}`)
      expect((await counts(page)).filter ?? 0).toBe(afterFilter.filter ?? 0)

      // Clear all filters restores every body row; the sort stays applied.
      await page.getByRole('button', { name: /^Filter column D/ }).click()
      await dialog.getByRole('button', { name: 'Clear all filters' }).click()
      await expect(status(page)).toContainText(`Visible body ${path.rows - 1} / complete body ${path.rows - 1}`, { timeout: 60_000 })
      await expect(page.getByRole('columnheader', { name: /^Column G/ })).toHaveAttribute('aria-sort', 'descending')
      if (path.worker) {
        const final = await counts(page)
        expect(final).toMatchObject({ open: 1, sort: 2, filter: 2 })
        // Bounded column preparation: never more than eight dense key columns per command.
        expect(final.prepareColumns ?? 0).toBeLessThanOrEqual(final.sort + final.filter)
      }
    })
  }

  for (const path of [{ name: 'f3sparse-small.xlsx', worker: false }, { name: 'f3-paged.xlsx', worker: true }] as const) {
    test(`F3 ${path.worker ? 'worker-backed dense 8x16384' : 'resident sparse-wide'}: horizontal recenter keeps headers and cells in agreement`, async ({ page }) => {
      await openViewer(page, path.name, { worker: path.worker })
      if (path.worker) expect((await counts(page)).open).toBe(1)
      const agree = async () => page.evaluate(() => {
        // Every rendered body cell's column index has a rendered column header at the same index.
        const headers = new Set([...document.querySelectorAll('[role=columnheader]')].map(h => h.getAttribute('aria-colindex')))
        const cells = [...document.querySelectorAll('[role=gridcell]')]
        return { cells: cells.length, headers: headers.size, orphans: cells.filter(c => !headers.has(c.getAttribute('aria-colindex'))).length }
      })
      for (const address of ['XFD1', 'A1', 'XFD8', 'H3']) {
        await goTo(page, address)
        await expect.poll(async () => (await activeCell(page))?.label.split(',')[0] ?? '').toBe(address)
        const letters = address.replace(/\d+$/, '')
        await expect(page.getByRole('columnheader', { name: new RegExp(`^Column ${letters}\\b`) })).toBeVisible()
        const state = await agree()
        expect(state.orphans).toBe(0)
        // Viewport-only DOM: a bounded owned tree, never a row of 16,384 cells.
        expect(state.cells).toBeLessThan(2000)
        expect(state.headers).toBeLessThan(200)
      }
      // Keyboard edge navigation after recenter: Left from XFD8 reaches XFC8.
      await goTo(page, 'XFD8')
      await expect.poll(async () => (await activeCell(page))?.label.split(',')[0] ?? '').toBe('XFD8')
      await page.getByRole('grid').focus()
      await page.keyboard.press('ArrowLeft')
      await expect.poll(async () => (await activeCell(page))?.label.split(',')[0] ?? '').toBe('XFC8')
      if (path.worker) {
        const cell = await activeCell(page)
        expect(cell!.label).toMatch(/^XFC8, value \S/) // dense: a populated, loaded value at the far edge
      }
    })
  }

  test('F4 12,000 unique strings: searched Select All past 10,000, reopen retains, Select All restores', async ({ page }) => {
    await openViewer(page, 'f4-12k.xlsx')
    await expect(status(page)).toContainText('Visible body 12000 / complete body 12000')
    const dialog = page.getByRole('dialog', { name: 'Column filter' })
    await page.getByRole('button', { name: /^Filter column A/ }).click()
    await expect(dialog.getByText('Showing first 10,000; search all values.')).toBeVisible({ timeout: 30_000 })
    await dialog.getByRole('button', { name: 'Deselect All', exact: true }).click()
    await expect(dialog.getByText(/^0 selected values/)).toBeVisible()
    // item-11000 lies beyond the first 10,000 displayed values; search finds it.
    await dialog.getByRole('textbox', { name: 'Search all values' }).fill('item-11000')
    await expect(dialog.getByRole('checkbox', { name: 'item-11000', exact: true })).toBeVisible({ timeout: 30_000 })
    await dialog.getByRole('button', { name: 'Select All', exact: true }).click()
    await expect(dialog.getByText(/^1 selected values; complete universe/)).toBeVisible()
    await dialog.getByRole('button', { name: 'OK' }).click()
    await expect(status(page)).toContainText('Visible body 1 / complete body 12000', { timeout: 30_000 })
    // Reopen: the applied choice is retained; Cancel publishes nothing.
    await page.getByRole('button', { name: /^Filter column A/ }).click()
    await expect(dialog.getByText(/^1 selected values; complete universe/)).toBeVisible({ timeout: 30_000 })
    await dialog.getByRole('button', { name: 'Cancel' }).click()
    await expect(status(page)).toContainText('Visible body 1 / complete body 12000')
    // First (unsearched) Select All covers the whole 12,000-value universe, not the 10,000 shown.
    await page.getByRole('button', { name: /^Filter column A/ }).click()
    await dialog.getByRole('button', { name: 'Select All', exact: true }).click()
    await expect(dialog.getByText(/^12000 selected values; complete universe/)).toBeVisible({ timeout: 30_000 })
    await dialog.getByRole('button', { name: 'OK' }).click()
    await expect(status(page)).toContainText('Visible body 12000 / complete body 12000', { timeout: 30_000 })
  })

  test('F5 structure: giant narrow-column merge and fixed array refuse sorts; child selection allowed', async ({ page }) => {
    await openViewer(page, 'f5-structure.xlsx')
    const before = await page.getByRole('gridcell', { name: /^A2, value / }).getAttribute('aria-label')
    // Header sort from A1: the current region reaches the fixed array C30:C32 and the
    // giant merge D1:D1048576, so the whole-row sort is refused, never partially applied.
    await page.getByRole('button', { name: 'Cycle sort column A' }).click()
    await expect(status(page)).toContainText('Operation refused: merge', { timeout: 30_000 })
    await expect(page.getByRole('columnheader', { name: /^Column A/ })).not.toHaveAttribute('aria-sort', /ascending|descending/)
    expect(await page.getByRole('gridcell', { name: /^A2, value / }).getAttribute('aria-label')).toBe(before)
    // An explicit A2:B7 selection: narrow keys, but the merge D1:D1048576 spans the
    // moved row band, so the band guard refuses too (AC-SEM-04).
    await goTo(page, 'A2')
    await page.getByRole('grid').focus()
    for (let i = 0; i < 5; i++) await page.keyboard.press('Shift+ArrowDown')
    await page.keyboard.press('Shift+ArrowRight')
    await page.getByRole('button', { name: 'Cycle sort column A' }).click()
    await expect(status(page)).toContainText('Operation refused: merge', { timeout: 30_000 })
    expect(await page.getByRole('gridcell', { name: /^A2, value / }).getAttribute('aria-label')).toBe(before)
    // Selecting a read-only array child is allowed.
    await goTo(page, 'C31')
    await expect.poll(async () => (await activeCell(page))?.label ?? '').toBe('C31, value 2')
  })

  test('F7 two sheets with identical first revisions: views, zoom and hidden-row inspection stay per sheet; reload resets', async ({ page }) => {
    await openViewer(page, 'f7-two-sheets.xlsx')
    // Reproduced gap: the 0.75pt row 10 (2 device px at DPR 2) made the whole
    // owned grid coarse. Now only that row is summarized; the rest stay owned.
    await expect(page.getByRole('rowheader', { name: 'Row 11' })).toBeVisible()
    await expect(page.getByRole('rowheader', { name: 'Row 10', exact: true })).toHaveCount(0)
    expect(await page.getByRole('gridcell').count()).toBeGreaterThan(20)
    const sheet = page.getByRole('combobox', { name: 'Worksheet' })
    const zoomLabel = page.getByRole('button', { name: 'Reset zoom' })
    await page.getByRole('button', { name: 'Zoom in' }).click()
    const zoomed = await zoomLabel.textContent()
    expect(zoomed).not.toBe('100%')
    const fontGen = await page.evaluate(() => window.__viewerE2E!.lastFrame!.identity.fontGen)
    await goTo(page, 'A1') // zoom anchors at the pane centre; bring column A back into view

    // Sheet A: sort score ascending (score falls as id rises, so id 40 leads).
    await page.getByRole('button', { name: 'Cycle sort column B' }).click()
    await expect(page.getByRole('columnheader', { name: /^Column B/ })).toHaveAttribute('aria-sort', 'ascending', { timeout: 30_000 })
    await goTo(page, 'A2')
    await expect.poll(async () => (await activeCell(page))?.label ?? '').toBe('A2, value 40; original A41')
    expect(await zoomLabel.textContent()).toBe(zoomed) // a committed view keeps zoom (§10.1)

    // Sheet B (also revision 1 after its own sort) must not adopt A's view or caches.
    await sheet.selectOption({ label: 'B' })
    await expect(page.getByRole('grid')).toHaveAttribute('aria-label', 'Worksheet B')
    await expect(page.getByRole('columnheader', { name: /^Column B/ })).not.toHaveAttribute('aria-sort', /ascending|descending/)
    await goTo(page, 'A1')
    await page.getByRole('button', { name: 'Cycle sort column A' }).click()
    await page.getByRole('button', { name: 'Cycle sort column A' }).click()
    await expect(page.getByRole('columnheader', { name: /^Column A/ })).toHaveAttribute('aria-sort', 'descending', { timeout: 30_000 })
    await goTo(page, 'B2')
    await expect.poll(async () => (await activeCell(page))?.label ?? '').toBe('B2, value 1000; original B41')

    // Back to A: its own sort and data, not B's.
    await sheet.selectOption({ label: 'A' })
    await expect(page.getByRole('grid')).toHaveAttribute('aria-label', 'Worksheet A')
    await expect(page.getByRole('columnheader', { name: /^Column B/ })).toHaveAttribute('aria-sort', 'ascending')
    await goTo(page, 'C2')
    await expect.poll(async () => (await activeCell(page))?.label ?? '').toBe('C2, value s0-40; original C41')
    expect(await page.evaluate(() => window.__viewerE2E!.lastFrame!.identity.fontGen)).toBe(fontGen) // views never bump fonts

    // Authored-hidden source row 8 (id 7): Original-cell Go To opens the read-only
    // inspector without unhiding it or making it the active descendant.
    const activeBefore = await activeCell(page)
    await goTo(page, 'A8', 'source')
    await expect(page.getByRole('status').filter({ hasText: 'Original cell A8' })).toContainText('read-only')
    expect((await activeCell(page))?.id).toBe(activeBefore?.id)

    // Reload replaces the document: zoom resets to 100% and no view survives.
    await page.reload()
    await expect(status(page)).toContainText('Visible body', { timeout: 60_000 })
    await expect(zoomLabel).toHaveText('100%')
    await expect(page.getByRole('columnheader', { name: /^Column B/ })).not.toHaveAttribute('aria-sort', /ascending|descending/)
  })


  // S13 native Excel checkpoint: a workbook written and reopened (no repair
  // prompt) by Microsoft Excel 16 for macOS. Expected labels are the values
  // Excel displayed. The file is far below every object cap, so plain
  // ?worker=1 still admits it resident; ?paged=1 is the debug-only admission
  // override of the e2e page (src/debug/forcePaged.ts) that sends it down the
  // worker-backed paged path.
  for (const leg of [
    { name: 'resident', worker: false, paged: false },
    { name: '?worker=1, admitted resident', worker: true, paged: false },
    { name: 'worker-backed paged (?worker=1&paged=1)', worker: true, paged: true },
  ] as const) {
    test(`native Excel native-corner.xlsx ${leg.name}: cached values, error, date display, Unicode, hidden row, far corner`, async ({ page }) => {
      await openViewer(page, 'native-corner.xlsx', { worker: leg.worker, paged: leg.paged })
      expect((await counts(page)).open ?? 0).toBe(leg.paged ? 1 : 0)
      if (leg.paged) await expect(page.getByRole('status').filter({ hasText: 'Showing saved values' }).first()).toBeVisible()
      await expect(page.getByRole('grid')).toHaveAttribute('aria-label', 'Worksheet 工作表1')
      await expect(page.getByRole('alert')).toHaveCount(0)

      const expectLabel = async (address: string, label: string) => {
        await goTo(page, address)
        await expect.poll(async () => (await activeCell(page))?.label ?? '').toBe(label)
      }
      await expectLabel('A1', 'A1, value 42')
      await expectLabel('A2', 'A2, value 2024-01-15 12:00')
      await expectLabel('A3', 'A3, value #DIV/0!')
      await expectLabel('A4', 'A4, value 84')
      await expectLabel('E3', 'E3, value 2')
      // Row 5 is authored-hidden: no row header for it, and moving down from A4 skips to A6.
      await expect(page.getByRole('rowheader', { name: 'Row 5', exact: true })).toHaveCount(0)
      await goTo(page, 'A4')
      await page.getByRole('grid').focus()
      await page.keyboard.press('ArrowDown')
      await expect.poll(async () => (await activeCell(page))?.label ?? '').toBe('A6, value Ünïcødé 漢字 🎉')
      // Its source cell is reachable only through the read-only Original-cell inspector.
      await goTo(page, 'A5', 'source')
      await expect(page.getByRole('status').filter({ hasText: 'Original cell A5' })).toContainText('read-only')

      await expectLabel('XFD1048576', 'XFD1048576, value corner')
      expect(cellIdOf((await activeCell(page))!.id)).toBe(17179869183)
      await expect(page.getByRole('rowheader', { name: 'Row 1048576' })).toBeVisible()
      if (leg.paged) expect((await counts(page)).open).toBe(1) // one open, every read went to the worker
    })
  }

  // S13: native-large.xlsx (Excel 16 for macOS, 200,001 cells) crosses the
  // object caps, so plain ?worker=1 (no debug override) opens it worker-backed.
  test('native Excel native-large.xlsx worker-backed without override: corners, date display, descending header sort', async ({ page }) => {
    await openViewer(page, 'large/native-large.xlsx', { worker: true })
    expect(await page.evaluate(() => window.__viewerE2E!.workersCreated)).toBe(1)
    expect((await counts(page)).open).toBe(1)
    await expect(page.getByRole('status').filter({ hasText: 'Showing saved values' }).first()).toBeVisible()
    await expect(page.getByRole('alert')).toHaveCount(0)
    const expectLabel = async (address: string, label: string) => {
      await goTo(page, address)
      await expect.poll(async () => (await activeCell(page))?.label ?? '', { timeout: 30_000 }).toBe(label)
    }
    await expectLabel('A50000', 'A50000, value 500001')
    await expectLabel('D50000', 'D50000, value 2160-11-23 12:00')
    await expectLabel('E50000', 'E50000, value last')
    await expectLabel('D2', 'D2, value 2024-01-03 12:00')

    await goTo(page, 'A1')
    const sortA = page.getByRole('button', { name: 'Cycle sort column A' })
    await sortA.click()
    await expect(page.getByText('View-only sort applied; formulas unchanged.').first()).toBeAttached({ timeout: 60_000 })
    await sortA.click()
    await expect(page.getByRole('columnheader', { name: /^Column A/ })).toHaveAttribute('aria-sort', 'descending', { timeout: 60_000 })
    await expectLabel('A1', 'A1, value 500001; original A50000')
    await expectLabel('E1', 'E1, value last; original E50000')
    expect((await counts(page)).open).toBe(1)
  })

  // S13: native-1904.xlsx (Excel 16 for macOS, date1904 on): the painted labels use the 1904 system.
  test('native Excel native-1904.xlsx resident: 1904 serials display as Excel showed them', async ({ page }) => {
    await openViewer(page, 'native-1904.xlsx')
    for (const [address, text] of [['A2', '2024-01-15 12:00:00'], ['A4', '2024-01-16 23:59:59'], ['A5', '2024-02-29 06:00:00'], ['A6', '1904-01-01 00:00:00'], ['A7', '2023-12-31 12:00:00']] as const) {
      await goTo(page, address)
      await expect.poll(async () => (await activeCell(page))?.label ?? '').toBe(`${address}, value ${text}`)
    }
  })
})
