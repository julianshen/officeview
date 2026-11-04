/**
 * S13 runtime device-pixel-ratio change in installed desktop Chrome
 * (spec §5.4/§10.1/§11.1 and the F7 "DPR changes" stress row; AC-CAP-02).
 *
 * EVIDENCE LABEL: emulated runtime DPR change (CDP) in desktop Chrome — not a
 * physical display move. The device scale is changed mid-session with
 * Emulation.setDeviceMetricsOverride (deviceScaleFactor 1 → 2 → 1.5 → 1) at a
 * constant 1440×900 CSS viewport, which is what moving a window between a DPR 1
 * and a DPR 2 display does to the page. This machine has one DPR 1 display, so
 * a physical move was not possible.
 *
 * Two questions per fixture (F1 resident and F1 worker-backed paged):
 *  1. Does a DPR-only change re-rasterize the viewer by itself? Before the
 *     D-DPR fix the viewer's only size input was a content-box ResizeObserver
 *     on the pane, so a DPR change with an unchanged CSS size woke nothing and
 *     this test failed. useDevicePixelRatio (src/components/xlsx/
 *     devicePixelRatio.ts) now follows the ratio. Under CDP emulation Chrome
 *     dispatches no resize/media-query event, so here it is the hook's
 *     visible-page poll that fires (the logged page signals stay empty).
 *  2. Once the viewer does render at the new DPR (driven here by a 100 px pane
 *     width change and back, the input it does listen to), are the raster, tile keys,
 *     alignment, view state, budget and console all correct?
 */
import { expect, test, type CDPSession, type Page } from '@playwright/test'
import { execFileSync } from 'node:child_process'
import type { SheetFrameStats } from '../../src/components/xlsx/InteractiveSheetDoc'
import { activeCell, ensureFixtures, goTo, openViewer, status, type FixtureName } from './viewer-e2e'

const VIEW = { width: 1440, height: 900 }
const STEPS = [2, 1.5, 1] as const
const TILE_DEVICE_PX = 512
const NUDGE_PX = 100
const LABEL = 'emulated runtime DPR change (CDP) in desktop Chrome — not a physical display move'

test.use({ viewport: VIEW, deviceScaleFactor: 1, channel: 'chrome' })
// File-level so the webkit project never launches with channel 'chrome' (an unsupported pair).
test.skip(({ browserName }) => browserName !== 'chromium', 'CDP device-metrics emulation and channel chrome are Chromium-only')
test.beforeAll(async () => { await ensureFixtures() })

interface DprProbe {
  events: string[]
  frames: SheetFrameStats[]
}
declare global {
  interface Window { __dprProbe?: DprProbe }
}

const FIXTURES: ReadonlyArray<{ name: FixtureName; worker: boolean }> = [
  { name: 'f1-small.xlsx', worker: false },
  { name: 'f1-paged.xlsx', worker: true },
]

function installedChromeVersion(): string | null {
  try {
    return execFileSync('plutil', ['-extract', 'CFBundleShortVersionString', 'raw', '/Applications/Google Chrome.app/Contents/Info.plist'], { encoding: 'utf8' }).trim()
  } catch {
    return null
  }
}

/** The browser under test must be the installed Google Chrome, not bundled Chromium. */
async function chromeIdentity(cdp: CDPSession): Promise<string> {
  const { product } = await cdp.send('Browser.getVersion') as { product: string }
  const installed = installedChromeVersion()
  expect(installed, 'Google Chrome must be installed in /Applications').not.toBeNull()
  expect(product).toBe(`Chrome/${installed}`)
  return product
}

/** Records DPR-change signals and every onFrame report from here on. */
async function installProbe(page: Page): Promise<void> {
  await page.evaluate(() => {
    const probe: DprProbe = { events: [], frames: [] }
    window.__dprProbe = probe
    window.addEventListener('resize', () => probe.events.push(`resize@${devicePixelRatio}`))
    for (const query of ['(resolution: 1dppx)', '(resolution: 1.5dppx)', '(resolution: 2dppx)']) {
      matchMedia(query).addEventListener('change', event => probe.events.push(`change ${query}=${event.matches}`))
    }
    const state = window.__viewerE2E!
    let last = state.lastFrame
    Object.defineProperty(state, 'lastFrame', {
      configurable: true,
      get: () => last,
      set: (frame: SheetFrameStats) => { last = frame; probe.frames.push(JSON.parse(JSON.stringify(frame)) as SheetFrameStats) },
    })
  })
}

async function setScale(cdp: CDPSession, deviceScaleFactor: number, width = VIEW.width): Promise<void> {
  await cdp.send('Emulation.setDeviceMetricsOverride', { width, height: VIEW.height, deviceScaleFactor, mobile: false })
}

async function surfaces(page: Page) {
  return page.evaluate(() => {
    const read = (id: string) => {
      const canvas = document.querySelector<HTMLCanvasElement>(`[data-testid=${id}]`)!
      const rect = canvas.getBoundingClientRect()
      return { width: canvas.width, height: canvas.height, cssWidth: rect.width, cssHeight: rect.height }
    }
    return { dpr: devicePixelRatio, base: read('officeview-sheet-base'), overlay: read('officeview-sheet-overlay') }
  })
}

/** View state that a DPR change must not disturb (§5.4: selected IDs, zoom,
 * active focus, visible data). Raw scrollTop is not compared: the scroll window
 * is rebuilt at the new device scale from the same logical origin. */
async function viewState(page: Page) {
  const active = await activeCell(page)
  const dom = await page.evaluate(() => {
    const labels = (role: string) => [...document.querySelectorAll(`[role=${role}]`)].map(e => e.getAttribute('aria-label') ?? '')
    return {
      selected: [...document.querySelectorAll('[role=gridcell][aria-selected=true]')].map(c => c.getAttribute('aria-label')).sort(),
      firstRowHeader: labels('rowheader').find(l => l.startsWith('Row ')) ?? null,
      firstColumnHeader: labels('columnheader').find(l => l.startsWith('Column ')) ?? null,
    }
  })
  return {
    active: active?.label ?? null,
    ...dom,
    zoom: await page.getByRole('button', { name: 'Reset zoom' }).textContent(),
    find: await page.getByLabel('Search document').inputValue(),
    status: await status(page).textContent(),
  }
}

/** Sets up find, selection, active cell and zoom before any DPR change. */
async function prepareState(page: Page): Promise<void> {
  await page.getByLabel('Search document').fill('A1048576')
  await expect.poll(async () => (await activeCell(page))?.label ?? '', { timeout: 30_000 }).toMatch(/^A1048576, value A1048576/)
  await goTo(page, 'C3')
  await expect.poll(async () => (await activeCell(page))?.label.split(',')[0] ?? '').toBe('C3')
  await page.getByRole('grid').focus()
  await page.keyboard.press('Shift+ArrowDown')
  await page.keyboard.press('Shift+ArrowRight')
  await expect.poll(() => page.locator('[role=gridcell][aria-selected=true]').count()).toBe(4)
  await page.keyboard.press('+')
  await expect(page.getByRole('button', { name: 'Reset zoom' })).not.toHaveText('100%')
}

/**
 * Gridline probe: the base canvas must carry the vertical gridline at the right
 * edge, and the horizontal one at the bottom edge, of an empty owned gridcell
 * whose DOM geometry is the hit-test geometry. Scans ±3 device px around the
 * expected device-pixel position for the darkest column/row.
 */
async function gridlineProbe(page: Page) {
  return page.evaluate(() => {
    const canvas = document.querySelector<HTMLCanvasElement>('[data-testid=officeview-sheet-base]')!
    const ctx = canvas.getContext('2d')!
    const box = canvas.getBoundingClientRect()
    const sx = canvas.width / box.width, sy = canvas.height / box.height
    const empty = [...document.querySelectorAll<HTMLElement>('[role=gridcell]')].filter(cell => {
      // A blank cell is announced with an empty value ("E4, value").
      const label = (cell.getAttribute('aria-label') ?? '').trimEnd()
      if (!/^[A-Z]+\d+, value$/.test(label) || cell.getAttribute('aria-selected') === 'true') return false
      const r = cell.getBoundingClientRect()
      return r.left > box.left + 80 && r.right < box.right - 80 && r.top > box.top + 40 && r.bottom < box.bottom - 40
    })
    const cell = empty[Math.floor(empty.length / 2)]
    if (!cell) return { cell: null }
    const r = cell.getBoundingClientRect()
    const lum = (d: Uint8ClampedArray, i: number) => 0.299 * d[i] + 0.587 * d[i + 1] + 0.114 * d[i + 2]
    const darkest = (values: number[], start: number) => {
      let at = 0
      for (let i = 1; i < values.length; i++) if (values[i] < values[at]) at = i
      return { at: start + at, lum: values[at] }
    }
    const ex = (r.right - box.left) * sx, ey = (r.bottom - box.top) * sy
    const cy = Math.round((r.top + r.height / 2 - box.top) * sy), cx = Math.round((r.left + r.width / 2 - box.left) * sx)
    const x0 = Math.round(ex) - 3, y0 = Math.round(ey) - 3
    const row = ctx.getImageData(x0, cy, 7, 1).data, col = ctx.getImageData(cx, y0, 1, 7).data
    const center = ctx.getImageData(cx, cy, 1, 1).data
    const v = darkest([0, 1, 2, 3, 4, 5, 6].map(i => lum(row, i * 4)), x0)
    const h = darkest([0, 1, 2, 3, 4, 5, 6].map(i => lum(col, i * 4)), y0)
    return { cell: cell.getAttribute('aria-label'), expectX: ex, vertical: v, expectY: ey, horizontal: h, centerLum: lum(center, 0) }
  })
}

const frames = (page: Page) => page.evaluate(() => window.__dprProbe!.frames.splice(0))
const events = (page: Page) => page.evaluate(() => window.__dprProbe!.events.splice(0))

function collectErrors(page: Page): string[] {
  const errors: string[] = []
  page.on('console', message => { if (message.type() === 'error') errors.push(message.text()) })
  page.on('pageerror', error => errors.push(`pageerror: ${error.message}`))
  return errors
}

test.describe(`S13 runtime DPR change — ${LABEL}`, () => {
  for (const fixture of FIXTURES) {
    const kind = fixture.worker ? 'worker-backed paged' : 'resident'

    test(`F1 ${kind}: a DPR-only change re-rasterizes without any other input`, async ({ page, browserName }) => {
      test.skip(browserName !== 'chromium', 'CDP device-metrics emulation and channel chrome are Chromium-only')
      const errors = collectErrors(page)
      await openViewer(page, fixture.name, { worker: fixture.worker })
      const cdp = await page.context().newCDPSession(page)
      const product = await chromeIdentity(cdp)
      await prepareState(page)
      await installProbe(page)
      for (const dsf of STEPS) {
        await setScale(cdp, dsf)
        await expect.poll(() => page.evaluate(() => devicePixelRatio)).toBe(dsf)
        let last = await surfaces(page)
        try {
          await expect.poll(async () => {
            last = await surfaces(page)
            return last.base.width === Math.round(last.base.cssWidth * dsf) && last.overlay.width === Math.round(last.overlay.cssWidth * dsf)
          }, { timeout: 5_000 }).toBe(true)
        } finally {
          console.log(`[dpr-evidence] ${LABEL}; ${product}; ${fixture.name}; unprompted step dsf=${dsf}: base ${last.base.width}x${last.base.height} for css ${last.base.cssWidth}x${last.base.cssHeight}, overlay ${last.overlay.width}x${last.overlay.height}; viewer frames since change=${(await frames(page)).length}; page signals=${JSON.stringify(await events(page))}`)
        }
      }
      expect(errors).toEqual([])
    })

    test(`F1 ${kind}: after re-rendering at the new DPR the raster, state, budget and alignment are correct`, async ({ page, browserName }) => {
      test.skip(browserName !== 'chromium', 'CDP device-metrics emulation and channel chrome are Chromium-only')
      const errors = collectErrors(page)
      await openViewer(page, fixture.name, { worker: fixture.worker })
      const cdp = await page.context().newCDPSession(page)
      const product = await chromeIdentity(cdp)
      await prepareState(page)
      const before = await viewState(page)
      await installProbe(page)
      const receipt: Record<string, unknown>[] = []
      for (const dsf of STEPS) {
        await setScale(cdp, dsf)
        await expect.poll(() => page.evaluate(() => devicePixelRatio)).toBe(dsf)
        // The nudge: a pane width change and back, which the pane's
        // ResizeObserver does deliver. It touches no view state under test.
        // Each leg waits until the viewer has laid its canvas out at that width
        // (100 px, not 1 px: a 1 px leg was sometimes absorbed by clientWidth
        // rounding at a fractional DPR and the viewer never re-rendered).
        const cssWidth = async () => (await surfaces(page)).base.cssWidth
        const settled = await cssWidth()
        await setScale(cdp, dsf, VIEW.width - NUDGE_PX)
        await expect.poll(cssWidth).toBe(settled - NUDGE_PX)
        await setScale(cdp, dsf, VIEW.width)
        await expect.poll(cssWidth).toBe(settled)

        // Backing stores = CSS size × DPR (within rounding) on both layers.
        const within = (layer: { width: number; height: number; cssWidth: number; cssHeight: number }) =>
          Math.abs(layer.width - layer.cssWidth * dsf) <= 1 && Math.abs(layer.height - layer.cssHeight * dsf) <= 1
        await expect.poll(async () => {
          const now = await surfaces(page)
          return { dsf, base: within(now.base) ? 'ok' : now.base, overlay: within(now.overlay) ? 'ok' : now.overlay }
        }, { timeout: 15_000 }).toEqual({ dsf, base: 'ok', overlay: 'ok' })
        await expect.poll(() => page.evaluate(() => window.__viewerE2E!.lastFrame!.pendingTiles), { timeout: 60_000 }).toBe(0)
        const s = await surfaces(page)
        const seen = await frames(page)
        expect(seen.length, 'the viewer reported frames after the change').toBeGreaterThan(0)
        const zoom = seen.at(-1)!.zoom
        // The first frame that rasterized at the new DPR: its tile requests are
        // 512-device-px tiles at zoom × dsf, and it reused no tile (a new DPR is
        // a re-admitted pool with new tile keys).
        const atNewScale = (f: SheetFrameStats) => f.lastRequests.length > 0 && f.lastRequests.every(b => Math.abs(b.width * f.zoom * dsf - TILE_DEVICE_PX) < 1e-6 && Math.abs(b.height * f.zoom * dsf - TILE_DEVICE_PX) < 1e-6)
        const firstNew = seen.find(atNewScale)
        expect(firstNew, 'a frame rasterized at the new DPR').toBeDefined()
        expect(firstNew!.tileHits).toBe(0)
        // No stale-DPR tile is ever presented after the change: every frame that
        // painted tiles from firstNew on painted them at the new scale.
        for (const frame of seen.slice(seen.indexOf(firstNew!))) if (frame.lastRequests.length) expect(atNewScale(frame)).toBe(true)
        // Bitmap budget accounting: every frame within budget; tile bytes = tiles × 1 MiB.
        for (const frame of seen) {
          expect(frame.bitmapBytes).toBeLessThanOrEqual(frame.budgetBytes)
          expect(frame.peakBitmapBytes).toBeLessThanOrEqual(frame.budgetBytes)
          expect(frame.tiles.bytes).toBe(frame.tiles.tiles * TILE_DEVICE_PX * TILE_DEVICE_PX * 4)
        }
        // Text/grid alignment: the painted gridline sits on the DOM (hit-test) cell edge.
        const probe = await gridlineProbe(page)
        expect(probe.cell, 'an empty owned gridcell inside the pane').not.toBeNull()
        expect(Math.abs(probe.vertical!.at - probe.expectX!)).toBeLessThanOrEqual(1.5)
        expect(Math.abs(probe.horizontal!.at - probe.expectY!)).toBeLessThanOrEqual(1.5)
        expect(probe.vertical!.lum).toBeLessThan(245)
        expect(probe.horizontal!.lum).toBeLessThan(245)
        expect(probe.centerLum!).toBeGreaterThan(250)
        // View state preserved.
        expect(await viewState(page)).toEqual(before)
        const last = seen.at(-1)!
        receipt.push({ dsf, base: `${s.base.width}x${s.base.height}`, overlay: `${s.overlay.width}x${s.overlay.height}`, css: `${s.base.cssWidth}x${s.base.cssHeight}`, zoom, framesSinceChange: seen.length, firstNewFrameTileHits: firstNew!.tileHits, tileBoxLogical: firstNew!.lastRequests[0].width, bitmapBytes: last.bitmapBytes, peakBitmapBytes: last.peakBitmapBytes, budgetBytes: last.budgetBytes, gridline: { cell: probe.cell, x: [probe.expectX, probe.vertical!.at], y: [probe.expectY, probe.horizontal!.at] } })
      }
      expect(errors).toEqual([])
      console.log(`[dpr-evidence] ${LABEL}; ${product}; ${fixture.name}${fixture.worker ? ' (worker)' : ''}; nudged by a 100px pane width change and back; state=${JSON.stringify(before)}; steps=${JSON.stringify(receipt)}`)
    })
  }
})
