/**
 * Stage 4a Task 14: a streamed XLSX in the interactive viewer, tiles mode, in a real browser against
 * server/test/app.ts. The harness page maps `?sheets=tiles` to `streamSheets: 'tiles'` (the client HELLO picks the
 * mode) and exposes the viewer's frame stats as `window.__ovSheetStats` (the last frame, plus the frame count and the
 * running `basePaints` total). Fixtures (`?doc=`): `million-rows` (A1:D2000 and A1048576), `grid500x100`, and
 * `frozen` (2 frozen columns and 3 frozen rows), served under the demo's visible-only policy.
 *
 * Tiles are real PNGs decoded by `createImageBitmap`. No fixed sleeps: every wait polls the stats or the DOM.
 */
import { expect, test, type Page } from '@playwright/test'
import { readFileSync } from 'node:fs'
import { resolve } from 'node:path'

const P = Number(process.env.OV_STREAM_PORT ?? 4181)
const open = (page: Page, q: string) => page.goto(`/src/debug/stream-viewer.html?port=${P}&sheets=tiles&${q}`)
/** The verbatim notice of the local static export (src/render/paint.ts); read from source so it is never retyped. */
const LARGE_EXPORT_MESSAGE = /export const LARGE_EXPORT_MESSAGE = '([^']+)'/.exec(readFileSync(resolve(process.cwd(), 'src/render/paint.ts'), 'utf8'))![1]!
const SHEET_ZOOM_STEPS = [0.25, 0.33, 0.5, 0.67, 0.75, 1, 1.25, 1.5, 2, 3, 4]

type Stats = { frames: number; basePaints: number; basePaintsTotal: number; pendingTiles: number; tileHits: number; zoom: number; activeSheetId: string }
const stats = (page: Page) => page.evaluate(() => (window as unknown as { __ovSheetStats?: Stats }).__ovSheetStats ?? null)
const scroller = (page: Page) => page.getByTestId('officeview-sheet-scroller')
/** Settled: the latest frame shows server tiles and nothing is pending. */
async function settled(page: Page) {
  await expect.poll(async () => {
    const s = await stats(page)
    return s ? s.frames > 0 && s.pendingTiles === 0 && s.tileHits > 0 : false
  }, { timeout: 60_000 }).toBe(true)
}

test.beforeEach(async ({ page }) => {
  page.on('pageerror', (e) => { throw e })
})

test('million-row sheet scrolls to row 1,048,576', async ({ page }) => {
  await open(page, 'doc=million-rows')
  await expect(page.getByRole('grid')).toBeVisible()
  await settled(page)
  // No keyboard (selection is 4b): the scroller is driven to its end directly. The physical scroll range is capped
  // (MAX_PHYSICAL_SCROLL_CSS_PX), so each end-of-segment lands, re-centres the window and moves the end further: the
  // end is re-applied on every poll until the last row's header is rendered.
  await expect.poll(async () => {
    await scroller(page).evaluate((s) => { s.scrollTop = s.scrollHeight - s.clientHeight })
    return page.locator('[role="rowheader"][aria-label="Row 1048576"]').count()
  }, { timeout: 60_000, intervals: [50] }).toBeGreaterThan(0)
  await expect(page.locator('[role="rowheader"][aria-label="Row 1048576"]')).toContainText('1048576')
  await settled(page)
  const s = (await stats(page))!
  expect(s.pendingTiles).toBe(0)
  expect(s.basePaintsTotal).toBe(0)
})

test('no too-large notice', async ({ page }) => {
  await open(page, 'doc=million-rows')
  await settled(page)
  expect(LARGE_EXPORT_MESSAGE.length).toBeGreaterThan(20)
  await expect(page.locator('body')).not.toContainText(LARGE_EXPORT_MESSAGE)
  // Nor any unit canvas: a tiles session never shows the Stage 3 units.
  await expect(page.locator('canvas[data-unit]')).toHaveCount(0)
})

test('no local painting', async ({ page }) => {
  await open(page, 'doc=grid500x100')
  await settled(page)
  await scroller(page).evaluate((s) => { s.scrollTop = 3000; s.scrollLeft = 2000 })
  await expect.poll(async () => (await stats(page))?.frames ?? 0).toBeGreaterThan(1)
  await settled(page)
  const s = (await stats(page))!
  expect(s.basePaints).toBe(0)
  expect(s.basePaintsTotal).toBe(0)
})

test('authored frozen panes stay pinned', async ({ page }) => {
  await open(page, 'doc=frozen')
  await settled(page)
  const grid = page.getByRole('grid')
  const top = async (label: string) => {
    const header = page.locator(`[role="rowheader"][aria-label="${label}"]`)
    if (await header.count() === 0) return null
    const g = (await grid.boundingBox())!, r = await header.boundingBox()
    return r ? r.y - g.y : null
  }
  const pinned = await top('Row 1')
  const body = await top('Row 4')
  expect(pinned).not.toBeNull()
  expect(body).not.toBeNull()
  await scroller(page).evaluate((s) => { s.scrollTop += 2000 })
  // The body moved away under the pinned strip; the pinned header row did not move.
  await expect.poll(() => top('Row 4')).not.toBe(body)
  await settled(page)
  expect(await top('Row 1')).toBe(pinned)
  expect(await top('Row 3')).not.toBeNull()
})

test('wheel zoom settles to a preset step', async ({ page }) => {
  await open(page, 'doc=grid500x100')
  await settled(page)
  const label = page.getByRole('button', { name: 'Reset zoom' })
  await expect(label).toHaveText('100%')
  const pane = (await page.getByTestId('officeview-sheet-pane').boundingBox())!
  await page.mouse.move(pane.x + pane.width / 2, pane.y + pane.height / 2)
  await page.keyboard.down('Control')
  await page.mouse.wheel(0, -120)
  await page.mouse.wheel(0, -45)
  await page.keyboard.up('Control')
  await expect.poll(async () => (await label.textContent())?.trim()).not.toBe('100%')
  await settled(page)
  const zoom = (await stats(page))!.zoom
  expect(SHEET_ZOOM_STEPS).toContain(zoom)
  expect(zoom).not.toBe(1)
  const shown = Number((await label.textContent())!.replace('%', '')) / 100
  expect(SHEET_ZOOM_STEPS.map((z) => Math.round(z * 100) / 100)).toContain(shown)
})

test.describe('print', () => {
  type Rec = { frames: number; windowPrints: number }
  const prints = (page: Page) => page.evaluate(() => (window as unknown as { __ovPrints: Rec }).__ovPrints)
  test.beforeEach(async ({ page }) => {
    // As stream-print.spec.ts: the print frame's print() is stubbed and counted; window.print must never run.
    await page.addInitScript(() => {
      const rec = { frames: 0, windowPrints: 0 }
      ;(window as unknown as { __ovPrints: Rec }).__ovPrints = rec
      window.print = () => { rec.windowPrints++ }
      new MutationObserver((records) => {
        for (const r of records) for (const n of r.addedNodes) {
          if (!(n instanceof HTMLIFrameElement) || !n.hasAttribute('data-officeview-print-frame')) continue
          const w = n.contentWindow as (Window & { __ovStubbed?: true }) | null
          if (!w || w.__ovStubbed) continue
          w.__ovStubbed = true
          w.print = () => { rec.frames++ }
        }
      }).observe(document, { childList: true, subtree: true })
    })
  })

  test('print prints the static unit', async ({ page }) => {
    await open(page, 'doc=frozen&print=on')
    await settled(page)
    await page.locator('[data-ov-print-button]').click()
    await expect.poll(async () => (await prints(page)).frames).toBe(1)
    const images = page.frameLocator('iframe[data-officeview-print-frame]').locator('img[data-page]')
    await expect(images).toHaveCount(1)
    const natural = await images.first().evaluate((i) => ({ w: (i as HTMLImageElement).naturalWidth, h: (i as HTMLImageElement).naturalHeight }))
    expect(natural.w).toBeGreaterThan(0)
    expect(natural.h).toBeGreaterThan(0)
    await expect(page.locator('[data-ov-print-notice]')).toHaveCount(0)
    expect((await prints(page)).windowPrints).toBe(0)
  })
})
