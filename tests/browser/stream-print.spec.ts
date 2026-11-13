/**
 * Stage 3b Task 8: printing a streamed document in a real browser against
 * server/test/app.ts (`?print=on|off|cap2|session3` sets the session's print
 * policy: true, false, {maxPages: 2}, {maxPages: 5, maxSessionPages: 3}).
 * Documents: pypptx-test.pptx (1 unit) for the uncapped job, and
 * lo-tdf104788.pptx (6 units) for the caps, which need more than 3 pages.
 *
 * An init script stubs `print` on every [data-officeview-print-frame] iframe's
 * window as it is added (a MutationObserver), recording calls on
 * window.__ovPrints; window.print is stubbed too and must never be called.
 * The stubbed print fires no afterprint, so the frame stays for inspection.
 */
import { expect, test, type Page } from '@playwright/test'

const P = Number(process.env.OV_STREAM_PORT ?? 4181)
const open = (page: Page, q: string) => page.goto(`/src/debug/stream-viewer.html?port=${P}&${q}`)
const painted = (page: Page) => page.locator('canvas[data-painted="1"]')
const printButton = (page: Page) => page.locator('[data-ov-print-button]')
const dialog = (page: Page) => page.getByRole('dialog', { name: 'Print pages' })
const notice = (page: Page) => page.locator('[data-ov-print-notice]')
const frameImages = (page: Page) => page.frameLocator('iframe[data-officeview-print-frame]').locator('img[data-page]')

type Rec = { frames: number; windowPrints: number }
const prints = (page: Page) => page.evaluate(() => (window as unknown as { __ovPrints: Rec }).__ovPrints)

test.beforeEach(async ({ page }) => {
  await page.addInitScript(() => {
    const rec = { frames: 0, windowPrints: 0 }
    ;(window as unknown as { __ovPrints: Rec }).__ovPrints = rec
    window.print = () => { rec.windowPrints++ }
    const stub = (f: HTMLIFrameElement) => {
      const w = f.contentWindow as (Window & { __ovStubbed?: true }) | null
      if (!w || w.__ovStubbed) return
      w.__ovStubbed = true
      w.print = () => { rec.frames++ }
    }
    new MutationObserver((records) => {
      for (const r of records) for (const n of r.addedNodes) {
        if (n instanceof HTMLIFrameElement && n.hasAttribute('data-officeview-print-frame')) stub(n)
      }
    }).observe(document, { childList: true, subtree: true })
  })
})

/** The manifest's unit sizes, as the view lays them out (each canvas's inline CSS width/height). */
async function unitSizes(page: Page): Promise<Array<{ w: number; h: number }>> {
  return page.locator('[data-unit-wrap] canvas').evaluateAll((cs) => cs.map((c) => {
    const s = (c as HTMLCanvasElement).style, [aw, ah] = s.aspectRatio.split('/').map((v) => Number(v.trim()))
    return { w: parseFloat(s.width), h: Math.round((parseFloat(s.width) * ah!) / aw!) }
  }))
}

/** Each printed image is the scale-2 print raster: (u − 1) × 2 < natural ≤ u × 2 on both axes. */
async function expectScale2(page: Page, units: Array<{ w: number; h: number }>, from: number) {
  const nat = await frameImages(page).evaluateAll((is) => is.map((i) => ({ w: (i as HTMLImageElement).naturalWidth, h: (i as HTMLImageElement).naturalHeight })))
  nat.forEach((n, k) => {
    const u = units[from + k]!
    expect(n.w).toBeGreaterThan((u.w - 1) * 2); expect(n.w).toBeLessThanOrEqual(u.w * 2)
    expect(n.h).toBeGreaterThan((u.h - 1) * 2); expect(n.h).toBeLessThanOrEqual(u.h * 2)
  })
}

test('uncapped: Print prints every unit once, from scale-2 print rasters', async ({ page }) => {
  await open(page, 'doc=pypptx-test.pptx&print=on')
  await expect(painted(page)).toHaveCount(1)
  const units = await unitSizes(page)
  expect(units).toHaveLength(1)
  await printButton(page).click()
  await expect.poll(async () => (await prints(page)).frames).toBe(1)
  await expect(dialog(page)).toHaveCount(0)
  await expect(frameImages(page)).toHaveCount(units.length)
  await expectScale2(page, units, 0)
  await expect(notice(page)).toHaveCount(0)
  expect((await prints(page)).windowPrints).toBe(0)
  // The frame is hosted outside the remote root, so print media (which hides
  // that root) leaves the frame's pages laid out.
  expect(await page.locator('iframe[data-officeview-print-frame]').evaluate((f) => f.parentElement === document.body && !f.closest('[data-ov-theme-root]'))).toBe(true)
  await page.emulateMedia({ media: 'print' })
  const laidOut = await frameImages(page).evaluateAll((is) => is.map((i) => getComputedStyle(i).display !== 'none' && i.getBoundingClientRect().width > 0))
  expect(laidOut).toEqual(units.map(() => true))
})

test('cap2: the dialog refuses 1–3 with the limit notice and prints 1–2', async ({ page }) => {
  await open(page, 'doc=lo-tdf104788.pptx&print=cap2')
  await expect(page.locator('[data-unit-wrap]')).toHaveCount(6)
  await expect(painted(page).first()).toBeVisible()
  const units = await unitSizes(page)
  await printButton(page).click()
  const d = dialog(page)
  await expect(d).toBeVisible()
  await expect(d.getByLabel('From')).toHaveValue('1')
  await expect(d.getByLabel('To')).toHaveValue('2')
  await expect(d.getByText('Up to 2 pages per print.')).toBeVisible()
  await d.getByLabel('To').fill('3')
  await d.getByRole('button', { name: 'Print' }).click()
  await expect(notice(page)).toHaveText('Too many pages to print at once (limit: 2 pages).')
  expect((await prints(page)).frames).toBe(0)

  await printButton(page).click()
  await expect(d).toBeVisible()
  await d.getByLabel('From').fill('1')
  await d.getByLabel('To').fill('2')
  await d.getByRole('button', { name: 'Print' }).click()
  await expect.poll(async () => (await prints(page)).frames).toBe(1)
  await expect(frameImages(page)).toHaveCount(2)
  await expectScale2(page, units, 0)
  await expect(notice(page)).toHaveCount(0)
  expect((await prints(page)).windowPrints).toBe(0)
})

test('session3: a second 2-page job over the session total shows the session-limit notice', async ({ page }) => {
  await open(page, 'doc=lo-tdf104788.pptx&print=session3')
  await expect(page.locator('[data-unit-wrap]')).toHaveCount(6)
  const d = dialog(page)
  const job = async () => {
    await printButton(page).click()
    await expect(d).toBeVisible()
    await d.getByLabel('From').fill('1')
    await d.getByLabel('To').fill('2')
    await d.getByRole('button', { name: 'Print' }).click()
  }
  await job()
  await expect(d).toHaveCount(0)
  await expect.poll(async () => (await prints(page)).frames).toBe(1)
  await expect(frameImages(page)).toHaveCount(2)
  await expect(printButton(page)).toBeEnabled()
  await job()
  await expect(d).toHaveCount(0)
  await expect(notice(page)).toHaveText('Print limit for this session reached (3 pages).')
  expect((await prints(page)).frames).toBe(1)
  expect((await prints(page)).windowPrints).toBe(0)
})

test('print off: no Print button, and Ctrl/Cmd+P opens nothing', async ({ page }) => {
  await open(page, 'doc=pypptx-test.pptx&print=off')
  await expect(painted(page)).toHaveCount(1)
  await expect(printButton(page)).toHaveCount(0)
  await page.locator('[role="document"]').focus()
  await page.keyboard.press('Control+p')
  await page.keyboard.press('Meta+p')
  await page.waitForTimeout(500)
  await expect(dialog(page)).toHaveCount(0)
  await expect(page.locator('iframe[data-officeview-print-frame]')).toHaveCount(0)
  expect(await prints(page)).toEqual({ frames: 0, windowPrints: 0 })
})

test('Ctrl+P with print allowed starts the same flow (cap2 opens the dialog)', async ({ page }) => {
  await open(page, 'doc=lo-tdf104788.pptx&print=cap2')
  await expect(page.locator('[data-unit-wrap]')).toHaveCount(6)
  await page.locator('[role="document"]').focus()
  await page.keyboard.press('Control+p')
  await expect(dialog(page)).toBeVisible()
  expect((await prints(page)).windowPrints).toBe(0)
})

test('in print media the remote root is display:none (screen canvases never print)', async ({ page }) => {
  await open(page, 'doc=pypptx-test.pptx&print=on')
  await expect(painted(page)).toHaveCount(1)
  const root = page.locator('[data-ov-theme-root][data-ov-print="false"]')
  await expect(root).toHaveCount(1)
  expect(await root.evaluate((e) => getComputedStyle(e).display)).not.toBe('none')
  await page.emulateMedia({ media: 'print' })
  expect(await root.evaluate((e) => getComputedStyle(e).display)).toBe('none')
})
