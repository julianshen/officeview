/**
 * S10.3 capability and failure scenarios for the real interactive viewer
 * (spec §§4.2, 10.3, 12.4; LIFE-03, AC-CAP-03, AC-BACKEND-01 fallback leg).
 *
 * Main-thread capabilities are removed with init scripts. Init scripts never
 * reach a dedicated worker, so worker-side denials are named in the Worker
 * `name` and applied by src/debug/workerCaps.ts before the worker entry runs.
 * Canvas context loss cannot be triggered from page script for a 2D canvas, so
 * that scenario is an emulation: the backing store is cleared and the
 * contextlost/contextrestored events are dispatched as the browser would.
 */
import { expect, test, type Page } from '@playwright/test'
import { activeCell, ensureFixtures, goTo, openViewer, status } from './viewer-e2e'

test.use({ viewport: { width: 1440, height: 900 }, deviceScaleFactor: 2 })
test.beforeAll(async () => { await ensureFixtures() })

const e2e = (page: Page) => page.evaluate(() => ({ created: window.__viewerE2E!.workersCreated, terminated: window.__viewerE2E!.workersTerminated, open: window.__viewerE2E!.counts().open ?? 0 }))
const alert = (page: Page) => page.getByTestId('e2e-error')
const cornerReachable = async (page: Page) => {
  await goTo(page, 'XFD1048576')
  await expect.poll(async () => (await activeCell(page))?.label ?? '', { timeout: 30_000 }).toMatch(/^XFD1048576, value 17179869183/)
}

test.describe('S10.3 capability fallbacks', () => {
  test('Worker unavailable: small file stays resident and correct; a large file keeps compatibility mode with its notice', async ({ page }) => {
    await page.addInitScript(() => { delete (window as { Worker?: unknown }).Worker })
    await openViewer(page, 'f1-small.xlsx', { worker: true })
    expect(await page.evaluate(() => typeof Worker)).toBe('undefined')
    await cornerReachable(page)
    await expect(status(page)).not.toContainText('compatibility mode')
    expect(await e2e(page)).toMatchObject({ created: 0 })

    // Reproduced gap: the host factory threw "Worker is not defined" and the
    // whole document failed. Now it is the same outcome as a dead worker.
    await openViewer(page, 'f1-paged.xlsx', { worker: true })
    await expect(status(page)).toContainText('Large file — opened in compatibility mode')
    await expect(alert(page)).toHaveCount(0)
    await cornerReachable(page)
    expect(await e2e(page)).toMatchObject({ created: 0 })
  })

  test('OffscreenCanvas absent on the main thread and in the worker: worker-backed viewing is unchanged', async ({ page }) => {
    await page.addInitScript(() => { delete (window as { OffscreenCanvas?: unknown }).OffscreenCanvas })
    await openViewer(page, 'f1-paged.xlsx', { worker: true, workerCaps: ['no-offscreen'] })
    expect(await page.evaluate(() => typeof OffscreenCanvas)).toBe('undefined')
    expect(await e2e(page)).toMatchObject({ created: 1, open: 1 })
    await cornerReachable(page)
    await expect(status(page)).not.toContainText('compatibility mode')
  })

  test('fonts unavailable (no FontFace/FontFaceSet on either side): text falls back and the grid stays owned', async ({ page }) => {
    await page.addInitScript(() => {
      delete (window as { FontFace?: unknown }).FontFace
      Object.defineProperty(document, 'fonts', { configurable: true, value: undefined })
    })
    await openViewer(page, 'f1-paged.xlsx', { worker: true, workerCaps: ['no-fonts'] })
    await cornerReachable(page)
    expect(await page.getByRole('gridcell').count()).toBeGreaterThan(20)
    expect((await page.evaluate(() => window.__viewerE2E!.lastFrame!)).basePaints).toBeGreaterThan(0)
  })

  test('small paged file with no persistent storage and no quota opens worker-backed from bounded memory pages', async ({ page }) => {
    await openViewer(page, 'f1-paged.xlsx', { worker: true, workerCaps: ['no-storage', 'quota-zero'] })
    expect(await e2e(page)).toMatchObject({ created: 1, open: 1 })
    await cornerReachable(page)
  })

  test('large paged file: no OPFS/IndexedDB, or zero quota, is a visible typed refusal — never a grid claiming the sheet', async ({ page }) => {
    await openViewer(page, 'f2-storage.xlsx', { worker: true, workerCaps: ['no-storage'], ready: false })
    await expect(alert(page)).toContainText('needs browser storage (OPFS or IndexedDB)', { timeout: 60_000 })
    await expect(page.getByRole('grid')).toHaveCount(0)
    await expect.poll(() => e2e(page)).toMatchObject({ created: 1, terminated: 1 })

    await openViewer(page, 'f2-storage.xlsx', { worker: true, workerCaps: ['quota-zero'], ready: false })
    await expect(alert(page)).toContainText(/browser storage has 0 bytes free; about \d+ are needed/, { timeout: 60_000 })
    await expect(page.getByRole('grid')).toHaveCount(0)
  })

  test('quota denied while the worker writes pages: the failure is stated and the body is incomplete, not an empty complete sheet', async ({ page }) => {
    await openViewer(page, 'f2-storage.xlsx', { worker: true, workerCaps: ['idb-quota'], ready: false })
    // Reproduced gap: "Visible body 0 / complete body 0" with cells "Loading…" forever.
    const line = page.locator('[data-officeview-ui=viewer-status]')
    await expect(line).toContainText('This sheet could not be fully loaded: browser storage quota exceeded while writing a source page', { timeout: 60_000 })
    await expect(line).toContainText('(incomplete)')
    await expect(line).not.toContainText('complete body')
  })

  test('without a viewer worker a large file is labelled compatibility mode', async ({ page }) => {
    await openViewer(page, 'f2-storage.xlsx')
    await expect(status(page)).toContainText('Large file — opened in compatibility mode')
    expect(await e2e(page)).toMatchObject({ created: 0 })
  })

  test('canvas context loss (emulated): restore repaints from records, not from lost tiles', async ({ page }) => {
    await openViewer(page, 'f2-1k.xlsx')
    const base = page.getByTestId('officeview-sheet-base')
    await expect.poll(() => page.evaluate(() => window.__viewerE2E!.lastFrame?.basePaints ?? 0)).toBeGreaterThan(0)
    const before = await page.evaluate(() => window.__viewerE2E!.lastFrame!.basePaints)
    const alphaAt = () => base.evaluate((canvas: HTMLCanvasElement) => canvas.getContext('2d')!.getImageData(20, 20, 1, 1).data[3])
    expect(await alphaAt()).toBe(255)
    // A lost context and its tiles come back blank: clear both, then restore.
    await page.evaluate(() => {
      const canvases = [...document.querySelectorAll('canvas')] as HTMLCanvasElement[]
      for (const canvas of canvases) canvas.getContext('2d')?.clearRect(0, 0, canvas.width, canvas.height)
      const surface = document.querySelector('[data-testid=officeview-sheet-base]')!
      surface.dispatchEvent(new Event('contextlost'))
      surface.dispatchEvent(new Event('contextrestored'))
    })
    await expect.poll(alphaAt).toBe(255)
    expect(await page.evaluate(() => window.__viewerE2E!.lastFrame!.basePaints)).toBeGreaterThan(before)
    // The viewer still works after the restore.
    await goTo(page, 'C5')
    await expect.poll(async () => (await activeCell(page))?.label.split(',')[0] ?? '').toBe('C5')
  })
})
